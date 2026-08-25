package it.vargacantieri.hera.biometric;

import android.content.Context;
import android.content.SharedPreferences;
import android.security.keystore.KeyGenParameterSpec;
import android.security.keystore.KeyProperties;
import android.util.Base64;

import androidx.annotation.NonNull;
import androidx.biometric.BiometricManager;
import androidx.biometric.BiometricPrompt;
import androidx.core.content.ContextCompat;

import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.annotation.CapacitorPlugin;

import org.json.JSONArray;
import org.json.JSONObject;

import java.nio.charset.StandardCharsets;
import java.security.KeyStore;
import java.util.concurrent.Executor;

import javax.crypto.Cipher;
import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import javax.crypto.spec.GCMParameterSpec;

@CapacitorPlugin(name = "HeraCredentialVault")
public class HeraCredentialVaultPlugin extends Plugin {
    private static final String KEY_ALIAS = "it.vargacantieri.hera.credential_vault";
    private static final String PREFS = "hera_credential_vault";
    private static final String CIPHER = "AES/GCM/NoPadding";
    private static final String PAYLOAD = "payload";
    private static final String IV = "iv";

    @com.getcapacitor.PluginMethod
    public void storeCredential(PluginCall call) {
        String email = normalizeEmail(call.getString("email", ""));
        String password = call.getString("password", "");
        if (email.isEmpty() || password.isEmpty()) {
            call.reject("Email o password mancanti.", "invalid_credential");
            return;
        }
        try {
            JSONArray items = readItems();
            JSONArray next = new JSONArray();
            JSONObject fresh = new JSONObject();
            fresh.put("email", email);
            fresh.put("password", password);
            fresh.put("savedAt", System.currentTimeMillis());
            next.put(fresh);
            for (int i = 0; i < items.length() && next.length() < 10; i++) {
                JSONObject item = items.optJSONObject(i);
                if (item == null) continue;
                if (email.equals(normalizeEmail(item.optString("email")))) continue;
                next.put(item);
            }
            writeItems(next);
            call.resolve();
        } catch (Exception error) {
            call.reject("Impossibile salvare la credenziale sul dispositivo.", "vault_store_failed", error);
        }
    }

    @com.getcapacitor.PluginMethod
    public void listCredentials(PluginCall call) {
        int availability = BiometricManager.from(getContext()).canAuthenticate(BiometricManager.Authenticators.BIOMETRIC_STRONG);
        if (availability != BiometricManager.BIOMETRIC_SUCCESS) {
            call.reject("Configura impronta o riconoscimento biometrico sul telefono.", "biometric_unavailable");
            return;
        }
        Executor executor = ContextCompat.getMainExecutor(getContext());
        BiometricPrompt prompt = new BiometricPrompt(getActivity(), executor, new BiometricPrompt.AuthenticationCallback() {
            @Override
            public void onAuthenticationError(int code, @NonNull CharSequence message) {
                call.reject(message.toString(), code == BiometricPrompt.ERROR_NEGATIVE_BUTTON || code == BiometricPrompt.ERROR_USER_CANCELED ? "cancelled" : "authentication_failed");
            }

            @Override
            public void onAuthenticationSucceeded(@NonNull BiometricPrompt.AuthenticationResult result) {
                try {
                    JSONArray items = readItems();
                    JSArray array = new JSArray();
                    for (int i = 0; i < items.length(); i++) {
                        JSONObject item = items.optJSONObject(i);
                        if (item == null) continue;
                        JSObject out = new JSObject();
                        out.put("email", item.optString("email"));
                        out.put("password", item.optString("password"));
                        out.put("savedAt", item.optLong("savedAt", 0L));
                        array.put(out);
                    }
                    JSObject response = new JSObject();
                    response.put("accounts", array);
                    call.resolve(response);
                } catch (Exception error) {
                    call.reject("Impossibile leggere le credenziali salvate.", "vault_read_failed", error);
                }
            }
        });
        BiometricPrompt.PromptInfo info = new BiometricPrompt.PromptInfo.Builder()
            .setTitle(call.getString("title", "Password salvate"))
            .setSubtitle(call.getString("subtitle", "Conferma la tua identità"))
            .setNegativeButtonText("ANNULLA")
            .setAllowedAuthenticators(BiometricManager.Authenticators.BIOMETRIC_STRONG)
            .build();
        prompt.authenticate(info);
    }

    @com.getcapacitor.PluginMethod
    public void deleteCredential(PluginCall call) {
        String email = normalizeEmail(call.getString("email", ""));
        if (email.isEmpty()) {
            call.reject("Email mancante.", "invalid_email");
            return;
        }
        try {
            JSONArray items = readItems();
            JSONArray next = new JSONArray();
            for (int i = 0; i < items.length(); i++) {
                JSONObject item = items.optJSONObject(i);
                if (item == null) continue;
                if (email.equals(normalizeEmail(item.optString("email")))) continue;
                next.put(item);
            }
            writeItems(next);
            call.resolve();
        } catch (Exception error) {
            call.reject("Impossibile eliminare la credenziale.", "vault_delete_failed", error);
        }
    }

    private JSONArray readItems() throws Exception {
        SharedPreferences prefs = prefs();
        if (!prefs.contains(PAYLOAD) || !prefs.contains(IV)) return new JSONArray();
        SecretKey key = getOrCreateKey();
        Cipher cipher = Cipher.getInstance(CIPHER);
        byte[] iv = Base64.decode(prefs.getString(IV, ""), Base64.NO_WRAP);
        cipher.init(Cipher.DECRYPT_MODE, key, new GCMParameterSpec(128, iv));
        byte[] clear = cipher.doFinal(Base64.decode(prefs.getString(PAYLOAD, ""), Base64.NO_WRAP));
        String json = new String(clear, StandardCharsets.UTF_8);
        return new JSONArray(json);
    }

    private void writeItems(JSONArray items) throws Exception {
        SecretKey key = getOrCreateKey();
        Cipher cipher = Cipher.getInstance(CIPHER);
        cipher.init(Cipher.ENCRYPT_MODE, key);
        byte[] encrypted = cipher.doFinal(items.toString().getBytes(StandardCharsets.UTF_8));
        prefs().edit()
            .putString(PAYLOAD, Base64.encodeToString(encrypted, Base64.NO_WRAP))
            .putString(IV, Base64.encodeToString(cipher.getIV(), Base64.NO_WRAP))
            .apply();
    }

    private SecretKey getOrCreateKey() throws Exception {
        KeyStore store = KeyStore.getInstance("AndroidKeyStore");
        store.load(null);
        if (store.containsAlias(KEY_ALIAS)) return (SecretKey) store.getKey(KEY_ALIAS, null);
        KeyGenerator generator = KeyGenerator.getInstance(KeyProperties.KEY_ALGORITHM_AES, "AndroidKeyStore");
        KeyGenParameterSpec spec = new KeyGenParameterSpec.Builder(
            KEY_ALIAS,
            KeyProperties.PURPOSE_ENCRYPT | KeyProperties.PURPOSE_DECRYPT
        ).setBlockModes(KeyProperties.BLOCK_MODE_GCM)
         .setEncryptionPaddings(KeyProperties.ENCRYPTION_PADDING_NONE)
         .build();
        generator.init(spec);
        return generator.generateKey();
    }

    private SharedPreferences prefs() {
        return getContext().getSharedPreferences(PREFS, Context.MODE_PRIVATE);
    }

    private String normalizeEmail(String value) {
        return value == null ? "" : value.trim().toLowerCase(java.util.Locale.ROOT);
    }
}
