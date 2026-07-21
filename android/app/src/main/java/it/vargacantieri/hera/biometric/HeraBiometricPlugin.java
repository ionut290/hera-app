package it.vargacantieri.hera.biometric;

import android.content.Context;
import android.content.SharedPreferences;
import android.os.Build;
import android.security.keystore.KeyGenParameterSpec;
import android.security.keystore.KeyProperties;
import android.util.Base64;

import androidx.annotation.NonNull;
import androidx.biometric.BiometricManager;
import androidx.biometric.BiometricPrompt;
import androidx.core.content.ContextCompat;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.nio.charset.StandardCharsets;
import java.security.KeyStore;
import java.util.concurrent.Executor;

import javax.crypto.Cipher;
import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import javax.crypto.spec.GCMParameterSpec;

/** Local Capacitor 7 plugin. The authorization marker is encrypted by an Android Keystore key. */
@CapacitorPlugin(name = "HeraBiometric")
public class HeraBiometricPlugin extends Plugin {
    private static final String PLUGIN_VERSION = "1.0.0";
    private static final String KEY_ALIAS = "it.vargacantieri.hera.biometric_access";
    private static final String PREFS = "hera_biometric_secure";
    private static final String CIPHER = "AES/GCM/NoPadding";
    private static final byte[] MARKER = "HERA_BIOMETRIC_AUTHORIZED".getBytes(StandardCharsets.UTF_8);

    @com.getcapacitor.PluginMethod
    public void status(PluginCall call) {
        int result = BiometricManager.from(getContext()).canAuthenticate(BiometricManager.Authenticators.BIOMETRIC_STRONG);
        JSObject data = new JSObject();
        data.put("available", result == BiometricManager.BIOMETRIC_SUCCESS);
        data.put("enabled", prefs().contains("payload") && prefs().contains("iv"));
        data.put("reason", reason(result));
        data.put("version", PLUGIN_VERSION);
        call.resolve(data);
    }

    @com.getcapacitor.PluginMethod
    public void enable(PluginCall call) {
        if (!ensureAvailable(call)) return;
        try {
            deleteKey();
            KeyGenerator generator = KeyGenerator.getInstance(KeyProperties.KEY_ALGORITHM_AES, "AndroidKeyStore");
            KeyGenParameterSpec.Builder spec = new KeyGenParameterSpec.Builder(KEY_ALIAS,
                    KeyProperties.PURPOSE_ENCRYPT | KeyProperties.PURPOSE_DECRYPT)
                    .setBlockModes(KeyProperties.BLOCK_MODE_GCM)
                    .setEncryptionPaddings(KeyProperties.ENCRYPTION_PADDING_NONE)
                    .setUserAuthenticationRequired(true)
                    .setInvalidatedByBiometricEnrollment(true);
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
                spec.setUserAuthenticationParameters(0, KeyProperties.AUTH_BIOMETRIC_STRONG);
            } else {
                spec.setUserAuthenticationValidityDurationSeconds(-1);
            }
            generator.init(spec.build());
            SecretKey key = generator.generateKey();
            Cipher cipher = Cipher.getInstance(CIPHER);
            cipher.init(Cipher.ENCRYPT_MODE, key);
            prompt(call, cipher, true);
        } catch (Exception error) {
            call.reject("Impossibile inizializzare Android Keystore.", "keystore_error", error);
        }
    }

    @com.getcapacitor.PluginMethod
    public void authenticate(PluginCall call) {
        if (!ensureAvailable(call)) return;
        SharedPreferences p = prefs();
        if (!p.contains("payload") || !p.contains("iv")) {
            call.reject("Accesso biometrico non attivo.", "not_enabled");
            return;
        }
        try {
            KeyStore store = KeyStore.getInstance("AndroidKeyStore");
            store.load(null);
            SecretKey key = (SecretKey) store.getKey(KEY_ALIAS, null);
            Cipher cipher = Cipher.getInstance(CIPHER);
            cipher.init(Cipher.DECRYPT_MODE, key,
                    new GCMParameterSpec(128, Base64.decode(p.getString("iv", ""), Base64.NO_WRAP)));
            prompt(call, cipher, false);
        } catch (Exception error) {
            disableInternal();
            call.reject("La configurazione biometrica è cambiata. Attivala nuovamente.", "key_invalidated", error);
        }
    }

    @com.getcapacitor.PluginMethod
    public void disable(PluginCall call) {
        disableInternal();
        call.resolve();
    }

    private void prompt(PluginCall call, Cipher cipher, boolean enabling) {
        Executor executor = ContextCompat.getMainExecutor(getContext());
        BiometricPrompt prompt = new BiometricPrompt(getActivity(), executor, new BiometricPrompt.AuthenticationCallback() {
            @Override public void onAuthenticationError(int code, @NonNull CharSequence message) {
                call.reject(message.toString(), code == BiometricPrompt.ERROR_NEGATIVE_BUTTON || code == BiometricPrompt.ERROR_USER_CANCELED ? "cancelled" : "authentication_failed");
            }
            @Override public void onAuthenticationSucceeded(@NonNull BiometricPrompt.AuthenticationResult result) {
                try {
                    Cipher authenticatedCipher = result.getCryptoObject().getCipher();
                    if (enabling) {
                        byte[] encrypted = authenticatedCipher.doFinal(MARKER);
                        prefs().edit().putString("payload", Base64.encodeToString(encrypted, Base64.NO_WRAP))
                                .putString("iv", Base64.encodeToString(authenticatedCipher.getIV(), Base64.NO_WRAP)).apply();
                    } else {
                        byte[] clear = authenticatedCipher.doFinal(Base64.decode(prefs().getString("payload", ""), Base64.NO_WRAP));
                        if (!java.util.Arrays.equals(clear, MARKER)) throw new SecurityException("Invalid marker");
                    }
                    call.resolve();
                } catch (Exception error) {
                    call.reject("Verifica biometrica non valida.", "authentication_failed", error);
                }
            }
        });
        BiometricPrompt.PromptInfo info = new BiometricPrompt.PromptInfo.Builder()
                .setTitle(call.getString("title", "Accedi a Varga Cantieri"))
                .setSubtitle(call.getString("subtitle", "Usa l’impronta digitale o il riconoscimento facciale"))
                .setNegativeButtonText("ANNULLA")
                .setAllowedAuthenticators(BiometricManager.Authenticators.BIOMETRIC_STRONG)
                .build();
        prompt.authenticate(info, new BiometricPrompt.CryptoObject(cipher));
    }

    private boolean ensureAvailable(PluginCall call) {
        int result = BiometricManager.from(getContext()).canAuthenticate(BiometricManager.Authenticators.BIOMETRIC_STRONG);
        if (result == BiometricManager.BIOMETRIC_SUCCESS) return true;
        call.reject(reason(result), result == BiometricManager.BIOMETRIC_ERROR_NONE_ENROLLED ? "none_enrolled" : "not_available");
        return false;
    }

    private String reason(int result) {
        return result == BiometricManager.BIOMETRIC_ERROR_NONE_ENROLLED ? "none_enrolled" :
                result == BiometricManager.BIOMETRIC_SUCCESS ? "" : "not_available";
    }
    private SharedPreferences prefs() { return getContext().getSharedPreferences(PREFS, Context.MODE_PRIVATE); }
    private void deleteKey() throws Exception { KeyStore s = KeyStore.getInstance("AndroidKeyStore"); s.load(null); if (s.containsAlias(KEY_ALIAS)) s.deleteEntry(KEY_ALIAS); }
    private void disableInternal() { prefs().edit().clear().apply(); try { deleteKey(); } catch (Exception ignored) {} }
}
