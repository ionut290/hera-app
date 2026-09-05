package it.vargacantieri.hera;

import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageInfo;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.view.Window;
import android.webkit.WebView;

import androidx.activity.EdgeToEdge;
import androidx.core.view.WindowCompat;
import androidx.core.view.WindowInsetsCompat;
import androidx.core.view.WindowInsetsControllerCompat;

import com.getcapacitor.BridgeActivity;

import it.vargacantieri.hera.camera.HeraContinuousCameraPlugin;
import it.vargacantieri.hera.geofence.HeraGeofencePlugin;
import it.vargacantieri.hera.biometric.HeraBiometricPlugin;
import it.vargacantieri.hera.biometric.HeraCredentialVaultPlugin;
import it.vargacantieri.hera.update.HeraAppUpdatePlugin;
import it.vargacantieri.hera.whatsapp.HeraWhazzupPhotosPlugin;
import it.vargacantieri.hera.whatsapp.HeraWhatsAppPlugin;

public class MainActivity extends BridgeActivity {
    private static final String CACHE_PREFS_NAME = "hera_native_cache";
    private static final String CACHE_VERSION_CODE_KEY = "clearedForVersionCode";
    private static final String TEMP_LOGIN_PARAM = "temp-login";

    @Override
    public void onCreate(Bundle savedInstanceState) {
        // Impedisce ad Android/AppCompat di creare la barra con il titolo
        // "Varga Cantieri" prima che venga inizializzata la WebView.
        supportRequestWindowFeature(Window.FEATURE_NO_TITLE);
        registerPlugin(HeraGeofencePlugin.class);
        registerPlugin(HeraBiometricPlugin.class);
        registerPlugin(HeraCredentialVaultPlugin.class);
        registerPlugin(HeraAppUpdatePlugin.class);
        registerPlugin(HeraWhazzupPhotosPlugin.class);
        registerPlugin(HeraWhatsAppPlugin.class);
        registerPlugin(HeraContinuousCameraPlugin.class);
        EdgeToEdge.enable(this);
        super.onCreate(savedInstanceState);
        if (getSupportActionBar() != null) {
            getSupportActionBar().hide();
        }
        hideAndroidStatusBar();
        clearWebViewCacheAfterAppUpdate();
        applyTemporaryLoginDeepLink(getIntent());
    }

    @Override
    public void onWindowFocusChanged(boolean hasFocus) {
        super.onWindowFocusChanged(hasFocus);
        if (hasFocus) {
            hideAndroidStatusBar();
        }
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        setIntent(intent);
        applyTemporaryLoginDeepLink(intent);
    }

    private void hideAndroidStatusBar() {
        try {
            WindowCompat.setDecorFitsSystemWindows(getWindow(), false);
            WindowInsetsControllerCompat controller = WindowCompat.getInsetsController(getWindow(), getWindow().getDecorView());
            controller.hide(WindowInsetsCompat.Type.statusBars());
            controller.setSystemBarsBehavior(
                WindowInsetsControllerCompat.BEHAVIOR_SHOW_TRANSIENT_BARS_BY_SWIPE
            );
        } catch (Exception ignored) {
            // La barra di stato non deve mai interferire con l'avvio dell'app.
        }
    }

    private void applyTemporaryLoginDeepLink(Intent intent) {
        try {
            Uri data = intent == null ? null : intent.getData();
            if (data == null) return;
            if (!"vargacantieri".equalsIgnoreCase(data.getScheme())) return;
            if (!"login".equalsIgnoreCase(data.getHost())) return;

            String payload = data.getQueryParameter(TEMP_LOGIN_PARAM);
            if (payload == null || payload.trim().isEmpty()) return;

            WebView webView = getBridge() == null ? null : getBridge().getWebView();
            if (webView == null) return;

            String safePayload = payload
                .replace("\\", "\\\\")
                .replace("'", "\\'")
                .replace("\n", "")
                .replace("\r", "");
            String javascript = "window.location.hash='temp-login=" + safePayload + "';";
            webView.postDelayed(() -> webView.evaluateJavascript(javascript, null), 350L);
        } catch (Exception ignored) {
            // Il deep link non deve mai impedire l'avvio normale dell'app.
        }
    }

    private void clearWebViewCacheAfterAppUpdate() {
        try {
            PackageInfo packageInfo = getPackageManager().getPackageInfo(getPackageName(), 0);
            long currentVersionCode = Build.VERSION.SDK_INT >= Build.VERSION_CODES.P
                ? packageInfo.getLongVersionCode()
                : packageInfo.versionCode;
            SharedPreferences preferences = getSharedPreferences(CACHE_PREFS_NAME, Context.MODE_PRIVATE);
            long clearedForVersionCode = preferences.getLong(CACHE_VERSION_CODE_KEY, -1L);
            if (clearedForVersionCode == currentVersionCode) return;

            WebView webView = getBridge() == null ? null : getBridge().getWebView();
            if (webView == null) return;

            webView.clearCache(true);
            preferences.edit().putLong(CACHE_VERSION_CODE_KEY, currentVersionCode).apply();
        } catch (Exception ignored) {
            // Non bloccare mai l'avvio: se la pulizia fallisce, verrà ritentata al prossimo avvio.
        }
    }
}
