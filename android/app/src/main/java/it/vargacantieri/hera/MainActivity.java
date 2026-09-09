package it.vargacantieri.hera;

import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageInfo;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.view.View;
import android.view.Window;
import android.view.WindowManager;
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
        // Blocca titolo e barra di stato prima che BridgeActivity crei la WebView.
        supportRequestWindowFeature(Window.FEATURE_NO_TITLE);
        getWindow().addFlags(WindowManager.LayoutParams.FLAG_FULLSCREEN);
        getWindow().clearFlags(WindowManager.LayoutParams.FLAG_FORCE_NOT_FULLSCREEN);
        registerPlugin(HeraGeofencePlugin.class);
        registerPlugin(HeraBiometricPlugin.class);
        registerPlugin(HeraCredentialVaultPlugin.class);
        registerPlugin(HeraAppUpdatePlugin.class);
        registerPlugin(HeraWhazzupPhotosPlugin.class);
        registerPlugin(HeraWhatsAppPlugin.class);
        registerPlugin(HeraContinuousCameraPlugin.class);
        EdgeToEdge.enable(this);
        super.onCreate(savedInstanceState);
        enforceFullscreenWithoutActionBar();
        clearWebViewCacheAfterAppUpdate();
        applyTemporaryLoginDeepLink(getIntent());
    }

    @Override
    public void onResume() {
        super.onResume();
        enforceFullscreenWithoutActionBar();
    }

    @Override
    protected void onPostResume() {
        super.onPostResume();
        enforceFullscreenWithoutActionBar();
    }

    @Override
    public void onWindowFocusChanged(boolean hasFocus) {
        super.onWindowFocusChanged(hasFocus);
        if (hasFocus) {
            enforceFullscreenWithoutActionBar();
        }
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        setIntent(intent);
        applyTemporaryLoginDeepLink(intent);
    }

    @SuppressWarnings("deprecation")
    private void enforceFullscreenWithoutActionBar() {
        try {
            if (getSupportActionBar() != null) {
                getSupportActionBar().hide();
            }

            Window window = getWindow();
            window.addFlags(WindowManager.LayoutParams.FLAG_FULLSCREEN);
            window.clearFlags(WindowManager.LayoutParams.FLAG_FORCE_NOT_FULLSCREEN);
            WindowCompat.setDecorFitsSystemWindows(window, false);

            View decorView = window.getDecorView();
            decorView.setSystemUiVisibility(
                View.SYSTEM_UI_FLAG_FULLSCREEN
                    | View.SYSTEM_UI_FLAG_IMMERSIVE_STICKY
                    | View.SYSTEM_UI_FLAG_LAYOUT_FULLSCREEN
                    | View.SYSTEM_UI_FLAG_LAYOUT_STABLE
            );

            WindowInsetsControllerCompat controller = WindowCompat.getInsetsController(window, decorView);
            controller.hide(WindowInsetsCompat.Type.statusBars());
            controller.setSystemBarsBehavior(
                WindowInsetsControllerCompat.BEHAVIOR_SHOW_TRANSIENT_BARS_BY_SWIPE
            );
        } catch (Exception ignored) {
            // Le barre native non devono mai interferire con i comandi della WebView.
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
