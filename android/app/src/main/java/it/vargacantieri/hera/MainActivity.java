package it.vargacantieri.hera;

import android.content.Context;
import android.content.SharedPreferences;
import android.content.pm.PackageInfo;
import android.os.Build;
import android.os.Bundle;
import android.webkit.WebView;

import com.getcapacitor.BridgeActivity;

import it.vargacantieri.hera.camera.HeraContinuousCameraPlugin;
import it.vargacantieri.hera.geofence.HeraGeofencePlugin;
import it.vargacantieri.hera.biometric.HeraBiometricPlugin;
import it.vargacantieri.hera.whatsapp.HeraWhazzupPhotosPlugin;
import it.vargacantieri.hera.whatsapp.HeraWhatsAppPlugin;

public class MainActivity extends BridgeActivity {
    private static final String CACHE_PREFS_NAME = "hera_native_cache";
    private static final String CACHE_VERSION_CODE_KEY = "clearedForVersionCode";

    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(HeraGeofencePlugin.class);
        registerPlugin(HeraBiometricPlugin.class);
        registerPlugin(HeraWhazzupPhotosPlugin.class);
        registerPlugin(HeraWhatsAppPlugin.class);
        registerPlugin(HeraContinuousCameraPlugin.class);
        super.onCreate(savedInstanceState);
        clearWebViewCacheAfterAppUpdate();
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
