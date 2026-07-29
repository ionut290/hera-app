package it.vargacantieri.hera.whatsapp;

import android.content.ActivityNotFoundException;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

@CapacitorPlugin(name = "HeraWhatsApp")
public class HeraWhatsAppPlugin extends Plugin {
    private static final String WHATSAPP = "com.whatsapp";
    private static final String WHATSAPP_BUSINESS = "com.whatsapp.w4b";

    @PluginMethod
    public void open(PluginCall call) {
        String url = call.getString("url", "");
        String packageName = resolveInstalledPackage();
        if (packageName == null) {
            call.reject("WhatsApp non è installato sul dispositivo.");
            return;
        }

        Uri uri = normalizeUri(url);
        Intent intent = new Intent(Intent.ACTION_VIEW, uri);
        intent.setPackage(packageName);
        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);

        try {
            getActivity().startActivity(intent);
            JSObject result = new JSObject();
            result.put("opened", true);
            result.put("packageName", packageName);
            call.resolve(result);
        } catch (ActivityNotFoundException error) {
            call.reject("Impossibile aprire WhatsApp installato.", error);
        } catch (Exception error) {
            call.reject("Errore durante l'apertura di WhatsApp.", error);
        }
    }

    private String resolveInstalledPackage() {
        PackageManager packageManager = getContext().getPackageManager();
        if (isInstalled(packageManager, WHATSAPP)) return WHATSAPP;
        if (isInstalled(packageManager, WHATSAPP_BUSINESS)) return WHATSAPP_BUSINESS;
        return null;
    }

    private boolean isInstalled(PackageManager packageManager, String packageName) {
        try {
            packageManager.getPackageInfo(packageName, 0);
            return true;
        } catch (PackageManager.NameNotFoundException error) {
            return false;
        }
    }

    private Uri normalizeUri(String url) {
        if (url == null || url.trim().isEmpty()) {
            return Uri.parse("https://wa.me/");
        }
        String trimmed = url.trim();
        if (trimmed.startsWith("whatsapp://")
                || trimmed.startsWith("https://wa.me/")
                || trimmed.startsWith("https://api.whatsapp.com/")) {
            return Uri.parse(trimmed);
        }
        return Uri.parse("https://wa.me/?text=" + Uri.encode(trimmed));
    }
}
