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

import java.util.Set;

@CapacitorPlugin(name = "HeraWhatsApp")
public class HeraWhatsAppPlugin extends Plugin {
    private static final String WHATSAPP = "com.whatsapp";
    private static final String WHATSAPP_BUSINESS = "com.whatsapp.w4b";

    @PluginMethod
    public void open(PluginCall call) {
        String rawUrl = call.getString("url", "");
        String packageName = resolveInstalledPackage();
        if (packageName == null) {
            call.reject("WhatsApp non è installato sul dispositivo.");
            return;
        }

        WhatsAppPayload payload = parsePayload(rawUrl);
        Intent intent = new Intent(Intent.ACTION_SEND);
        intent.setType("text/plain");
        intent.setPackage(packageName);
        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
        intent.putExtra(Intent.EXTRA_TEXT, payload.text);

        if (!payload.phone.isEmpty()) {
            intent.putExtra("jid", payload.phone + "@s.whatsapp.net");
        }

        try {
            getActivity().startActivity(intent);
            JSObject result = new JSObject();
            result.put("opened", true);
            result.put("packageName", packageName);
            result.put("phone", payload.phone);
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

    private WhatsAppPayload parsePayload(String rawUrl) {
        String value = rawUrl == null ? "" : rawUrl.trim();
        if (value.isEmpty()) return new WhatsAppPayload("", "");

        if (!(value.startsWith("whatsapp://")
                || value.startsWith("https://wa.me/")
                || value.startsWith("https://api.whatsapp.com/"))) {
            return new WhatsAppPayload(value, "");
        }

        try {
            Uri uri = Uri.parse(value);
            String text = firstNonEmpty(uri.getQueryParameter("text"), uri.getQueryParameter("message"));
            String phone = firstNonEmpty(uri.getQueryParameter("phone"), uri.getQueryParameter("send"));

            if (phone.isEmpty() && "wa.me".equalsIgnoreCase(uri.getHost())) {
                for (String segment : uri.getPathSegments()) {
                    String candidate = digitsOnly(segment);
                    if (!candidate.isEmpty()) {
                        phone = candidate;
                        break;
                    }
                }
            }

            if (phone.isEmpty() && "api.whatsapp.com".equalsIgnoreCase(uri.getHost())) {
                Set<String> names = uri.getQueryParameterNames();
                if (names.contains("phone")) phone = digitsOnly(uri.getQueryParameter("phone"));
            }

            return new WhatsAppPayload(text, digitsOnly(phone));
        } catch (Exception ignored) {
            return new WhatsAppPayload(value, "");
        }
    }

    private String firstNonEmpty(String first, String second) {
        if (first != null && !first.trim().isEmpty()) return first;
        return second == null ? "" : second;
    }

    private String digitsOnly(String value) {
        if (value == null) return "";
        return value.replaceAll("[^0-9]", "");
    }

    private static final class WhatsAppPayload {
        final String text;
        final String phone;

        WhatsAppPayload(String text, String phone) {
            this.text = text == null ? "" : text;
            this.phone = phone == null ? "" : phone;
        }
    }
}
