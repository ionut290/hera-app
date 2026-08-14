package it.vargacantieri.hera.whatsapp;

import android.content.ActivityNotFoundException;
import android.content.ClipData;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Handler;
import android.os.Looper;
import android.util.Base64;

import androidx.activity.result.ActivityResult;
import androidx.core.content.FileProvider;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.ActivityCallback;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.UUID;

@CapacitorPlugin(name = "HeraWhazzupPhotos")
public class HeraWhazzupPhotosPlugin extends Plugin {
    private static final String WHATSAPP = "com.whatsapp";
    private static final String WHATSAPP_BUSINESS = "com.whatsapp.w4b";
    private static final String SHARE_ROOT = "hera-whazzup-native";
    private static final long CLEANUP_DELAY_MS = 5 * 60 * 1000L;
    private static final long EXPIRED_SESSION_MS = 12 * 60 * 60 * 1000L;
    private static final int MAX_PHOTO_BYTES = 35 * 1024 * 1024;

    @PluginMethod
    public void begin(PluginCall call) {
        purgeExpiredSessions();
        String sessionId = UUID.randomUUID().toString();
        File folder = sessionFolder(sessionId);
        if (!folder.mkdirs() && !folder.isDirectory()) {
            call.reject("Impossibile preparare le foto sul dispositivo.");
            return;
        }
        JSObject result = new JSObject();
        result.put("sessionId", sessionId);
        call.resolve(result);
    }

    @PluginMethod
    public void addPhoto(PluginCall call) {
        String sessionId = validatedSessionId(call.getString("sessionId", ""));
        String data = call.getString("data", "");
        String fileName = safeFileName(call.getString("fileName", "foto.jpg"));
        if (sessionId == null || data.isEmpty()) {
            call.reject("Foto o sessione non valida.");
            return;
        }

        byte[] bytes;
        try {
            bytes = Base64.decode(data, Base64.DEFAULT);
        } catch (IllegalArgumentException error) {
            call.reject("Formato della foto non valido.", error);
            return;
        }
        if (bytes.length == 0 || bytes.length > MAX_PHOTO_BYTES) {
            call.reject("La foto è vuota o troppo grande.");
            return;
        }

        File folder = sessionFolder(sessionId);
        if (!folder.isDirectory() && !folder.mkdirs()) {
            call.reject("Archivio temporaneo delle foto non disponibile.");
            return;
        }
        File destination = new File(folder, fileName);
        try (FileOutputStream output = new FileOutputStream(destination, false)) {
            output.write(bytes);
            output.flush();
            JSObject result = new JSObject();
            result.put("stored", true);
            result.put("fileName", fileName);
            call.resolve(result);
        } catch (IOException error) {
            call.reject("Salvataggio temporaneo della foto non riuscito.", error);
        }
    }

    @PluginMethod
    public void share(PluginCall call) {
        String sessionId = validatedSessionId(call.getString("sessionId", ""));
        if (sessionId == null) {
            call.reject("Sessione foto Whazzup non disponibile.");
            return;
        }

        String packageName = resolveInstalledPackage();
        if (packageName == null) {
            call.reject("WhatsApp non è installato sul dispositivo.");
            return;
        }

        File folder = sessionFolder(sessionId);
        File[] storedFiles = folder.listFiles(File::isFile);
        if (storedFiles == null || storedFiles.length == 0) {
            call.reject("Nessuna foto disponibile per la condivisione.");
            return;
        }
        Arrays.sort(storedFiles, Comparator.comparing(File::getName));

        ArrayList<Uri> photoUris = new ArrayList<>();
        try {
            for (File file : storedFiles) {
                photoUris.add(FileProvider.getUriForFile(
                    getContext(),
                    getContext().getPackageName() + ".fileprovider",
                    file
                ));
            }

            Intent intent = new Intent(photoUris.size() > 1 ? Intent.ACTION_SEND_MULTIPLE : Intent.ACTION_SEND);
            intent.setType("image/*");
            intent.setPackage(packageName);
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);

            ClipData clipData = ClipData.newRawUri("Foto impianto", photoUris.get(0));
            for (int index = 1; index < photoUris.size(); index++) {
                clipData.addItem(new ClipData.Item(photoUris.get(index)));
            }
            intent.setClipData(clipData);

            if (photoUris.size() == 1) {
                intent.putExtra(Intent.EXTRA_STREAM, photoUris.get(0));
            } else {
                intent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, photoUris);
            }

            for (Uri uri : photoUris) {
                getContext().grantUriPermission(packageName, uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);
            }

            startActivityForResult(call, intent, "photosActivityResult");
            scheduleCleanup(folder);
        } catch (ActivityNotFoundException error) {
            call.reject("Impossibile aprire WhatsApp installato con le foto.", error);
        } catch (Exception error) {
            call.reject("Errore durante la condivisione nativa delle foto.", error);
        }
    }

    @ActivityCallback
    private void photosActivityResult(PluginCall call, ActivityResult result) {
        if (call == null) return;
        JSObject callResult = new JSObject();
        callResult.put("opened", true);
        callResult.put("returned", true);
        callResult.put("separateTextRequired", true);
        callResult.put("fallback", false);
        call.resolve(callResult);
    }

    @PluginMethod
    public void discard(PluginCall call) {
        String sessionId = validatedSessionId(call.getString("sessionId", ""));
        if (sessionId != null) deleteRecursively(sessionFolder(sessionId));
        call.resolve();
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

    private File shareRoot() {
        return new File(getContext().getCacheDir(), SHARE_ROOT);
    }

    private File sessionFolder(String sessionId) {
        return new File(shareRoot(), sessionId);
    }

    private String validatedSessionId(String value) {
        String sessionId = value == null ? "" : value.trim();
        return sessionId.matches("[a-fA-F0-9-]{36}") ? sessionId : null;
    }

    private String safeFileName(String value) {
        String fileName = value == null ? "foto.jpg" : value.trim();
        fileName = fileName.replaceAll("[^a-zA-Z0-9._-]", "_");
        if (fileName.isEmpty() || fileName.equals(".") || fileName.equals("..")) return "foto.jpg";
        return fileName.length() > 120 ? fileName.substring(fileName.length() - 120) : fileName;
    }

    private void scheduleCleanup(File folder) {
        new Handler(Looper.getMainLooper()).postDelayed(() -> deleteRecursively(folder), CLEANUP_DELAY_MS);
    }

    private void purgeExpiredSessions() {
        File root = shareRoot();
        File[] folders = root.listFiles(File::isDirectory);
        if (folders == null) return;
        long cutoff = System.currentTimeMillis() - EXPIRED_SESSION_MS;
        for (File folder : folders) {
            if (folder.lastModified() < cutoff) deleteRecursively(folder);
        }
    }

    private void deleteRecursively(File file) {
        if (file == null || !file.exists()) return;
        File[] children = file.listFiles();
        if (children != null) {
            for (File child : children) deleteRecursively(child);
        }
        file.delete();
    }
}
