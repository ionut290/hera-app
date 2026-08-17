package it.vargacantieri.hera.camera;

import android.Manifest;
import android.app.Activity;
import android.content.Intent;
import android.os.Handler;
import android.os.Looper;

import androidx.activity.result.ActivityResult;

import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.PermissionState;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.ActivityCallback;
import com.getcapacitor.annotation.CapacitorPlugin;
import com.getcapacitor.annotation.Permission;
import com.getcapacitor.annotation.PermissionCallback;

import java.io.File;
import java.util.ArrayList;

@CapacitorPlugin(
    name = "HeraContinuousCamera",
    permissions = {
        @Permission(alias = "camera", strings = { Manifest.permission.CAMERA })
    }
)
public class HeraContinuousCameraPlugin extends Plugin {
    private static final int DEFAULT_MAX_PHOTOS = 10;
    private static final long CLEANUP_DELAY_MS = 10 * 60 * 1000L;

    @PluginMethod
    public void capture(PluginCall call) {
        if (getPermissionState("camera") != PermissionState.GRANTED) {
            requestPermissionForAlias("camera", call, "cameraPermissionCallback");
            return;
        }
        launchCamera(call);
    }

    @PermissionCallback
    private void cameraPermissionCallback(PluginCall call) {
        if (getPermissionState("camera") != PermissionState.GRANTED) {
            JSObject result = new JSObject();
            result.put("cancelled", true);
            result.put("permissionDenied", true);
            result.put("photos", new JSArray());
            call.resolve(result);
            return;
        }
        launchCamera(call);
    }

    private void launchCamera(PluginCall call) {
        int requested = call.getInt("maxPhotos", DEFAULT_MAX_PHOTOS);
        int maxPhotos = Math.max(1, Math.min(requested, DEFAULT_MAX_PHOTOS));
        Intent intent = new Intent(getActivity(), HeraContinuousCameraActivity.class);
        intent.putExtra(HeraContinuousCameraActivity.EXTRA_MAX_PHOTOS, maxPhotos);
        startActivityForResult(call, intent, "cameraActivityResult");
    }

    @ActivityCallback
    private void cameraActivityResult(PluginCall call, ActivityResult result) {
        if (call == null) return;

        Intent data = result.getData();
        if (result.getResultCode() != Activity.RESULT_OK || data == null) {
            JSObject cancelled = new JSObject();
            cancelled.put("cancelled", true);
            cancelled.put("photos", new JSArray());
            call.resolve(cancelled);
            return;
        }

        ArrayList<String> paths = data.getStringArrayListExtra(HeraContinuousCameraActivity.EXTRA_PHOTO_PATHS);
        JSArray photos = new JSArray();
        if (paths != null) {
            for (String path : paths) {
                if (path == null || path.trim().isEmpty()) continue;
                File file = new File(path);
                if (!file.isFile()) continue;
                JSObject photo = new JSObject();
                photo.put("path", file.getAbsolutePath());
                photo.put("name", file.getName());
                photo.put("type", "image/jpeg");
                photos.put(photo);
                scheduleCleanup(file);
            }
        }

        JSObject response = new JSObject();
        response.put("cancelled", false);
        response.put("photos", photos);
        call.resolve(response);
    }

    private void scheduleCleanup(File file) {
        new Handler(Looper.getMainLooper()).postDelayed(() -> {
            try {
                File parent = file.getParentFile();
                file.delete();
                if (parent != null) parent.delete();
            } catch (Exception ignored) {}
        }, CLEANUP_DELAY_MS);
    }
}
