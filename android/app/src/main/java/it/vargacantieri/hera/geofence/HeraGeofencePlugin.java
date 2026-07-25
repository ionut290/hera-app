package it.vargacantieri.hera.geofence;

import android.Manifest;
import android.os.Build;
import android.util.Log;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.PermissionState;
import com.getcapacitor.annotation.CapacitorPlugin;
import com.getcapacitor.annotation.Permission;
import com.getcapacitor.annotation.PermissionCallback;

@CapacitorPlugin(
        name = "HeraGeofence",
        permissions = {
                @Permission(strings = {Manifest.permission.ACCESS_COARSE_LOCATION, Manifest.permission.ACCESS_FINE_LOCATION}, alias = "location"),
                @Permission(strings = {Manifest.permission.ACCESS_BACKGROUND_LOCATION}, alias = "backgroundLocation"),
                @Permission(strings = {Manifest.permission.POST_NOTIFICATIONS}, alias = "notifications")
        }
)
public class HeraGeofencePlugin extends Plugin {
    private static final String TAG = "HeraGeofencePlugin";

    @PluginMethod
    public void activate(PluginCall call) {
        HeraGeofenceManager manager = new HeraGeofenceManager(getContext());

        if (!manager.hasForegroundLocationPermission()) {
            requestPermissionForAlias("location", call, "foregroundPermissionCallback");
            return;
        }

        if (!manager.hasBackgroundLocationPermission()) {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
                JSObject result = new JSObject();
                result.put("active", false);
                result.put("needsBackgroundSettings", true);
                result.put("message", "Consenti Sempre dalle impostazioni Android per attivare la posizione in background.");
                call.resolve(result);
            } else {
                requestPermissionForAlias("backgroundLocation", call, "backgroundPermissionCallback");
            }
            return;
        }

        activateGeofence(call, manager);
    }

    @PermissionCallback
    private void foregroundPermissionCallback(PluginCall call) {
        HeraGeofenceManager manager = new HeraGeofenceManager(getContext());
        if (!manager.hasForegroundLocationPermission()) {
            call.reject("Permesso posizione Android non concesso.");
            return;
        }
        activate(call);
    }

    @PermissionCallback
    private void backgroundPermissionCallback(PluginCall call) {
        HeraGeofenceManager manager = new HeraGeofenceManager(getContext());
        if (!manager.hasBackgroundLocationPermission()) {
            call.reject("Permesso posizione in background non concesso.");
            return;
        }
        activateGeofence(call, manager);
    }

    private void activateGeofence(PluginCall call, HeraGeofenceManager manager) {
        HeraGeofenceNotifier notifier = new HeraGeofenceNotifier(getContext());
        notifier.ensureChannel();

        manager.registerGeofence(new HeraGeofenceManager.GeofenceRegistrationCallback() {
            @Override
            public void onSuccess() {
                manager.setActive(true);
                JSObject result = new JSObject();
                result.put("active", true);
                call.resolve(result);
            }

            @Override
            public void onFailure(Exception exception) {
                manager.setActive(false);
                Log.e(TAG, "Geofence activation failed.", exception);
                call.reject("Attivazione geofence fallita.", exception);
            }
        });
    }

    @PluginMethod
    public void deactivate(PluginCall call) {
        HeraGeofenceManager manager = new HeraGeofenceManager(getContext());
        boolean wasActive = manager.isActive();

        manager.unregisterGeofence(new HeraGeofenceManager.GeofenceRegistrationCallback() {
            @Override
            public void onSuccess() {
                manager.setActive(false);
                JSObject result = new JSObject();
                result.put("active", false);
                call.resolve(result);
            }

            @Override
            public void onFailure(Exception exception) {
                manager.setActive(wasActive);
                Log.e(TAG, "Geofence deactivation failed.", exception);
                call.reject("Disattivazione geofence fallita.", exception);
            }
        });
    }

    @PluginMethod
    public void status(PluginCall call) {
        HeraGeofenceManager manager = new HeraGeofenceManager(getContext());
        JSObject result = new JSObject();
        result.put("active", manager.isActive());
        result.put("hasLocationPermission", manager.hasForegroundLocationPermission());
        result.put("hasBackgroundLocationPermission", manager.hasBackgroundLocationPermission());
        result.put("needsBackgroundSettings", Build.VERSION.SDK_INT >= Build.VERSION_CODES.R
                && !manager.hasBackgroundLocationPermission());
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
            result.put("needsNotificationPermission", getPermissionState("notifications") != PermissionState.GRANTED);
        } else {
            result.put("needsNotificationPermission", false);
        }
        call.resolve(result);
    }
}
