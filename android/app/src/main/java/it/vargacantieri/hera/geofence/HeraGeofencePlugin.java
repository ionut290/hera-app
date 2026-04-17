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

        if (!manager.hasRequiredPermissions()) {
            call.reject("Permessi posizione mancanti (foreground/background).");
            return;
        }

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
        result.put("hasLocationPermission", manager.hasRequiredPermissions());
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
            result.put("needsNotificationPermission", getPermissionState("notifications") != PermissionState.GRANTED);
        } else {
            result.put("needsNotificationPermission", false);
        }
        call.resolve(result);
    }
}
