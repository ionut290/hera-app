package it.vargacantieri.hera.geofence;

import android.Manifest;
import android.os.Build;

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

    @PluginMethod
    public void activate(PluginCall call) {
        HeraGeofenceManager manager = new HeraGeofenceManager(getContext());

        if (!manager.hasRequiredPermissions()) {
            call.reject("Permessi posizione mancanti (foreground/background).");
            return;
        }

        HeraGeofenceNotifier notifier = new HeraGeofenceNotifier(getContext());
        notifier.ensureChannel();

        manager.setActive(true);
        manager.registerGeofence();

        JSObject result = new JSObject();
        result.put("active", true);
        call.resolve(result);
    }

    @PluginMethod
    public void deactivate(PluginCall call) {
        HeraGeofenceManager manager = new HeraGeofenceManager(getContext());
        manager.unregisterGeofence();
        manager.setActive(false);

        JSObject result = new JSObject();
        result.put("active", false);
        call.resolve(result);
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
