package it.vargacantieri.hera.geofence;

import android.Manifest;
import android.app.PendingIntent;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.os.Build;

import androidx.core.content.ContextCompat;

import com.google.android.gms.location.Geofence;
import com.google.android.gms.location.GeofencingClient;
import com.google.android.gms.location.GeofencingRequest;
import com.google.android.gms.location.LocationServices;

public class HeraGeofenceManager {
    private final Context context;
    private final GeofencingClient geofencingClient;

    public HeraGeofenceManager(Context context) {
        this.context = context.getApplicationContext();
        this.geofencingClient = LocationServices.getGeofencingClient(this.context);
    }

    public void setActive(boolean active) {
        SharedPreferences preferences = context.getSharedPreferences(HeraGeofenceConstants.PREFS_NAME, Context.MODE_PRIVATE);
        preferences.edit().putBoolean(HeraGeofenceConstants.PREF_ACTIVE, active).apply();
    }

    public boolean isActive() {
        SharedPreferences preferences = context.getSharedPreferences(HeraGeofenceConstants.PREFS_NAME, Context.MODE_PRIVATE);
        return preferences.getBoolean(HeraGeofenceConstants.PREF_ACTIVE, false);
    }

    public boolean hasRequiredPermissions() {
        boolean fineLocation = ContextCompat.checkSelfPermission(context, Manifest.permission.ACCESS_FINE_LOCATION)
                == PackageManager.PERMISSION_GRANTED;
        boolean backgroundLocation = Build.VERSION.SDK_INT < Build.VERSION_CODES.Q
                || ContextCompat.checkSelfPermission(context, Manifest.permission.ACCESS_BACKGROUND_LOCATION)
                == PackageManager.PERMISSION_GRANTED;
        return fineLocation && backgroundLocation;
    }

    public void registerGeofence() {
        if (!hasRequiredPermissions()) {
            return;
        }

        Geofence geofence = new Geofence.Builder()
                .setRequestId(HeraGeofenceConstants.GEOFENCE_REQUEST_ID)
                .setCircularRegion(
                        HeraGeofenceConstants.TARGET_LAT,
                        HeraGeofenceConstants.TARGET_LNG,
                        HeraGeofenceConstants.TARGET_RADIUS_METERS
                )
                .setTransitionTypes(Geofence.GEOFENCE_TRANSITION_ENTER | Geofence.GEOFENCE_TRANSITION_DWELL)
                .setLoiteringDelay(60_000)
                .setExpirationDuration(Geofence.NEVER_EXPIRE)
                .build();

        GeofencingRequest request = new GeofencingRequest.Builder()
                .setInitialTrigger(GeofencingRequest.INITIAL_TRIGGER_ENTER)
                .addGeofence(geofence)
                .build();

        geofencingClient.removeGeofences(getPendingIntent());
        geofencingClient.addGeofences(request, getPendingIntent());
    }

    public void unregisterGeofence() {
        geofencingClient.removeGeofences(getPendingIntent());
    }

    private PendingIntent getPendingIntent() {
        Intent intent = new Intent(context, HeraGeofenceBroadcastReceiver.class);
        intent.setAction("it.vargacantieri.hera.GEOFENCE_EVENT");

        int flags = PendingIntent.FLAG_UPDATE_CURRENT;
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            flags |= PendingIntent.FLAG_IMMUTABLE;
        }

        return PendingIntent.getBroadcast(context, 1106, intent, flags);
    }
}
