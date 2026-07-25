package it.vargacantieri.hera.geofence;

import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.util.Log;

public class HeraGeofenceBootReceiver extends BroadcastReceiver {
    private static final String TAG = "HeraGeofenceBoot";

    @Override
    public void onReceive(Context context, Intent intent) {
        if (intent == null || intent.getAction() == null) {
            return;
        }

        String action = intent.getAction();
        if (!Intent.ACTION_BOOT_COMPLETED.equals(action)
                && !Intent.ACTION_MY_PACKAGE_REPLACED.equals(action)) {
            return;
        }

        HeraGeofenceManager manager = new HeraGeofenceManager(context);
        if (!manager.isActive()) {
            return;
        }

        PendingResult pendingResult = goAsync();
        manager.registerGeofence(new HeraGeofenceManager.GeofenceRegistrationCallback() {
            @Override
            public void onSuccess() {
                Log.i(TAG, "Geofence restored after boot or app update.");
                pendingResult.finish();
            }

            @Override
            public void onFailure(Exception exception) {
                Log.e(TAG, "Unable to restore geofence after boot or app update.", exception);
                pendingResult.finish();
            }
        });
    }
}
