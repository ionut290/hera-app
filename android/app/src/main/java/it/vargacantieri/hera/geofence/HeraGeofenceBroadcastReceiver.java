package it.vargacantieri.hera.geofence;

import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;

import com.google.android.gms.location.Geofence;
import com.google.android.gms.location.GeofencingEvent;

import java.time.LocalDate;
import java.time.LocalTime;

public class HeraGeofenceBroadcastReceiver extends BroadcastReceiver {
    @Override
    public void onReceive(Context context, Intent intent) {
        GeofencingEvent event = GeofencingEvent.fromIntent(intent);
        if (event == null || event.hasError()) {
            return;
        }

        int transition = event.getGeofenceTransition();
        if (transition != Geofence.GEOFENCE_TRANSITION_ENTER
                && transition != Geofence.GEOFENCE_TRANSITION_DWELL) {
            return;
        }

        HeraGeofenceManager manager = new HeraGeofenceManager(context);
        if (!manager.isActive()) {
            return;
        }

        LocalTime now = LocalTime.now();
        String slot = resolveSlot(now);
        if (slot == null) {
            return;
        }

        LocalDate today = LocalDate.now();
        HeraGeofenceDedupStore dedupStore = new HeraGeofenceDedupStore(context);
        if (!dedupStore.shouldNotify(slot, today)) {
            return;
        }

        HeraGeofenceNotifier notifier = new HeraGeofenceNotifier(context);
        notifier.ensureChannel();
        notifier.notifySlot(slot);
        dedupStore.markNotified(slot, today);
    }

    private String resolveSlot(LocalTime localTime) {
        int minutes = localTime.getHour() * 60 + localTime.getMinute();

        if (minutes >= HeraGeofenceConstants.ENTRY_START_MINUTES
                && minutes <= HeraGeofenceConstants.ENTRY_END_MINUTES) {
            return HeraGeofenceConstants.SLOT_ENTRY;
        }

        if (minutes >= HeraGeofenceConstants.EXIT_START_MINUTES
                && minutes <= HeraGeofenceConstants.EXIT_END_MINUTES) {
            return HeraGeofenceConstants.SLOT_EXIT;
        }

        return null;
    }
}
