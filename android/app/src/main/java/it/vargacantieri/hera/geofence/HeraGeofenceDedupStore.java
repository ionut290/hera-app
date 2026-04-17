package it.vargacantieri.hera.geofence;

import android.content.Context;
import android.content.SharedPreferences;

import java.time.LocalDate;

public class HeraGeofenceDedupStore {
    private final SharedPreferences preferences;

    public HeraGeofenceDedupStore(Context context) {
        this.preferences = context.getSharedPreferences(HeraGeofenceConstants.PREFS_NAME, Context.MODE_PRIVATE);
    }

    public boolean shouldNotify(String slot, LocalDate date) {
        String dedupKey = buildDedupKey(slot, date);
        return !preferences.getBoolean(dedupKey, false);
    }

    public void markNotified(String slot, LocalDate date) {
        String dedupKey = buildDedupKey(slot, date);
        preferences.edit().putBoolean(dedupKey, true).apply();
    }

    private String buildDedupKey(String slot, LocalDate date) {
        return "notified:" + date + ":" + slot;
    }
}
