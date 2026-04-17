package it.vargacantieri.hera.geofence;

public final class HeraGeofenceConstants {
    private HeraGeofenceConstants() {}

    public static final String GEOFENCE_REQUEST_ID = "hera-site-main";
    public static final double TARGET_LAT = 44.562504656236015d;
    public static final double TARGET_LNG = 11.356961975643515d;
    public static final float TARGET_RADIUS_METERS = 200f;

    public static final String NOTIFICATION_CHANNEL_ID = "hera_geofence_channel";
    public static final String NOTIFICATION_CHANNEL_NAME = "Hera Geofence";

    public static final String PREFS_NAME = "hera_geofence_prefs";
    public static final String PREF_ACTIVE = "active";

    public static final String SLOT_ENTRY = "entry";
    public static final String SLOT_EXIT = "exit";

    public static final int ENTRY_START_MINUTES = 6 * 60 + 15;
    public static final int ENTRY_END_MINUTES = 7 * 60 + 30;

    public static final int EXIT_START_MINUTES = 15 * 60 + 30;
    public static final int EXIT_END_MINUTES = 17 * 60;
}
