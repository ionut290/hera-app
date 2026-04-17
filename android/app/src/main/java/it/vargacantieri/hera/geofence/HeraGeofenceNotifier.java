package it.vargacantieri.hera.geofence;

import android.Manifest;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.content.Context;
import android.content.pm.PackageManager;
import android.os.Build;

import androidx.core.app.NotificationCompat;
import androidx.core.app.NotificationManagerCompat;
import androidx.core.content.ContextCompat;


public class HeraGeofenceNotifier {
    private final Context context;

    public HeraGeofenceNotifier(Context context) {
        this.context = context.getApplicationContext();
    }

    public void ensureChannel() {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.O) {
            return;
        }

        NotificationChannel channel = new NotificationChannel(
                HeraGeofenceConstants.NOTIFICATION_CHANNEL_ID,
                HeraGeofenceConstants.NOTIFICATION_CHANNEL_NAME,
                NotificationManager.IMPORTANCE_DEFAULT
        );
        channel.setDescription("Notifiche presenza/uscita geofence Hera");

        NotificationManager manager = context.getSystemService(NotificationManager.class);
        if (manager != null) {
            manager.createNotificationChannel(channel);
        }
    }

    public void notifySlot(String slot) {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU
                && ContextCompat.checkSelfPermission(context, Manifest.permission.POST_NOTIFICATIONS)
                != PackageManager.PERMISSION_GRANTED) {
            return;
        }

        String title;
        String body;
        int notificationId;

        if (HeraGeofenceConstants.SLOT_ENTRY.equals(slot)) {
            title = "Promemoria entrata";
            body = "Sei nell'area cantiere nella fascia 06:15-07:30.";
            notificationId = 615730;
        } else {
            title = "Promemoria uscita";
            body = "Sei nell'area cantiere nella fascia 15:30-17:00.";
            notificationId = 1530170;
        }

        NotificationCompat.Builder builder = new NotificationCompat.Builder(context, HeraGeofenceConstants.NOTIFICATION_CHANNEL_ID)
                .setSmallIcon(android.R.drawable.ic_dialog_info)
                .setContentTitle(title)
                .setContentText(body)
                .setPriority(NotificationCompat.PRIORITY_DEFAULT)
                .setAutoCancel(true);

        NotificationManagerCompat.from(context).notify(notificationId, builder.build());
    }
}
