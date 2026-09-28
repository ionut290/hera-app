package it.vargacantieri.hera.navigation;

import android.content.ActivityNotFoundException;
import android.content.Intent;
import android.net.Uri;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.util.Locale;

@CapacitorPlugin(name = "HeraNavigation")
public class HeraNavigationPlugin extends Plugin {
    @PluginMethod
    public void openWaze(PluginCall call) {
        Double latitude = call.getDouble("latitude");
        Double longitude = call.getDouble("longitude");
        if (latitude == null || longitude == null
                || !Double.isFinite(latitude) || !Double.isFinite(longitude)
                || latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) {
            call.reject("Coordinate di navigazione non valide.");
            return;
        }

        String point = String.format(Locale.US, "%.7f,%.7f", latitude, longitude);
        Uri uri = Uri.parse("waze://?ll=" + point + "&navigate=yes");
        Intent intent = new Intent(Intent.ACTION_VIEW, uri);
        intent.setPackage("com.waze");
        try {
            getActivity().startActivity(intent);
            JSObject result = new JSObject();
            result.put("opened", true);
            call.resolve(result);
        } catch (ActivityNotFoundException | SecurityException error) {
            call.reject("Waze non è installato o non può essere aperto su questo telefono.", error);
        }
    }
}
