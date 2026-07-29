package it.vargacantieri.hera;

import android.os.Bundle;

import com.getcapacitor.BridgeActivity;

import it.vargacantieri.hera.geofence.HeraGeofencePlugin;
import it.vargacantieri.hera.biometric.HeraBiometricPlugin;
import it.vargacantieri.hera.whatsapp.HeraWhatsAppPlugin;

public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(HeraGeofencePlugin.class);
        registerPlugin(HeraBiometricPlugin.class);
        registerPlugin(HeraWhatsAppPlugin.class);
        super.onCreate(savedInstanceState);
    }
}
