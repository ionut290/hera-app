package it.vargacantieri.hera;

import android.os.Bundle;

import com.getcapacitor.BridgeActivity;

import it.vargacantieri.hera.geofence.HeraGeofencePlugin;
import it.vargacantieri.hera.biometric.HeraBiometricPlugin;

public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(HeraGeofencePlugin.class);
        registerPlugin(HeraBiometricPlugin.class);
        super.onCreate(savedInstanceState);
    }
}
