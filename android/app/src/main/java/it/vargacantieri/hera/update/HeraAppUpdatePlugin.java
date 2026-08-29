package it.vargacantieri.hera.update;

import android.content.pm.PackageInfo;
import android.os.Build;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import com.google.android.play.core.appupdate.AppUpdateInfo;
import com.google.android.play.core.appupdate.AppUpdateManager;
import com.google.android.play.core.appupdate.AppUpdateManagerFactory;
import com.google.android.play.core.appupdate.AppUpdateOptions;
import com.google.android.play.core.install.model.AppUpdateType;
import com.google.android.play.core.install.model.UpdateAvailability;

@CapacitorPlugin(name = "HeraAppUpdate")
public class HeraAppUpdatePlugin extends Plugin {
    private static final int UPDATE_REQUEST_CODE = 9438;

    private AppUpdateManager manager() {
        return AppUpdateManagerFactory.create(getContext());
    }

    private long currentVersionCode() {
        try {
            PackageInfo info = getContext().getPackageManager().getPackageInfo(getContext().getPackageName(), 0);
            return Build.VERSION.SDK_INT >= Build.VERSION_CODES.P ? info.getLongVersionCode() : info.versionCode;
        } catch (Exception ignored) {
            return 0L;
        }
    }

    private JSObject resultFor(AppUpdateInfo info) {
        JSObject result = new JSObject();
        boolean available = info.updateAvailability() == UpdateAvailability.UPDATE_AVAILABLE;
        result.put("available", available);
        result.put("currentVersionCode", currentVersionCode());
        result.put("availableVersionCode", info.availableVersionCode());
        result.put("immediateAllowed", available && info.isUpdateTypeAllowed(AppUpdateType.IMMEDIATE));
        return result;
    }

    @PluginMethod
    public void checkForUpdate(PluginCall call) {
        manager().getAppUpdateInfo()
            .addOnSuccessListener(info -> call.resolve(resultFor(info)))
            .addOnFailureListener(error -> call.reject("Controllo aggiornamento Google Play non riuscito.", error));
    }

    @PluginMethod
    public void startUpdate(PluginCall call) {
        AppUpdateManager updateManager = manager();
        updateManager.getAppUpdateInfo()
            .addOnSuccessListener(info -> {
                JSObject result = resultFor(info);
                boolean available = info.updateAvailability() == UpdateAvailability.UPDATE_AVAILABLE;
                boolean immediateAllowed = available && info.isUpdateTypeAllowed(AppUpdateType.IMMEDIATE);
                if (!available || !immediateAllowed) {
                    result.put("started", false);
                    call.resolve(result);
                    return;
                }
                try {
                    boolean started = updateManager.startUpdateFlowForResult(
                        info,
                        getActivity(),
                        AppUpdateOptions.newBuilder(AppUpdateType.IMMEDIATE).build(),
                        UPDATE_REQUEST_CODE
                    );
                    result.put("started", started);
                    call.resolve(result);
                } catch (Exception error) {
                    call.reject("Apertura aggiornamento Google Play non riuscita.", error);
                }
            })
            .addOnFailureListener(error -> call.reject("Controllo aggiornamento Google Play non riuscito.", error));
    }
}
