import Foundation
import UIKit
import Capacitor
import CoreLocation

@objc(VargaGeofencePlugin)
public class VargaGeofencePlugin: CAPPlugin, CAPBridgedPlugin {
    public let identifier = "VargaGeofencePlugin"
    public let jsName = "VargaGeofence"
    public let pluginMethods: [CAPPluginMethod] = [
        CAPPluginMethod(name: "configure", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "status", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "disable", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "openEasyBoxShortcut", returnType: CAPPluginReturnPromise)
    ]

    // Called only after an explicit user tap on a Varga timbratura notification
    // or on the manual "Apri Easy-Box" button. This is an Apple-documented
    // shortcut URL, not an unverified deep link owned by ICT-GROUP.
    @objc func openEasyBoxShortcut(_ call: CAPPluginCall) {
        guard let url = URL(string: "shortcuts://run-shortcut?name=Apri%20Easy-Box") else {
            call.reject("Collegamento Comandi Rapidi non valido.")
            return
        }
        DispatchQueue.main.async {
            UIApplication.shared.open(url, options: [:]) { opened in
                call.resolve(["opened": opened])
            }
        }
    }

    @objc func configure(_ call: CAPPluginCall) {
        guard let latitude = call.getDouble("latitude"),
              let longitude = call.getDouble("longitude"),
              let radius = call.getDouble("radius"),
              let from = call.getString("arrivalFrom"),
              let to = call.getString("arrivalTo"),
              let returnFrom = call.getString("returnFrom"),
              let returnTo = call.getString("returnTo"),
              let rawDays = call.getArray("days"),
              let days = rawDays as? [Int] else {
            call.reject("Dati GPS, raggio, giorni o orari mancanti.")
            return
        }
        let values = VargaGeofenceSettings(enabled: true, latitude: latitude,
            longitude: longitude, radius: radius, days: days,
            arrivalFrom: from, arrivalTo: to, returnFrom: returnFrom, returnTo: returnTo)
        DispatchQueue.main.async {
            do {
                call.resolve(try VargaGeofenceManager.shared.configure(values))
            } catch {
                call.reject(error.localizedDescription)
            }
        }
    }

    @objc func status(_ call: CAPPluginCall) {
        let consume = call.getBool("consumePending") ?? false
        DispatchQueue.main.async {
            call.resolve(VargaGeofenceManager.shared.status(consumePending: consume))
        }
    }

    @objc func disable(_ call: CAPPluginCall) {
        DispatchQueue.main.async {
            call.resolve(VargaGeofenceManager.shared.disable())
        }
    }
}
