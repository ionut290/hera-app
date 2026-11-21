import Foundation
import UIKit
import CoreLocation
import UserNotifications

// iOS owns the circular region even if the Capacitor WKWebView is suspended.
// No continuous location tracking, network transfers or automatic clock-in.
struct VargaGeofenceSettings: Codable {
    let enabled: Bool
    let latitude: Double
    let longitude: Double
    let radius: Double
    let days: [Int] // 0=Sunday, same convention as JavaScript Date.getDay()
    let arrivalFrom: String
    let arrivalTo: String
    let returnFrom: String
    let returnTo: String
}

final class VargaGeofenceManager: NSObject, CLLocationManagerDelegate {
    static let shared = VargaGeofenceManager()
    private let key = "varga.geofence.config.v1"
    private let pendingKey = "varga.geofence.pending.v1"
    private let regionID = "it.vargacantieri.magazzino.timbratura"
    private let defaults = UserDefaults.standard
    private lazy var locationManager: CLLocationManager = {
        let value = CLLocationManager()
        value.delegate = self
        return value
    }()

    private override init() {
        super.init()
    }

    // Called from AppDelegate at app launch, including background relaunch on
    // a Core Location region event. This MUST NOT wait for the WebView to load.
    func bootstrap() {
        assert(Thread.isMainThread)
        _ = locationManager
        if let settings = readSettings(), settings.enabled {
            reconcileMonitoring(settings)
            // A background relaunch may occur while the phone is already inside.
            // Ask iOS for the current region state instead of relying only on
            // a new boundary crossing.
            for region in locationManager.monitoredRegions where region.identifier == regionID {
                locationManager.requestState(for: region)
            }
        }
    }

    private func readSettings() -> VargaGeofenceSettings? {
        guard let data = defaults.data(forKey: key) else { return nil }
        return try? JSONDecoder().decode(VargaGeofenceSettings.self, from: data)
    }

    func configure(_ settings: VargaGeofenceSettings) throws -> [String: Any] {
        assert(Thread.isMainThread)
        guard settings.latitude.isFinite, (-90...90).contains(settings.latitude),
              settings.longitude.isFinite, (-180...180).contains(settings.longitude),
              settings.radius.isFinite, (100...2000).contains(settings.radius),
              !settings.days.isEmpty, settings.days.allSatisfy({ (0...6).contains($0) }),
              validWindow(settings.arrivalFrom, settings.arrivalTo),
              validWindow(settings.returnFrom, settings.returnTo) else {
            throw NSError(domain: "VargaGeofence", code: 1,
                          userInfo: [NSLocalizedDescriptionKey: "Coordinate, raggio (100–2000 m), giorni oppure orari non validi."])
        }
        guard CLLocationManager.isMonitoringAvailable(for: CLCircularRegion.self) else {
            throw NSError(domain: "VargaGeofence", code: 2,
                          userInfo: [NSLocalizedDescriptionKey: "Monitoraggio geografico non disponibile su questo iPhone."])
        }
        let data = try JSONEncoder().encode(settings)
        defaults.set(data, forKey: key)
        // Explicit user opt-in from the Save button; iOS manages both dialogs.
        UNUserNotificationCenter.current().requestAuthorization(options: [.alert, .sound, .badge]) { _, _ in }
        if locationManager.authorizationStatus != .authorizedAlways {
            locationManager.requestAlwaysAuthorization()
        }
        reconcileMonitoring(settings)
        return status(consumePending: false)
    }

    func disable() -> [String: Any] {
        assert(Thread.isMainThread)
        defaults.removeObject(forKey: key)
        defaults.removeObject(forKey: pendingKey)
        stopOurRegion()
        return status(consumePending: false)
    }

    func status(consumePending: Bool) -> [String: Any] {
        assert(Thread.isMainThread)
        let permission: String
        switch locationManager.authorizationStatus {
        case .authorizedAlways: permission = "always"
        case .authorizedWhenInUse: permission = "whenInUse"
        case .denied: permission = "denied"
        case .restricted: permission = "restricted"
        case .notDetermined: permission = "notDetermined"
        @unknown default: permission = "unknown"
        }
        let regionActive = locationManager.monitoredRegions.contains { $0.identifier == regionID }
        var output: [String: Any] = [
            "supported": CLLocationManager.isMonitoringAvailable(for: CLCircularRegion.self),
            "permission": permission,
            "configured": readSettings()?.enabled == true,
            "monitoring": regionActive,
            "radius": (locationManager.monitoredRegions.first { $0.identifier == regionID } as? CLCircularRegion)?.radius ?? 0
        ]
        if let pending = defaults.dictionary(forKey: pendingKey) {
            output["pending"] = pending
            if consumePending { defaults.removeObject(forKey: pendingKey) }
        }
        return output
    }

    private func stopOurRegion() {
        for region in locationManager.monitoredRegions where region.identifier == regionID {
            locationManager.stopMonitoring(for: region)
        }
    }

    private func reconcileMonitoring(_ settings: VargaGeofenceSettings) {
        guard settings.enabled, locationManager.authorizationStatus == .authorizedAlways,
              CLLocationManager.isMonitoringAvailable(for: CLCircularRegion.self) else {
            stopOurRegion()
            return
        }
        let limit = locationManager.maximumRegionMonitoringDistance
        guard limit > 0 else { return }
        let radius = min(settings.radius, limit)
        let center = CLLocationCoordinate2D(latitude: settings.latitude, longitude: settings.longitude)
        let region = CLCircularRegion(center: center, radius: radius, identifier: regionID)
        region.notifyOnEntry = true
        region.notifyOnExit = false
        // Same identifier updates a pre-existing region after a configuration change.
        locationManager.startMonitoring(for: region)
    }

    private func minutes(_ value: String) -> Int? {
        let parts = value.split(separator: ":")
        guard parts.count == 2, parts[0].count == 2, parts[1].count == 2,
              let h = Int(parts[0]), let m = Int(parts[1]),
              (0...23).contains(h), (0...59).contains(m) else { return nil }
        return h * 60 + m
    }

    private func validWindow(_ start: String, _ end: String) -> Bool {
        guard let a = minutes(start), let b = minutes(end) else { return false }
        return a < b
    }

    // Returns a reminder type ONLY when the region is entered on a selected
    // weekday and within the configured local-time window.
    private func eligibleReminder(_ config: VargaGeofenceSettings, now: Date) -> String? {
        let calendar = Calendar.current
        let weekday = calendar.component(.weekday, from: now) - 1
        guard config.days.contains(weekday) else { return nil }
        let current = calendar.component(.hour, from: now) * 60 + calendar.component(.minute, from: now)
        if let from = minutes(config.arrivalFrom), let to = minutes(config.arrivalTo),
           current >= from && current <= to { return "arrival" }
        if let from = minutes(config.returnFrom), let to = minutes(config.returnTo),
           current >= from && current <= to { return "return" }
        return nil
    }

    func locationManagerDidChangeAuthorization(_ manager: CLLocationManager) {
        if let settings = readSettings() { reconcileMonitoring(settings) }
        else { stopOurRegion() }
    }

    func locationManager(_ manager: CLLocationManager, didStartMonitoringFor region: CLRegion) {
        guard region.identifier == regionID else { return }
        // Newly enabled monitoring does not guarantee an entry callback if
        // the device was already inside when monitoring started.
        manager.requestState(for: region)
    }

    func locationManager(_ manager: CLLocationManager, didDetermineState state: CLRegionState, for region: CLRegion) {
        guard region.identifier == regionID, state == .inside else { return }
        // Reuse the identical weekday/window/deduplication checks.
        locationManager(manager, didEnterRegion: region)
    }

    func locationManager(_ manager: CLLocationManager, didEnterRegion region: CLRegion) {
        guard region.identifier == regionID,
              manager.authorizationStatus == .authorizedAlways,
              let settings = readSettings(), settings.enabled,
              let which = eligibleReminder(settings, now: Date()) else { return }

        let now = Date()
        let formatter = DateFormatter()
        formatter.calendar = Calendar(identifier: .gregorian)
        formatter.locale = Locale(identifier: "en_US_POSIX")
        formatter.timeZone = .current
        formatter.dateFormat = "yyyy-MM-dd"
        let keyForDay = "varga.geofence.sent.\(which).\(formatter.string(from: now))"
        guard !defaults.bool(forKey: keyForDay) else { return }
        defaults.set(true, forKey: keyForDay)
        defaults.set(["which": which, "time": now.timeIntervalSince1970], forKey: pendingKey)

        let content = UNMutableNotificationContent()
        content.title = which == "arrival" ? "🔔 Sei al magazzino: timbra ENTRATA" : "🔔 Sei al magazzino: timbra USCITA"
        content.body = "Varga Cantieri ha rilevato l'ingresso nell'area GPS durante il tuo orario. Tocca per aprire e raggiungere Easy-Box."
        content.sound = .default
        content.userInfo = ["feature": "varga-timbratura", "which": which, "destination": "easy-box"]
        content.threadIdentifier = "varga-geofence-timbratura"
        let request = UNNotificationRequest(identifier: "varga-geofence-\(which)-\(formatter.string(from: now))",
                                            content: content,
                                            trigger: UNTimeIntervalNotificationTrigger(timeInterval: 1, repeats: false))
        UNUserNotificationCenter.current().add(request) { error in
            if let error = error {
                NSLog("Varga geofence: local notification failed: %@", String(describing: error))
                // Preserve the pending event so it can be shown on the next open.
            }
        }
    }

    func locationManager(_ manager: CLLocationManager, monitoringDidFailFor region: CLRegion?, withError error: Error) {
        NSLog("Varga geofence region monitoring error: %@", String(describing: error))
    }
}
