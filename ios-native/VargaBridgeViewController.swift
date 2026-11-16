import Capacitor
import UIKit

@objc(VargaBridgeViewController)
class VargaBridgeViewController: CAPBridgeViewController {
    override open func capacitorDidLoad() {
        super.capacitorDidLoad()
        bridge?.registerPluginInstance(VargaGeofencePlugin())
    }
}
