import plistlib
import unittest
from prepare_v2_export import options

class ExportOptionsTests(unittest.TestCase):
    def test_exact_v2_only(self):
        opts = plistlib.loads(plistlib.dumps(options()))
        self.assertEqual(opts["provisioningProfiles"], {"it.vargacantieri.staging": "Varga Cantieri 2 Staging TestFlight"})
        self.assertEqual(opts["teamID"], "NJ64236Q54")
        self.assertEqual(opts["signingStyle"], "manual")
    def test_never_upload(self):
        opts = options()
        self.assertEqual(opts["destination"], "export")
        self.assertIs(opts["testFlightInternalTestingOnly"], True)
        self.assertIs(opts["manageAppVersionAndBuildNumber"], False)
        self.assertEqual(opts["method"], "app-store-connect")

if __name__ == "__main__":
    unittest.main()
