import plistlib
import unittest
from prepare_v2_export import options as local_options
from prepare_v2_internal_upload import options

class InternalUploadOptionsTests(unittest.TestCase):
    def test_denies_without_explicit_enable(self):
        for value in (None, "", "false", "True", True):
            with self.subTest(value=value), self.assertRaises(ValueError):
                options(value)
    def test_internal_only_exact_v2(self):
        opts = plistlib.loads(plistlib.dumps(options("true")))
        self.assertEqual(opts["destination"], "upload")
        self.assertIs(opts["testFlightInternalTestingOnly"], True)
        self.assertEqual(opts["method"], "app-store-connect")
        self.assertEqual(opts["provisioningProfiles"], {"it.vargacantieri.staging": "Varga Cantieri 2 Staging TestFlight"})
        self.assertEqual(opts["teamID"], "NJ64236Q54")
        self.assertIs(opts["manageAppVersionAndBuildNumber"], False)
    def test_local_export_unchanged(self):
        options("true")
        self.assertEqual(local_options()["destination"], "export")

if __name__ == "__main__":
    unittest.main()
