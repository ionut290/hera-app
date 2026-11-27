import datetime
import unittest
from validate_ios_profile import validate, TEAM, BUNDLE, PROFILE

class ProfileTests(unittest.TestCase):
    def setUp(self):
        self.now = datetime.datetime(2026, 11, 26)
        self.p = {"Name": PROFILE, "TeamIdentifier": [TEAM], "ExpirationDate": datetime.datetime(2027, 11, 11), "UUID": "12345678-1234-1234-1234-123456789012", "DeveloperCertificates": [b"test-leaf"], "Entitlements": {"application-identifier": TEAM+"."+BUNDLE, "get-task-allow": False}}
    def test_valid(self):
        self.assertEqual(validate(self.p, b"test-leaf", self.now), self.p["UUID"])
    def test_original_bundle_denied(self):
        self.p["Entitlements"]["application-identifier"] = TEAM+".it.vargacantieri.hera"
        with self.assertRaises(ValueError): validate(self.p,b"test-leaf",self.now)
    def test_wrong_cert_denied(self):
        with self.assertRaises(ValueError): validate(self.p,b"other",self.now)
    def test_wrong_team_denied(self):
        self.p["TeamIdentifier"]=["OTHER"]
        with self.assertRaises(ValueError): validate(self.p,b"test-leaf",self.now)
    def test_development_denied(self):
        self.p["Entitlements"]["get-task-allow"]=True
        with self.assertRaises(ValueError): validate(self.p,b"test-leaf",self.now)
    def test_adhoc_denied(self):
        self.p["ProvisionedDevices"]=["test"]
        with self.assertRaises(ValueError): validate(self.p,b"test-leaf",self.now)
    def test_expired_denied(self):
        with self.assertRaises(ValueError): validate(self.p,b"test-leaf",datetime.datetime(2028,1,1))
    def test_wrong_profile_denied(self):
        self.p["Name"]="VARGA CANTIERI - TestFlight"
        with self.assertRaises(ValueError): validate(self.p,b"test-leaf",self.now)

if __name__ == "__main__":
    unittest.main()
