#!/usr/bin/env ruby
# Configure manual Apple Distribution signing for ONLY the generated Capacitor
# iOS App target in Release. Never touch Pods, Debug, app JS or production data.
require "xcodeproj"

root = File.expand_path("..", __dir__)
project_path = File.join(root, "ios", "App", "App.xcodeproj")
abort "Generated App.xcodeproj not found" unless File.directory?(project_path)
team = ENV.fetch("IOS_APPLE_TEAM_ID", "")
profile = ENV.fetch("IOS_PROFILE_NAME", "")
abort "Missing valid IOS_APPLE_TEAM_ID" unless team.match?(/\A[A-Za-z0-9]{10}\z/)
abort "Missing App Store Connect provisioning profile name" if profile.strip.empty?

project = Xcodeproj::Project.open(project_path)
app_targets = project.targets.select { |t| t.name == "App" && t.respond_to?(:build_configurations) }
abort "Expected exactly one App target" unless app_targets.size == 1
app = app_targets.first
release = app.build_configurations.find { |config| config.name == "Release" }
abort "App Release configuration missing" unless release

# An Xcode command-line CODE_SIGN_IDENTITY=Apple Distribution override leaks into
# every CocoaPods target, triggering conflicting-provisioning errors.
# Instead bind a real distribution certificate + App Store provisioning profile
# to the App Release target alone. The profile must already be installed.
settings = release.build_settings
settings["DEVELOPMENT_TEAM"] = team
settings["CODE_SIGN_STYLE"] = "Manual"
settings["CODE_SIGN_IDENTITY"] = "Apple Distribution"
settings["CODE_SIGN_IDENTITY[sdk=iphoneos*]"] = "Apple Distribution"
settings["PROVISIONING_PROFILE_SPECIFIER"] = profile
settings.delete("PROVISIONING_PROFILE")

target_attributes = project.root_object.attributes["TargetAttributes"] ||= {}
target_attributes[app.uuid] ||= {}
target_attributes[app.uuid]["ProvisioningStyle"] = "Manual"

project.save
verified = Xcodeproj::Project.open(project_path)
target = verified.targets.find { |t| t.name == "App" }
config = target.build_configurations.find { |c| c.name == "Release" }
abort "App Distribution profile signing not persisted" unless
  config.build_settings["CODE_SIGN_IDENTITY[sdk=iphoneos*]"] == "Apple Distribution" &&
  config.build_settings["DEVELOPMENT_TEAM"] == team &&
  config.build_settings["CODE_SIGN_STYLE"] == "Manual" &&
  config.build_settings["PROVISIONING_PROFILE_SPECIFIER"] == profile

puts "Configured App Release for manual Apple Distribution; Pods and Debug unchanged."
