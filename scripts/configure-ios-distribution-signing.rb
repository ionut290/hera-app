#!/usr/bin/env ruby
# Configure the generated Capacitor App target for distribution archives.
# Never change the CocoaPods project or application JavaScript.
require "xcodeproj"

root = File.expand_path("..", __dir__)
project_path = File.join(root, "ios", "App", "App.xcodeproj")
abort "Generated App.xcodeproj not found" unless File.directory?(project_path)
team = ENV.fetch("IOS_APPLE_TEAM_ID", "")
abort "Missing valid IOS_APPLE_TEAM_ID" unless team.match?(/\A[A-Za-z0-9]{10}\z/)

project = Xcodeproj::Project.open(project_path)
app_targets = project.targets.select { |t| t.name == "App" && t.respond_to?(:build_configurations) }
abort "Expected exactly one App target" unless app_targets.size == 1
app = app_targets.first
release = app.build_configurations.find { |config| config.name == "Release" }
abort "App Release configuration missing" unless release

# Scope the identity to the app target's Release settings. Passing
# CODE_SIGN_IDENTITY on the xcodebuild command line affects ALL workspace
# targets, including Pods, and causes conflicting provisioning errors.
settings = release.build_settings
settings["DEVELOPMENT_TEAM"] = team
settings["CODE_SIGN_STYLE"] = "Automatic"
settings["CODE_SIGN_IDENTITY"] = "Apple Distribution"
settings["CODE_SIGN_IDENTITY[sdk=iphoneos*]"] = "Apple Distribution"
settings.delete("PROVISIONING_PROFILE_SPECIFIER")
settings.delete("PROVISIONING_PROFILE")

project.save
verified = Xcodeproj::Project.open(project_path)
config = verified.targets.find { |t| t.name == "App" }.build_configurations.find { |c| c.name == "Release" }
abort "Distribution signing target not persisted" unless
  config.build_settings["CODE_SIGN_IDENTITY[sdk=iphoneos*]"] == "Apple Distribution" &&
  config.build_settings["DEVELOPMENT_TEAM"] == team

puts "Configured only App Release for Apple Distribution; Pods and Debug unchanged."
