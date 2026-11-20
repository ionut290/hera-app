#!/usr/bin/env ruby
# Register Google Sign-In in a freshly generated Capacitor 7 iOS Podfile.
# This is NOT a Google secret and never changes the production Firebase project.
# Called after 'npx cap add ios' and before 'npx cap sync ios' (pod install).
root = File.expand_path("..", __dir__)
podfile = ARGV.fetch(0, File.join(root, "ios", "App", "Podfile"))
abort "iOS Podfile missing" unless File.file?(podfile)
source = File.read(podfile)
pod = "pod 'CapacitorFirebaseAuthentication/Google', :path => '../../node_modules/@capacitor-firebase/authentication'"
marker = "  # Add your Pods here"
abort "Unrecognized generated Capacitor App target" unless
  source.include?("target 'App' do") && source.include?("  capacitor_pods") && source.include?(marker)
# Idempotent: repeated local build commands do not add a second Google pod.
unless source.include?(pod)
  source = source.sub(marker, marker + "\n  " + pod)
end
# Google and Firebase CocoaPods may include resource bundles with their own
# signing settings. Disable signing ONLY for bundle targets, not the App target.
signing_marker = "# VARGA_GOOGLE_BUNDLE_SIGNING_GUARD"
unless source.include?(signing_marker)
  install_marker = "post_install do |installer|"
  abort "Missing generated CocoaPods post_install block" unless source.include?(install_marker)
  code = <<~RUBY
      #{signing_marker}
      installer.pods_project.targets.each do |target|
        next unless target.respond_to?(:product_type) && target.product_type == 'com.apple.product-type.bundle'
        target.build_configurations.each do |config|
          config.build_settings['CODE_SIGNING_ALLOWED'] = 'NO'
        end
      end
  RUBY
  source = source.sub(install_marker, install_marker + "\n" + code.lines.map { |l| "  " + l }.join.rstrip)
end
File.write(podfile, source)
verified = File.read(podfile)
abort "Google pod must be declared exactly once" unless verified.scan(pod).length == 1
abort "CocoaPods bundle signing protection missing" unless verified.include?(signing_marker)
puts "iOS native Google Sign-In pod and resource-bundle signing settings installed."
