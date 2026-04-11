import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'it.vargacantieri.hera',
  appName: 'Hera App',
  webDir: '.',
  server: {
    androidScheme: 'https'
  }
};

export default config;
