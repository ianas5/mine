// Drizzle on Expo: .sql migration files are bundled as source (DATABASE §1).
const { getDefaultConfig } = require('expo/metro-config');

const config = getDefaultConfig(__dirname);
config.resolver.sourceExts.push('sql');
// expo-sqlite web driver ships a wasm binary.
config.resolver.assetExts.push('wasm');

module.exports = config;
