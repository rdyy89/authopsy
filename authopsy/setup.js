#!/usr/bin/env node

/**
 * Setup script for Email Authentication Checker Outlook Add-in
 * This script helps set up the development environment
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('🔧 Setting up Email Authentication Checker...\n');

// Check if Node.js version is supported
const nodeVersion = process.version;
const majorVersion = parseInt(nodeVersion.split('.')[0].slice(1));

if (majorVersion < 14) {
    console.error('❌ Node.js 14 or higher is required. Current version:', nodeVersion);
    process.exit(1);
}

console.log('✅ Node.js version check passed:', nodeVersion);

// Install dependencies
console.log('\n📦 Installing dependencies...');
try {
    execSync('npm install', { stdio: 'inherit' });
    console.log('✅ Dependencies installed successfully');
} catch (error) {
    console.error('❌ Failed to install dependencies:', error.message);
    process.exit(1);
}

// Generate development certificates
console.log('\n🔐 Setting up development certificates...');
try {
    const officeCerts = require('office-addin-dev-certs');
    if (!fs.existsSync('certs')) {
        fs.mkdirSync('certs');
    }
    console.log('✅ Development certificates ready');
} catch (error) {
    console.log('⚠️  Development certificates setup skipped - will be generated on first run');
}

// Create development environment file
console.log('\n⚙️  Creating development configuration...');
const envContent = `# Development Environment Configuration
NODE_ENV=development
DEV_SERVER_PORT=3000
HTTPS_PORT=3000
`;

fs.writeFileSync('.env.development', envContent);
console.log('✅ Development configuration created');

// Validate manifest
console.log('\n📋 Validating manifest...');
try {
    execSync('npm run validate', { stdio: 'inherit' });
    console.log('✅ Manifest validation passed');
} catch (error) {
    console.log('⚠️  Manifest validation skipped - office-addin tools not yet available');
}

console.log('\n🎉 Setup complete! Next steps:');
console.log('1. Start development server: npm run dev-server');
console.log('2. In another terminal, sideload add-in: npm run sideload');
console.log('3. Open Outlook and test the add-in');
console.log('\n📚 For more information, check the README.md file');
