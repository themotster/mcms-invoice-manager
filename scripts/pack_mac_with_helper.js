const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const releaseDir = path.join(root, 'release');
const arch = process.arch;
const mainName = 'AhMen Booking Manager';
const helperName = 'AhMen Reminders';
const iconPath = path.join(root, 'icons', 'ahmen.icns');

const run = (cmd) => {
  execSync(cmd, { stdio: 'inherit', cwd: root });
};

run('npm run build');
run(`./node_modules/.bin/electron-packager . "${mainName}" --platform=darwin --arch=${arch} --out=release --icon="${iconPath}" --overwrite`);
run(`./node_modules/.bin/electron-packager . "${helperName}" --platform=darwin --arch=${arch} --out=release --icon="${iconPath}" --overwrite`);

const mainApp = path.join(releaseDir, `${mainName}-darwin-${arch}`, `${mainName}.app`);
const helperApp = path.join(releaseDir, `${helperName}-darwin-${arch}`, `${helperName}.app`);

if (!fs.existsSync(mainApp)) {
  throw new Error(`Main app not found at ${mainApp}`);
}
if (!fs.existsSync(helperApp)) {
  throw new Error(`Helper app not found at ${helperApp}`);
}

const loginItemsDir = path.join(mainApp, 'Contents', 'Library', 'LoginItems');
fs.mkdirSync(loginItemsDir, { recursive: true });
const target = path.join(loginItemsDir, `${helperName}.app`);

if (fs.existsSync(target)) {
  fs.rmSync(target, { recursive: true, force: true });
}

fs.cpSync(helperApp, target, { recursive: true });
console.log(`Embedded helper app at ${target}`);
