const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const releaseDir = path.join(root, 'release');
const arch = process.arch;
const mainName = 'AhMen Booking Manager';
const iconPath = path.join(root, 'icons', 'ahmen.icns');

const run = (cmd) => {
  execSync(cmd, { stdio: 'inherit', cwd: root });
};

run('npm run build');
run(`./node_modules/.bin/electron-packager . "${mainName}" --platform=darwin --arch=${arch} --out=release --icon="${iconPath}" --overwrite`);
const mainApp = path.join(releaseDir, `${mainName}-darwin-${arch}`, `${mainName}.app`);

if (!fs.existsSync(mainApp)) {
  throw new Error(`Main app not found at ${mainApp}`);
}
