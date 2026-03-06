import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(scriptDir, '..');
const packageJsonPath = path.join(projectRoot, 'package.json');
const manifestPath = path.join(projectRoot, 'src', 'manifest.json');

async function readJson(filePath) {
  const raw = await readFile(filePath, 'utf8');
  return JSON.parse(raw);
}

async function main() {
  const packageJson = await readJson(packageJsonPath);
  const manifest = await readJson(manifestPath);

  const packageVersion = String(packageJson?.version || '').trim();
  if (!packageVersion) {
    throw new Error('package.json version is missing');
  }

  if (String(manifest?.version || '').trim() === packageVersion) {
    console.log(`[sync:version] versions already aligned (${packageVersion})`);
    return;
  }

  manifest.version = packageVersion;
  await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
  console.log(`[sync:version] updated src/manifest.json -> ${packageVersion}`);
}

await main();
