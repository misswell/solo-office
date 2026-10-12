import { execFileSync } from "node:child_process";
import { cpSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

/** Assembles dist/Solo Office.app from the Electron already installed and the built app/ dir. */

const root = dirname(dirname(fileURLToPath(import.meta.url)));
// An isolated output lets packaging coexist with a previously launched dist app.
const output = resolve(root, process.env.SOLO_PACKAGE_DIR || "dist");
const bundle = join(output, "Solo Office.app");
// Package a verified snapshot without rebuilding another agent's active sources.
const appSource = resolve(root, process.env.SOLO_APP_DIR || "app");
for (const file of ["main.cjs", "preload.cjs", "renderer.js", "index.html"]) {
  if (!existsSync(join(appSource, file))) throw new Error(`Missing built app file: ${join(appSource, file)}`);
}
const { version } = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
const resources = join(bundle, "Contents", "Resources");

rmSync(bundle, { recursive: true, force: true });
mkdirSync(output, { recursive: true });
// ditto, not cp: the framework keeps symlinks that a plain recursive copy flattens.
execFileSync("ditto", [join(root, "node_modules", "electron", "dist", "Electron.app"), bundle]);

rmSync(join(resources, "default_app.asar"), { force: true });
mkdirSync(join(resources, "app"), { recursive: true });
cpSync(appSource, join(resources, "app", "app"), { recursive: true });
cpSync(join(root, "package.json"), join(resources, "app", "package.json"));
cpSync(join(root, "build", "icon.icns"), join(resources, "electron.icns"));

const plist = join(bundle, "Contents", "Info.plist");
const fields = {
  CFBundleName: "Solo Office",
  CFBundleDisplayName: "Solo Office",
  CFBundleIdentifier: "com.solo-office.app",
  CFBundleShortVersionString: version,
  CFBundleVersion: version,
  CFBundleExecutable: "Solo Office",
};
for (const [key, value] of Object.entries(fields)) execFileSync("plutil", ["-replace", key, "-string", value, plist]);
renameSync(join(bundle, "Contents", "MacOS", "Electron"), join(bundle, "Contents", "MacOS", "Solo Office"));

// npm Electron's linker-signed framework lacks a complete resource seal. Sign
// the copied nested code too; never modify the source Electron installation.
execFileSync("codesign", ["--force", "--deep", "--sign", "-", "--timestamp=none", bundle]);
execFileSync("codesign", ["--verify", "--deep", "--strict", bundle]);
console.log(bundle);
