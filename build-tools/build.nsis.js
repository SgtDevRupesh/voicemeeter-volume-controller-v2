#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const buildDir = __dirname;
const installerDir = path.join(buildDir, "..");
const logDir = path.join(installerDir, "dev-logs");
const logFile = path.join(logDir, `build-nsis-${new Date().toISOString().replace(/[:.]/g, "-")}.log`);
const publishDir = path.join(installerDir, "artifacts", "nsis-publish");

if (!fs.existsSync(logDir)) fs.mkdirSync(logDir, { recursive: true });

const logStream = fs.createWriteStream(logFile, { flags: "a" });

function log(msg) {
  const line = `[${new Date().toISOString()}] ${msg}`;
  console.log(line);
  logStream.write(line + "\n");
}

function die(msg) {
  const line = `[${new Date().toISOString()}] ERROR: ${msg}`;
  console.error(line);
  logStream.write(line + "\n");
  logStream.write("--- BUILD FAILED ---\n");
  logStream.end();
  process.exit(1);
}

log("=== NSIS Build Started ===");
log("Log file: " + logFile);


const payloadDir = path.join(buildDir, "payload");

log("Preparing payload...");

const projectPath = path.join(installerDir, "VMWV.Modern", "VMWV.Modern.csproj");

log("Publishing fresh application binaries...");
try {
  fs.rmSync(publishDir, { recursive: true, force: true });
  fs.mkdirSync(publishDir, { recursive: true });
  execSync(
    `dotnet publish "${projectPath}" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -p:PublishTrimmed=false -o "${publishDir}"`,
    { stdio: "inherit" }
  );
} catch (e) {
  die("Failed to publish app for installer payload: " + e.message);
}

const appSrcDir = publishDir;
if (!fs.existsSync(path.join(appSrcDir, "VVC.exe"))) {
  die("Published app is missing VVC.exe.");
}

// Create payload.zip from the app binaries
const payloadZip = path.join(buildDir, "payload.zip");
log("Creating payload.zip from: " + appSrcDir);
try {
  if (fs.existsSync(payloadZip)) fs.unlinkSync(payloadZip);
  execSync(
    `powershell -NoProfile -Command "Compress-Archive -Path '${appSrcDir}\\*' -DestinationPath '${payloadZip}' -Force"`,
    { stdio: "inherit" }
  );
} catch (e) {
  die("Failed to create payload.zip: " + e.message);
}
log("payload.zip created.");

log("Payload ready.");

// Check for makensis

// Check for makensis
try {
  execSync("makensis.exe /VERSION", { stdio: "ignore" });
} catch {
  die("NSIS (makensis.exe) not found on PATH. Install from: https://nsis.sourceforge.io/");
}

// Build NSIS installer
log("Compiling installer with makensis...");
const nsiFile = path.join(buildDir, "installer.nsi");
const outDir = path.join(installerDir, "release");

if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

try {
  const output = execSync(`makensis.exe /DOUTDIR="${outDir}" "${nsiFile}"`, { encoding: "utf8", stderr: "pipe" });
  if (output) log(output);
  const exePath = path.join(outDir, "VVC-Installer.exe");
  log("✓ Installer created: " + exePath);
  log("=== BUILD SUCCESS ===");
  logStream.end();
} catch (e) {
  die("makensis failed:\n" + (e.stdout || "") + "\n" + (e.stderr || "") + "\n" + e.message);
}
