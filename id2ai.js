#!/usr/bin/env node

const { execSync } = require("child_process");
const path = require("path");
const fs = require("fs");
const os = require("os");

const args = process.argv.slice(2);

if (args.length === 0 || args.includes("--help") || args.includes("-h")) {
  console.log(`
Usage: id2ai <input.indd|input.idml> [output.ai]

Converts an InDesign file (.indd or .idml) to Illustrator (.ai).
Exports via EPS per-page, then combines in Illustrator.

Requires Adobe InDesign and Illustrator to be installed.

Options:
  --debug           Keep intermediate files
  --help, -h        Show this help
`);
  process.exit(0);
}

const inputFile = path.resolve(args[0]);
const debug = args.includes("--debug");

if (!fs.existsSync(inputFile)) {
  console.error(`Error: File not found: ${inputFile}`);
  process.exit(1);
}

const ext = path.extname(inputFile).toLowerCase();
if (ext !== ".indd" && ext !== ".idml") {
  console.error(`Error: Input must be an .indd or .idml file`);
  process.exit(1);
}

// Determine output path
let outputFile;
if (args[1] && !args[1].startsWith("--")) {
  outputFile = path.resolve(args[1]);
} else {
  outputFile = inputFile.replace(/\.(indd|idml)$/i, ".ai");
}

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "id2ai-"));
const manifestFile = path.join(tmpDir, "manifest.json");

console.log(`Converting: ${path.basename(inputFile)}`);
console.log(`Output:     ${path.basename(outputFile)}`);
console.log("");

// ---- STEP 1: Open in InDesign, gather info & export per-page EPS ----
console.log("[1/3] Opening in InDesign & exporting pages as EPS...");

const indesignScript = `
var inputPath = "${jsEsc(inputFile)}";
var tmpDir = "${jsEsc(tmpDir)}";
var manifestPath = "${jsEsc(manifestFile)}";

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var doc = app.open(new File(inputPath));

// Gather page info
var pages = [];
for (var i = 0; i < doc.pages.length; i++) {
  var pg = doc.pages[i];
  var b = pg.bounds;
  pages.push({
    index: i,
    name: pg.name,
    width: b[3] - b[1],
    height: b[2] - b[0]
  });
}

// Write manifest
var f = new File(manifestPath);
f.open("w");
f.encoding = "UTF-8";
var json = '{"pageCount":' + doc.pages.length + ',"pages":[';
for (var i = 0; i < pages.length; i++) {
  if (i > 0) json += ",";
  json += '{"index":' + pages[i].index + ',"name":"' + pages[i].name + '","width":' + pages[i].width + ',"height":' + pages[i].height + '}';
}
json += ']}';
f.write(json);
f.close();

// Export each page as EPS
for (var i = 0; i < doc.pages.length; i++) {
  var epsPath = tmpDir + "/page_" + i + ".eps";

  // Set page range to just this page
  app.epsExportPreferences.pageRange = doc.pages[i].name;
  app.epsExportPreferences.epsColor = EPSColorSpace.CMYK;
  app.epsExportPreferences.preview = PreviewTypes.TIFF_PREVIEW;
  app.epsExportPreferences.postscriptLevel = PostScriptLevels.LEVEL_3;
  app.epsExportPreferences.fontEmbedding = FontEmbedding.COMPLETE;
  app.epsExportPreferences.dataFormat = DataFormat.BINARY;

  // Key: set flattener to high quality for gradient fidelity
  try {
    var flattener = app.flattenerPresets.itemByName("[High Resolution]");
    if (flattener.isValid) {
      app.epsExportPreferences.appliedFlattenerPreset = flattener;
    }
  } catch(e) {}

  doc.exportFile(ExportFormat.EPS_TYPE, new File(epsPath));
  $.writeln("Exported page " + doc.pages[i].name + " -> " + epsPath);
}

doc.close(SaveOptions.NO);
`;

runInDesign(indesignScript);

if (!fs.existsSync(manifestFile)) {
  console.error("Error: InDesign failed - no manifest created.");
  cleanup();
  process.exit(1);
}

const manifest = JSON.parse(fs.readFileSync(manifestFile, "utf8"));
console.log(`   Pages: ${manifest.pageCount}`);
manifest.pages.forEach((p) => {
  const epsPath = path.join(tmpDir, `page_${p.index}.eps`);
  const exists = fs.existsSync(epsPath);
  const size = exists ? (fs.statSync(epsPath).size / 1024).toFixed(0) + "KB" : "MISSING";
  console.log(`   Page ${p.name}: ${p.width.toFixed(0)}x${p.height.toFixed(0)}pt (${size})`);
});
console.log("");

// ---- STEP 2: Open first EPS in Illustrator, place remaining pages ----
console.log("[2/3] Building Illustrator document from EPS pages...");

const firstEps = path.join(tmpDir, "page_0.eps");
if (!fs.existsSync(firstEps)) {
  console.error("Error: First page EPS not found.");
  cleanup();
  process.exit(1);
}

// Script to open first page and add remaining as artboards
const buildScript = path.join(tmpDir, "ai_build.jsx");

let aiScript = `
// Open first page
var firstEps = new File("${jsEsc(firstEps)}");
var doc = app.open(firstEps);
doc.artboards[0].name = "Page ${manifest.pages[0].name}";

`;

// Layout pages in a grid (3 columns) to stay within Illustrator canvas limits
const COLS = 3;
const GAP = 50;

for (let i = 1; i < manifest.pages.length; i++) {
  const pg = manifest.pages[i];
  const col = i % COLS;
  const row = Math.floor(i / COLS);
  const ox = col * (manifest.pages[0].width + GAP);
  const oy = -(row * (manifest.pages[0].height + GAP));

  const epsPath = path.join(tmpDir, `page_${i}.eps`);
  if (!fs.existsSync(epsPath)) continue;

  aiScript += `
// Page ${pg.name} (grid pos: col ${col}, row ${row})
var ab${i} = doc.artboards.add([${ox}, ${oy}, ${ox + pg.width}, ${oy - pg.height}]);
ab${i}.name = "Page ${pg.name}";

// Open page EPS in separate doc, copy all, paste into main doc
var pageSrc = app.open(new File("${jsEsc(epsPath)}"));
pageSrc.selectObjectsOnActiveArtboard();
if (pageSrc.selection.length > 0) {
  app.copy();
  pageSrc.close(SaveOptions.DONOTSAVECHANGES);

  app.activeDocument = doc;
  doc.artboards.setActiveArtboardIndex(doc.artboards.length - 1);
  app.paste();

  if (doc.selection.length > 0) {
    var sel = doc.selection;
    var minX = Infinity, maxY = -Infinity;
    for (var s = 0; s < sel.length; s++) {
      if (sel[s].left < minX) minX = sel[s].left;
      if (sel[s].top > maxY) maxY = sel[s].top;
    }
    var dx = ${ox} - minX;
    var dy = ${oy} - maxY;
    for (var s = 0; s < sel.length; s++) {
      sel[s].left += dx;
      sel[s].top += dy;
    }
  }
} else {
  pageSrc.close(SaveOptions.DONOTSAVECHANGES);
}

`;
}

// Save as AI
aiScript += `
// Save as Illustrator
var saveOpts = new IllustratorSaveOptions();
saveOpts.compatibility = Compatibility.ILLUSTRATOR17;
saveOpts.flattenOutput = OutputFlattening.PRESERVEAPPEARANCE;
saveOpts.pdfCompatible = true;
saveOpts.embedICCProfile = true;
doc.saveAs(new File("${jsEsc(outputFile)}"), saveOpts);
`;

fs.writeFileSync(buildScript, aiScript, "utf8");

if (debug) {
  console.log(`   Build script: ${buildScript}`);
}

runIllustrator(buildScript);

console.log("[3/3] Saved!");
console.log("");

if (!fs.existsSync(outputFile)) {
  console.error("Error: Output file not created.");
  cleanup();
  process.exit(1);
}

if (debug) {
  console.log(`   Debug files kept in: ${tmpDir}`);
} else {
  cleanup();
}

const inputSize = (fs.statSync(inputFile).size / 1024).toFixed(0);
const outputSize = (fs.statSync(outputFile).size / 1024 / 1024).toFixed(1);
console.log(`Done! ${inputSize}KB -> ${outputSize}MB`);
console.log(`Output: ${outputFile}`);

// ---- Helper Functions ----

function runInDesign(script) {
  const scriptFile = path.join(tmpDir, "indesign_script.jsx");
  fs.writeFileSync(scriptFile, script, "utf8");

  const appNames = [
    "Adobe InDesign 2026",
    "Adobe InDesign 2025",
    "Adobe InDesign 2024",
  ];
  let lastErr;

  for (const appName of appNames) {
    try {
      execSync(
        `osascript -l JavaScript -e '
          var id = Application("${appName}");
          id.activate();
          id.doScript(Path("${scriptFile.replace(/"/g, '\\"')}"), {language: "javascript"});
        '`,
        {
          stdio: "pipe",
          timeout: 300000,
        }
      );
      return;
    } catch (err) {
      const errMsg = err.stderr?.toString() || err.message || "";
      if (!errMsg.includes("can't be found") && !errMsg.includes("not found")) {
        console.error(`Error in InDesign script: ${errMsg}`);
        cleanup();
        process.exit(1);
      }
      lastErr = err;
      continue;
    }
  }

  console.error("Error: Failed to run script in InDesign.");
  console.error(lastErr?.stderr?.toString() || lastErr?.message);
  cleanup();
  process.exit(1);
}

function runIllustrator(scriptPath) {
  const appNames = [
    "Adobe Illustrator",
    "Adobe Illustrator 2026",
    "Adobe Illustrator 2025",
    "Adobe Illustrator 2024",
  ];
  let lastErr;

  for (const appName of appNames) {
    try {
      execSync(
        `osascript -l JavaScript -e '
          var ai = Application("${appName}");
          ai.activate();
          ai.doJavascript("$.evalFile(\\x27${jsEsc(scriptPath)}\\x27);");
        '`,
        {
          stdio: "pipe",
          timeout: 600000,
        }
      );
      return;
    } catch (err) {
      const errMsg = err.stderr?.toString() || err.message || "";
      if (!errMsg.includes("can't be found") && !errMsg.includes("not found")) {
        console.error(`Error in Illustrator script: ${errMsg}`);
        cleanup();
        process.exit(1);
      }
      lastErr = err;
      continue;
    }
  }

  console.error("Error: Failed to run script in Illustrator.");
  console.error(lastErr?.stderr?.toString() || lastErr?.message);
  cleanup();
  process.exit(1);
}

function jsEsc(str) {
  return str.replace(/\\/g, "/").replace(/"/g, '\\"');
}

function cleanup() {
  if (!debug) {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {}
  }
}
