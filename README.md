# indesign-to-illustrator

Convert Adobe InDesign files (`.indd`) to Illustrator (`.ai`) from the command line.

## The problem

There's no "Save As Illustrator" in InDesign. Never has been. If you need to hand off an InDesign layout to someone who only has Illustrator — or you need editable vectors from a multi-page `.indd` file — you're stuck with a surprisingly annoying workflow.

We tried a few things before landing on something that actually works:

**Clipboard copy-paste** — Script InDesign to select all objects on a page, copy, switch to Illustrator, paste. Sounds reasonable. In practice? Empty artboards. The macOS clipboard doesn't reliably transfer complex InDesign objects between apps via scripting.

**PDF intermediate** — Export from InDesign as PDF, open in Illustrator. Gets you *something*, but gradients come out banded and stepped. InDesign's gradient feather effects get decomposed into discrete strips during PDF conversion. A brand book full of smooth gradient overlays turns into a blocky mess.

**EPS per-page** — This is what actually works. Export each page from InDesign as a PostScript Level 3 EPS with high-resolution flattening. Open the first page in Illustrator, then copy-paste the remaining pages into new artboards. Gradients come through clean. Text stays live. Vectors stay vectors.

That's what this tool does.

## Requirements

- macOS or Windows
- Node.js >= 14
- Adobe InDesign (2024, 2025, or 2026)
- Adobe Illustrator (2024, 2025, or 2026)

On macOS it uses `osascript` (JXA) for app automation. On Windows it uses COM automation via `cscript`.

## Install

```bash
npm install -g indesign-to-illustrator
```

Or clone and link:

```bash
git clone https://github.com/aedev-tools/indesign-to-illustrator.git
cd indesign-to-illustrator
npm link
```

## Usage

```bash
id2ai myfile.indd
```

Output goes next to the input file as `myfile.ai`.

Custom output path:

```bash
id2ai myfile.indd /path/to/output.ai
```

Works with `.idml` files too:

```bash
id2ai myfile.idml
```

### Options

```
--debug    Keep intermediate EPS files (useful for troubleshooting)
--help     Show help
```

## What it does

1. Opens the `.indd` file in InDesign
2. Exports each page as a high-quality EPS (PostScript Level 3, binary, complete font embedding, high-res flattener)
3. Opens the first page EPS in Illustrator
4. Copies remaining pages into the same document as separate artboards (arranged in a grid)
5. Saves as `.ai`

Pages are laid out in a 3-column grid to stay within Illustrator's canvas limits.

## What it preserves

- Live editable text
- Vector paths and shapes
- Gradient fills (the whole reason this exists)
- CMYK color space
- Embedded fonts

## Limitations

- macOS and Windows only (no Linux — Adobe apps don't run on Linux)
- Both InDesign and Illustrator must be installed and licensed
- Transparency effects get flattened (EPS limitation) — but at high resolution, so it looks good
- Multi-page documents get combined into one Illustrator file with multiple artboards

## License

MIT
