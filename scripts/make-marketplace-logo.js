const { PNG } = require("pngjs");
const fs = require("fs");
const path = require("path");

// Same design as make-icon.js (the 20x20 pbiviz icon), redrawn on a 20x20 grid of
// scaled-up blocks so it stays perfectly crisp at marketplace logo sizes - these are all
// flat rectangles, so integer scaling has no blur/aliasing to worry about.
const UNIT = 20;
const SCALE = 15; // 20 * 15 = 300
const SIZE = UNIT * SCALE;

const GREEN = [74, 94, 58, 255];
const WHITE = [255, 255, 255, 255];
const FOLD = [197, 217, 160, 255];
const ROW = [197, 217, 160, 255];
const ARROW_BG = [230, 126, 34, 255];
const TRANSPARENT = [0, 0, 0, 0];

const png = new PNG({ width: SIZE, height: SIZE });

function setUnitPx(x, y, color) {
    if (x < 0 || y < 0 || x >= UNIT || y >= UNIT) {
        return;
    }
    for (let dy = 0; dy < SCALE; dy++) {
        for (let dx = 0; dx < SCALE; dx++) {
            const px = x * SCALE + dx;
            const py = y * SCALE + dy;
            const idx = (SIZE * py + px) << 2;
            png.data[idx] = color[0];
            png.data[idx + 1] = color[1];
            png.data[idx + 2] = color[2];
            png.data[idx + 3] = color[3];
        }
    }
}

function fillRect(x0, y0, x1, y1, color) {
    for (let y = y0; y <= y1; y++) {
        for (let x = x0; x <= x1; x++) {
            setUnitPx(x, y, color);
        }
    }
}

fillRect(0, 0, UNIT - 1, UNIT - 1, TRANSPARENT);

fillRect(1, 0, 18, 19, GREEN);
fillRect(0, 1, 19, 18, GREEN);
setUnitPx(0, 0, TRANSPARENT); setUnitPx(19, 0, TRANSPARENT);
setUnitPx(0, 19, TRANSPARENT); setUnitPx(19, 19, TRANSPARENT);
setUnitPx(1, 1, GREEN); setUnitPx(18, 1, GREEN);
setUnitPx(1, 18, GREEN); setUnitPx(18, 18, GREEN);

fillRect(4, 3, 12, 15, WHITE);
setUnitPx(11, 3, FOLD); setUnitPx(12, 3, FOLD);
setUnitPx(12, 4, FOLD);

fillRect(6, 6, 10, 6, ROW);
fillRect(6, 9, 10, 9, ROW);
fillRect(6, 12, 10, 12, ROW);

fillRect(11, 11, 18, 18, ARROW_BG);
setUnitPx(11, 11, TRANSPARENT); setUnitPx(18, 11, TRANSPARENT);
setUnitPx(11, 18, TRANSPARENT); setUnitPx(18, 18, TRANSPARENT);

fillRect(14, 12, 15, 14, WHITE);
setUnitPx(13, 15, WHITE); setUnitPx(16, 15, WHITE);
fillRect(13, 15, 16, 15, WHITE);
setUnitPx(14, 16, WHITE); setUnitPx(15, 16, WHITE);

// outDir/outPath are built solely from __dirname + fixed segments (no external/user input);
// this is a build-time script, not part of the shipped visual bundle.
const outDir = path.join(__dirname, "..", "assets", "marketplace");
// eslint-disable-next-line powerbi-visuals/non-literal-fs-path
if (!fs.existsSync(outDir)) {
    // eslint-disable-next-line powerbi-visuals/non-literal-fs-path
    fs.mkdirSync(outDir, { recursive: true });
}
const outPath = path.join(outDir, `logo-${SIZE}.png`);
// eslint-disable-next-line powerbi-visuals/non-literal-fs-path
png.pack().pipe(fs.createWriteStream(outPath)).on("finish", () => {
    console.log("Wrote", outPath);
});
