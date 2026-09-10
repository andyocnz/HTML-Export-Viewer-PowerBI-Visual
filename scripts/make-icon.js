const { PNG } = require("pngjs");
const fs = require("fs");
const path = require("path");

const SIZE = 20;
const png = new PNG({ width: SIZE, height: SIZE });

const GREEN = [74, 94, 58, 255];      // brand green (#4a5e3a)
const WHITE = [255, 255, 255, 255];
const FOLD = [197, 217, 160, 255];    // brand light green (#c5d9a0) - page fold
const ROW = [197, 217, 160, 255];     // table row lines
const ARROW_BG = [230, 126, 34, 255]; // orange badge
const TRANSPARENT = [0, 0, 0, 0];

function setPx(x, y, color) {
    if (x < 0 || y < 0 || x >= SIZE || y >= SIZE) {
        return;
    }
    const idx = (SIZE * y + x) << 2;
    png.data[idx] = color[0];
    png.data[idx + 1] = color[1];
    png.data[idx + 2] = color[2];
    png.data[idx + 3] = color[3];
}

function fillRect(x0, y0, x1, y1, color) {
    for (let y = y0; y <= y1; y++) {
        for (let x = x0; x <= x1; x++) {
            setPx(x, y, color);
        }
    }
}

// start fully transparent
fillRect(0, 0, SIZE - 1, SIZE - 1, TRANSPARENT);

// rounded background square (corners chamfered by skipping 1px)
fillRect(1, 0, 18, 19, GREEN);
fillRect(0, 1, 19, 18, GREEN);
// chamfer corners
setPx(0, 0, TRANSPARENT); setPx(19, 0, TRANSPARENT);
setPx(0, 19, TRANSPARENT); setPx(19, 19, TRANSPARENT);
setPx(1, 1, GREEN); setPx(18, 1, GREEN);
setPx(1, 18, GREEN); setPx(18, 18, GREEN);

// document/page (white) with folded top-right corner
fillRect(4, 3, 12, 15, WHITE);
// fold triangle (top-right corner of page)
setPx(11, 3, FOLD); setPx(12, 3, FOLD);
setPx(12, 4, FOLD);

// table rows inside the page
fillRect(6, 6, 10, 6, ROW);
fillRect(6, 9, 10, 9, ROW);
fillRect(6, 12, 10, 12, ROW);

// export badge (bottom-right), circle-ish square with downward arrow
fillRect(11, 11, 18, 18, ARROW_BG);
setPx(11, 11, TRANSPARENT); setPx(18, 11, TRANSPARENT);
setPx(11, 18, TRANSPARENT); setPx(18, 18, TRANSPARENT);

// arrow: vertical stem + arrow head pointing down, in white
fillRect(14, 12, 15, 14, WHITE);
setPx(13, 15, WHITE); setPx(16, 15, WHITE);
fillRect(13, 15, 16, 15, WHITE);
setPx(14, 16, WHITE); setPx(15, 16, WHITE);

// outPath is built solely from __dirname + fixed segments (no external/user input); this is
// a build-time script, not part of the shipped visual bundle.
const outPath = path.join(__dirname, "..", "assets", "icon.png");
// eslint-disable-next-line powerbi-visuals/non-literal-fs-path
png.pack().pipe(fs.createWriteStream(outPath)).on("finish", () => {
    console.log("Wrote", outPath);
});
