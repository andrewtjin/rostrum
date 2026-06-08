// Generate the Settings ribbon GEAR icons — assets/gear-16.png, gear-32.png, gear-80.png.
//
// Deterministic, dependency-free (Node built-ins only: `zlib` for PNG deflate + a tiny CRC32), so the
// gear is reproducible in CI and on any machine with no image toolchain. It draws an 8-tooth gear (a
// donut with a center hole) in a neutral dark gray — the universally understood "settings" glyph — with
// 4× supersampled anti-aliasing so the 16px size stays legible. Office ribbon icons must be raster PNG at
// 16/32/80 (SVG isn't supported for ribbon buttons), which is why this exists at all.
//
// Run once after changing the gear: `node scripts/gen-icons.mjs`, then commit the PNGs. (They're served
// by the dev server + copied to dist by copy-webpack-plugin, so prod/Pages picks them up automatically.)
import { deflateSync } from "zlib";
import { writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const ASSETS = resolve(dirname(fileURLToPath(import.meta.url)), "..", "assets");

// ── minimal PNG encoder (RGBA, 8-bit, no interlace) ────────────────────────────
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([len, body, crc]);
}
function encodePng(size, rgba) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type: RGBA
  // 10/11/12 = compression/filter/interlace = 0
  const stride = 1 + size * 4;
  const raw = Buffer.alloc(size * stride);
  for (let y = 0; y < size; y++) {
    raw[y * stride] = 0; // filter type 0 (None)
    rgba.copy(raw, y * stride + 1, y * size * 4, (y + 1) * size * 4);
  }
  return Buffer.concat([sig, chunk("IHDR", ihdr), chunk("IDAT", deflateSync(raw, { level: 9 })), chunk("IEND", Buffer.alloc(0))]);
}

// ── gear renderer (supersampled coverage → alpha) ──────────────────────────────
const TEETH = 8;
const COLOR = [0x44, 0x44, 0x44]; // neutral dark gray — the standard settings gear
function renderGear(size) {
  const ss = 4; // supersample factor
  const S = size * ss;
  const c = S / 2;
  const Rtip = 0.45 * S; // outer (tooth) radius
  const Rroot = 0.35 * S; // valley radius between teeth
  const hole = 0.16 * S; // center bore (transparent)
  const out = new Uint8Array(size * size * 4);
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let cov = 0;
      for (let sy = 0; sy < ss; sy++) {
        for (let sx = 0; sx < ss; sx++) {
          const dx = x * ss + sx + 0.5 - c;
          const dy = y * ss + sy + 0.5 - c;
          const rad = Math.hypot(dx, dy);
          // Tooth profile: the outer limit alternates between tip and root around the circle, giving TEETH
          // rounded teeth (cos has TEETH positive arcs over a full turn).
          const limit = Math.cos(Math.atan2(dy, dx) * TEETH) > 0 ? Rtip : Rroot;
          if (rad <= limit && rad >= hole) cov++;
        }
      }
      const di = (y * size + x) * 4;
      out[di] = COLOR[0];
      out[di + 1] = COLOR[1];
      out[di + 2] = COLOR[2];
      out[di + 3] = Math.round((cov / (ss * ss)) * 255);
    }
  }
  return Buffer.from(out);
}

for (const size of [16, 32, 80]) {
  const png = encodePng(size, renderGear(size));
  const file = resolve(ASSETS, `gear-${size}.png`);
  writeFileSync(file, png);
  // eslint-disable-next-line no-console
  console.log(`wrote ${file} (${png.length} bytes)`);
}
