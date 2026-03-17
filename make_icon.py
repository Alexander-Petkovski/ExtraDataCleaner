"""
make_icon.py  —  BUILD-TIME ONLY
Generates icon.ico for ExtraDataCleaner.
Run automatically by build_exe.bat before PyInstaller.

Tries Pillow first (full colour brush icon).
Falls back to pure-Python BMP if Pillow is unavailable.
"""

from pathlib import Path
import struct
import zlib

OUT = Path(__file__).parent / "icon.ico"


# ── attempt full-colour icon with Pillow ──────────────────────────────────────

def _pillow_icon():
    from PIL import Image, ImageDraw

    def draw(size: int) -> "Image.Image":
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d   = ImageDraw.Draw(img)
        k   = size / 64.0

        def sc(v):        return int(round(v * k))
        def poly(pts, **kw): d.polygon([(sc(x), sc(y)) for x, y in pts], **kw)
        def ln(x0,y0,x1,y1,w=1,**kw):
            d.line([(sc(x0),sc(y0)),(sc(x1),sc(y1))],
                   width=max(1,int(round(w*k))),**kw)
        def ell(x0,y0,x1,y1,**kw):
            d.ellipse([sc(x0),sc(y0),sc(x1),sc(y1)],**kw)

        # drop shadow
        poly([(10,62),(60,62),(62,64),(8,64)], fill=(0,0,0,50))

        # handle — diagonal wooden rod
        poly([(4,2),(17,2),(46,42),(33,42)],  fill=(192,124,48))   # body
        poly([(5,3),(10,3),(39,42),(34,42)],  fill=(222,165,88))   # light
        poly([(14,2),(17,2),(46,41),(43,41)], fill=(138,82,20))    # dark
        ell(4,0,17,7, fill=(170,104,35))
        poly([(4,4),(17,4),(17,7),(4,7)],     fill=(170,104,35))   # end cap

        # ferrule — chrome collar
        poly([(31,40),(47,40),(52,52),(26,52)], fill=(178,182,198))
        poly([(32,41),(47,41),(48,45),(31,45)], fill=(222,226,240))
        poly([(27,49),(52,49),(52,52),(26,52)], fill=(118,122,140))
        ln(28,47,51,47, w=0.8, fill=(155,158,175))

        # bristle base wedge
        poly([(26,51),(52,51),(63,59),(15,59)], fill=(45,29,9))

        # bristle strands
        strands = [
            (-13,15,62,(58,40,14)), (-10,20,63,(78,55,21)),
            ( -7,25,63,(98,72,30)), ( -4,30,63,(68,47,17)),
            ( -1,35,63,(88,62,24)), (  2,40,62,(78,55,21)),
            (  5,45,62,(98,72,30)), (  8,50,62,(68,47,17)),
            ( 11,55,61,(58,40,14)), ( 14,60,59,(78,55,21)),
            ( 12,63,55,(58,40,14)),
        ]
        bx, by = 39, 52
        for dx,tx,ty,col in strands:
            ln(bx+dx,by,tx,ty, w=1.6, fill=col)
        for dx,tx,ty,col in strands[::3]:
            lighter = tuple(min(255,c+45) for c in col)
            ln(bx+dx-1,by,tx-1,ty, w=0.7, fill=lighter)

        return img

    # Draw at full 256px resolution; Pillow downsamples to each ICO size
    base   = draw(256)
    sizes  = [(16,16), (24,24), (32,32), (48,48), (64,64), (256,256)]
    base.save(str(OUT), format="ICO", sizes=sizes)
    print(f"  icon.ico  written (Pillow, {len(sizes)} sizes)")


# ── pure-Python fallback — simple BMP brush silhouette ───────────────────────

def _make_bmp_32(size: int, pixels) -> bytes:
    """Build a 32-bit ARGB BMP (for embedding in ICO)."""
    w = h = size
    row_bytes = w * 4
    bmp_data  = bytearray()
    # rows are stored bottom-up in BMP
    for y in range(h - 1, -1, -1):
        for x in range(w):
            r, g, b, a = pixels[y * w + x]
            bmp_data += bytes([b, g, r, a])
    # BITMAPINFOHEADER (40 bytes) — height doubled for mask
    hdr = struct.pack("<IiiHHIIiiII",
        40, w, h * 2, 1, 32, 0,
        len(bmp_data), 0, 0, 0, 0)
    # AND mask (all zeros = fully opaque where alpha > 0)
    mask = b'\x00' * (((w + 31) // 32) * 4 * h)
    return bytes(hdr) + bytes(bmp_data) + mask


def _make_pixel_brush(size: int):
    """Draw a minimal brush silhouette in RGBA pixels."""
    px = [(0, 0, 0, 0)] * (size * size)
    k  = size / 32

    def dot(x, y, r, g, b, a=255):
        xi, yi = int(x * k), int(y * k)
        for dy in range(max(1, int(k))):
            for dx in range(max(1, int(k))):
                idx = (yi + dy) * size + (xi + dx)
                if 0 <= idx < len(px):
                    px[idx] = (r, g, b, a)

    # handle (diagonal, brown)
    for t in range(14):
        dot(2 + t, 2 + t * 1.2, 192, 124, 48)
    # ferrule (silver)
    for t in range(6):
        dot(15 + t, 17 + t * 0.3, 178, 182, 198)
    # bristles (dark brown, fan)
    for i, (tx, ty) in enumerate([(14,28),(17,29),(20,30),(23,30),
                                   (26,30),(29,29),(31,27)]):
        dot(18, 19, 45, 29, 9)
        dot(tx, ty, 68 + i * 4, 47, 17)
    return px


def _fallback_icon():
    for size in [16, 32, 48]:
        px  = _make_pixel_brush(size)
        bmp = _make_bmp_32(size, px)

    # Build a minimal multi-size ICO with three BMP images
    sizes = [16, 32, 48]
    bmps  = [_make_bmp_32(s, _make_pixel_brush(s)) for s in sizes]

    # ICO header
    num   = len(sizes)
    ico   = struct.pack("<HHH", 0, 1, num)
    # ICONDIRENTRY per image — offset begins after header + all entries
    offset = 6 + num * 16
    entries = b""
    for i, (s, bmp) in enumerate(zip(sizes, bmps)):
        entries += struct.pack("<BBBBHHII",
            s, s, 0, 0, 1, 32, len(bmp), offset)
        offset += len(bmp)
    OUT.write_bytes(ico + entries + b"".join(bmps))
    print("  icon.ico  written (pure-Python fallback)")


# ── entry point ───────────────────────────────────────────────────────────────

def main():
    try:
        _pillow_icon()
    except ImportError:
        print("  Pillow not available — using pure-Python fallback icon.")
        _fallback_icon()
    except Exception as exc:
        print(f"  Pillow icon failed ({exc}) — using pure-Python fallback icon.")
        _fallback_icon()


if __name__ == "__main__":
    main()
