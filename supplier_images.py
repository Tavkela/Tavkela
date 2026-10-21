"""მომწოდებლების სურათების ჩამოწერა ოფიციალური საიტებიდან Excel-ის არტიკულების მიხედვით.

ბრენდი ავტომატურად ირჩევა Excel-ის სვეტიდან „ფირმა“ / „ბრენდი“ (ან --brand).
ახალი Shopify-ბრენდის დასამატებლად საკმარისია BRANDS-ში ერთი ჩანაწერი.

გამოყენება:
    python3 supplier_images.py FILE.xlsx --out OUT_DIR [--brand lenz] [--size 1500] [--odoo https://ODOO_DOMAIN]

შედეგი OUT_DIR-ში:
    სურათები/<კოდი>.jpg, <კოდი>_A.jpg, ...   (პირველი: მთავარი სურათი)
    <ბრენდი> - შედეგი.xlsx                    (🟢 ნაპოვნია / 🟡 შესამოწმებელი / 🔴 ვერ ვიპოვე / 🔵 Odoo-ში უკვე არის)

Excel-ში უკვე მწვანედ მონიშნული სტრიქონები (სურათი უკვე აქვს) გამოიტოვება.
"""
import argparse
import io
import json
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote

import openpyxl
import requests
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/141.0 Safari/537.36"}

FILLS = {
    "green": PatternFill("solid", fgColor="FFC6EFCE", bgColor="FFC6EFCE"),
    "yellow": PatternFill("solid", fgColor="FFFFEB9C", bgColor="FFFFEB9C"),
    "red": PatternFill("solid", fgColor="FFFFC7CE", bgColor="FFFFC7CE"),
    "blue": PatternFill("solid", fgColor="FFBDD7EE", bgColor="FFBDD7EE"),
}
ICONS = {"green": "🟢", "yellow": "🟡", "red": "🔴", "blue": "🔵", "skip": "⚪"}

# Excel-ის სვეტების შესაძლო სახელები
COLS = {
    "code": ["კოდი"],
    "article": ["მოდელის კოდი", "მომწ. კოდი", "მომწოდებლის კოდი"],
    "model": ["მოდელი", "სახელი"],
    "color": ["ფერი"],
    "brand": ["ფირმა", "ბრენდი"],
}

# ინგლისური ფერი (Excel) -> გერმანული (საიტი)
COLOR_DE = {
    "BLACK": "schwarz", "WHITE": "weiss", "GREY": "grau", "GRAY": "grau",
    "LIGHTGREY": "hellgrau", "LIGHTGRAY": "hellgrau", "DARKGREY": "dunkelgrau",
    "BLUE": "blau", "NAVY": "marine", "RED": "rot", "PINK": "pink", "ROSE": "rosa",
    "GREEN": "gruen", "OLIVE": "oliv", "LIME": "lime", "YELLOW": "gelb",
    "ORANGE": "orange", "BROWN": "braun", "BEIGE": "beige", "ANTHRACITE": "anthrazit",
    "PURPLE": "violett", "VIOLET": "violett", "TURQUOISE": "tuerkis", "BERRY": "beere",
    "PETROL": "petrol", "SILVER": "silber",
}


# ---------- ბრენდები ----------

def casco_base(s):
    m = re.search(r"(\d{2})\.(\d{4})", str(s or "").replace(",", "."))
    return f"{m.group(1)}.{m.group(2)}" if m else None


class Casco:
    """casco-helmets.com: SKU 07.1902.L — '07.1902' = მოდელი+ფერი, ბოლო = ზომა."""
    name, site = "Casco", "https://casco-helmets.com"
    check_color = False  # ფერი არტიკულშივეა

    @staticmethod
    def article(excel_value):
        return casco_base(excel_value)

    @staticmethod
    def split_sku(sku, product=None):
        a = casco_base(sku)
        return (a, a) if a else (None, None)

    @staticmethod
    def image_key(fname, article):
        m = re.match(r"(\d{2}\.\d{4})_", fname)
        return m.group(1) if m else None

    @staticmethod
    def handle_article(handle):
        m = re.search(r"(\d{2})-(\d{4})(?:-|$)", handle)
        return f"{m.group(1)}.{m.group(2)}" if m else None


class Lenz:
    """lenzproducts.com: SKU 090_64 — '090' = მოდელი, '6' = ფერი, '4' = ზომა.
    სურათის სახელში ფერი ასე წერია: ...-090-60-1.webp"""
    name, site = "Lenz", "https://lenzproducts.com"
    check_color = True

    @staticmethod
    def norm(a):
        a = str(a or "").strip().upper()
        return a.lstrip("0") or a

    @classmethod
    def article(cls, excel_value):
        return cls.norm(excel_value) if excel_value not in (None, "") else None

    @classmethod
    def split_sku(cls, sku, product=None):
        if "_" not in (sku or ""):
            return None, None
        art, rest = sku.split("_", 1)
        return cls.norm(art), rest[:-1]

    @classmethod
    def image_key(cls, fname, article):
        m = re.search(rf"(?<![0-9a-z])0*{re.escape(article.lower())}[-_](\d+)0(?=[-_.])",
                      fname.lower())
        return m.group(1) if m else None

    @staticmethod
    def handle_article(handle):
        return None


class Shred:
    """shredoptics.com: Excel-ის „მომწ. კოდი“ = ვარიანტის სრული SKU (GOGRAO21A, HENOTP12L).
    სურათის სახელი ფერის კოდით იწყება: GOGRAO21A_21.jpg, HENOTP12_11.jpg (ზომის გარეშე)."""
    name, site = "Shred", "https://shredoptics.com"
    check_color = False  # SKU-ში ფერიცაა და ზომაც
    _prefixes = {}

    @staticmethod
    def article(excel_value):
        return str(excel_value).strip().upper() if excel_value not in (None, "") else None

    @classmethod
    def image_key(cls, fname, article=None):
        m = re.match(r"([A-Z]{2}[A-Z0-9]{4,})_", fname)
        return m.group(1) if m else None

    @classmethod
    def split_sku(cls, sku, product=None):
        sku = str(sku or "").strip().upper()
        if not sku:
            return None, None
        key = sku
        if product is not None:
            if product["id"] not in cls._prefixes:
                cls._prefixes[product["id"]] = {
                    k for k in (cls.image_key(image_name(i["src"])) for i in product.get("images", [])) if k}
            pref = [p for p in cls._prefixes[product["id"]] if sku.startswith(p)]
            if pref:
                key = max(pref, key=len)
        return sku, key

    @staticmethod
    def handle_article(handle):
        return None


BRANDS = {"casco": Casco, "lenz": Lenz, "shred": Shred}


# ---------- ქსელი ----------

def get(url, session, **kw):
    for attempt in range(6):
        try:
            r = session.get(url, headers=UA, timeout=60, **kw)
        except requests.ConnectionError:
            if attempt == 5:
                raise
            time.sleep(2 ** attempt)
            continue
        if r.status_code in (429, 502, 503):
            time.sleep(2 ** attempt + 1)
            continue
        r.raise_for_status()
        return r
    r.raise_for_status()
    return r


def load_catalog(brand, session, cache):
    if cache.exists():
        return json.loads(cache.read_text(encoding="utf-8"))
    products, page = [], 1
    while True:
        batch = get(f"{brand.site}/products.json?limit=250&page={page}", session).json()["products"]
        if not batch:
            break
        products += batch
        print(f"  კატალოგი: გვერდი {page}, სულ {len(products)} პროდუქტი")
        page += 1
        time.sleep(0.5)
    cache.write_text(json.dumps(products, ensure_ascii=False), encoding="utf-8")
    return products


# ---------- შედარება ----------

def image_name(src):
    return src.split("?")[0].rsplit("/", 1)[-1]


def color_index(product):
    for i, o in enumerate(product.get("options", [])):
        if "farb" in o["name"].lower() or "color" in o["name"].lower():
            return i
    return None


def build_groups(brand, products):
    """არტიკული -> [{product, key, color, variant_ids}] (ერთი ჩანაწერი = პროდუქტის ერთი ფერი)."""
    index = {}
    for p in products:
        ci = color_index(p)
        groups = {}
        for v in p.get("variants", []):
            art, key = brand.split_sku(v.get("sku"), p)
            if not art:
                continue
            g = groups.setdefault((art, key), {
                "product": p, "article": art, "key": key, "variant_ids": set(),
                "color": v.get(f"option{ci + 1}") if ci is not None else "",
            })
            g["variant_ids"].add(v["id"])
        if not groups and brand.handle_article(p.get("handle", "")):
            art = brand.handle_article(p["handle"])
            groups[(art, art)] = {"product": p, "article": art, "key": art,
                                  "variant_ids": set(), "color": ""}
        for g in groups.values():
            index.setdefault(g["article"], []).append(g)
    return index


def images_for(brand, group):
    """ამ ფერის სურათები. ფერი დგინდება: ვარიანტთან მიბმით, სახელში კოდით,
    ან (თუ სახელში კოდი არ წერია) წინა ფერიანი სურათის მიხედვით."""
    p = group["product"]
    imgs = sorted(p.get("images", []), key=lambda i: i.get("position", 0))
    keys_of = {}
    for v in p.get("variants", []):
        art, key = brand.split_sku(v.get("sku"), p)
        if art:
            keys_of[v["id"]] = key
    all_keys = set(keys_of.values())
    if len(all_keys) <= 1:
        return [i["src"] for i in imgs]
    out, current = [], None
    for i in imgs:
        linked = {keys_of[x] for x in i.get("variant_ids") or [] if x in keys_of}
        named = brand.image_key(image_name(i["src"]), group["article"])
        if linked:
            current = group["key"] if group["key"] in linked else next(iter(linked))
            mine = group["key"] in linked
        elif named is not None:
            if named not in all_keys:
                continue  # სხვა პროდუქტის ფერი
            current = named
            mine = named == group["key"]
        else:
            mine = current == group["key"]
        if mine and i["src"] not in out:
            out.append(i["src"])
    return out


def norm_de(s):
    s = str(s).lower().replace("ß", "ss").replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
    return s


def color_tokens_excel(color):
    out = []
    for part in str(color or "").upper().split("/"):
        k = re.sub(r"[^A-Z]", "", part)
        if k:
            out.append(COLOR_DE.get(k, k.lower()))
    return out


def color_tokens_site(color):
    return [norm_de(t).strip() for t in str(color or "").split("/") if t.strip()]


def color_score(excel_color, site_color):
    a, b = color_tokens_excel(excel_color), color_tokens_site(site_color)
    if not a or not b:
        return 0.0
    if sorted(a) == sorted(b):
        return 1.0
    common = len(set(a) & set(b))
    return common / max(len(set(a)), len(set(b)))


def tokens(s):
    s = re.sub(r"[^A-Z0-9]+", " ", str(s).upper().replace("+", " PLUS "))
    return [re.sub(r"(.)\1+", r"\1", t) for t in s.split()
            if t not in {"ML", "MAGNET", "LINK", "CASCO", "LENZ"}]


def title_score(model, title):
    want, have = set(tokens(model)), tokens(title)
    return sum(t in want for t in have) - 2 * sum(t not in want for t in have)


def name_candidates(model, products, limit=3):
    want = tokens(model)
    found = []
    for p in products:
        title = tokens(p.get("title", ""))
        have = set(title + tokens(p.get("handle", "")))
        if want and title[:1] == want[:1] and all(t in have for t in want):
            found.append(p)
    return found[:limit]


def suffix(i):
    """0 -> '', 1 -> '_A', 26 -> '_Z', 27 -> '_AA'"""
    s = ""
    while i > 0:
        i -= 1
        s = chr(65 + i % 26) + s
        i //= 26
    return "_" + s if s else ""


def is_marked_green(row):
    for c in row:
        f = c.fill
        if f is not None and f.fill_type == "solid" and f.fgColor is not None and f.fgColor.type == "rgb":
            rgb = str(f.fgColor.rgb).upper()[-6:]
            r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
            if g > 120 and g > r + 40 and g > b + 40:
                return True
    return False


# ---------- Odoo-ს საჯარო საიტი ----------

def odoo_has(session, odoo, code):
    r = get(f"{odoo}/shop?search={quote(str(code))}", session)
    if "No results" in r.text:
        return False
    return bool(re.search(r'href="(?:/[a-z]{2}(?:_[A-Z]{2})?)?/shop/(?!cart|checkout|category|page|wishlist)[^"?#]+-\d+',
                          r.text))


# ---------- სურათები ----------

def fetch_jpeg(session, src, size):
    url = "https:" + src if src.startswith("//") else src
    if size:
        url += ("&" if "?" in url else "?") + f"width={size}"
    im = Image.open(io.BytesIO(get(url, session).content))
    im.load()
    if size and max(im.size) > size:
        im.thumbnail((size, size), Image.LANCZOS)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=92, optimize=True)
    return buf.getvalue()


# ---------- მთავარი ----------

def find_header(ws):
    for row in ws.iter_rows(min_row=1, max_row=15):
        names = {str(c.value).strip(): c.column for c in row if c.value is not None}
        cols = {k: next((names[n] for n in v if n in names), None) for k, v in COLS.items()}
        if cols["code"] and cols["article"]:
            return row[0].row, cols
    sys.exit("სათაურის სტრიქონი ვერ ვიპოვე (საჭიროა 'კოდი' და 'მოდელის კოდი' / 'მომწ. კოდი')")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel")
    ap.add_argument("--out", required=True)
    ap.add_argument("--brand", choices=sorted(BRANDS))
    ap.add_argument("--size", type=int, default=1500, help="მაქს. ზომა პიქსელებში; 0 = ორიგინალი")
    ap.add_argument("--odoo", default=None)
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.excel)
    ws = wb.active
    header_row, cols = find_header(ws)

    brand_key = args.brand
    if not brand_key and cols["brand"]:
        for r in range(header_row + 1, ws.max_row + 1):
            v = str(ws.cell(r, cols["brand"]).value or "").strip().lower()
            if v in BRANDS:
                brand_key = v
                break
    if not brand_key:
        sys.exit(f"ბრენდი ვერ დავადგინე; მიუთითეთ --brand ({', '.join(sorted(BRANDS))})")
    brand = BRANDS[brand_key]

    out = Path(args.out)
    img_dir = out / "სურათები"
    img_dir.mkdir(parents=True, exist_ok=True)
    s = requests.Session()

    c_status = ws.max_column + 1
    bold = ws.cell(header_row, 1).font.copy(bold=True)
    for i, title in enumerate(["სტატუსი", "საიტის ბმული", "სურათები"]):
        ws.cell(header_row, c_status + i, title).font = bold

    print(f"{brand.name}: კატალოგის ჩამოტვირთვა...")
    products = load_catalog(brand, s, out / f"{brand_key}_products.json")
    index = build_groups(brand, products)
    print(f"  სულ {len(products)} პროდუქტი, {len(index)} არტიკული")

    cache = {}
    stats = {k: 0 for k in ICONS}
    for r in range(header_row + 1, ws.max_row + 1):
        row = ws[r]
        code = ws.cell(r, cols["code"]).value
        art_raw = ws.cell(r, cols["article"]).value
        if code in (None, "") or art_raw in (None, ""):
            continue
        code = str(code).strip()
        model = ws.cell(r, cols["model"]).value if cols["model"] else ""
        color = ws.cell(r, cols["color"]).value if cols["color"] else ""
        link, n_img, kind = "", 0, None

        if is_marked_green(row):
            kind, status = "skip", "უკვე აქვს სურათი (მწვანედ იყო მონიშნული) — გამოტოვდა"
        elif args.odoo and odoo_has(s, args.odoo, code):
            kind, status = "blue", "Odoo-ს საიტზე უკვე დევს — სურათები არ ჩამოიწერა"
        else:
            art = brand.article(art_raw)
            groups = index.get(art, [])
            chosen, sure = None, True
            if groups and brand.check_color and color:
                scored = sorted(((color_score(color, g["color"]), g) for g in groups),
                                key=lambda t: -t[0])
                best, chosen = scored[0]
                if best < 1.0:
                    sure = False
            elif groups:
                # ერთი კოდი რამდენიმე პროდუქტშია (მაგ. GRATIFY და GRATIFY+): ვირჩევთ სახელით
                chosen = max(groups, key=lambda g: title_score(model, g["product"]["title"]))

            if chosen and sure:
                p = chosen["product"]
                link = f"{brand.site}/products/{p['handle']}"
                srcs = images_for(brand, chosen)
                if not srcs:
                    kind, status = "yellow", f"საიტზე არის, მაგრამ სურათი არ აქვს — {p['title']}"
                else:
                    ck = (p["id"], chosen["key"])
                    if ck not in cache:
                        cache[ck] = [fetch_jpeg(s, u, args.size) for u in srcs]
                    for i, data in enumerate(cache[ck]):
                        (img_dir / f"{code}{suffix(i)}.jpg").write_bytes(data)
                    n_img = len(cache[ck])
                    kind = "green"
                    status = f"ნაპოვნია: {p['title']}" + (f" — {chosen['color']}" if chosen["color"] else "")
            elif chosen:
                p = chosen["product"]
                link = f"{brand.site}/products/{p['handle']}"
                kind = "yellow"
                status = (f"შესამოწმებელი: მოდელი საიტზე არის, მაგრამ ფერი '{color}' ზუსტად არ ემთხვევა. "
                          "საიტზე არის: " + "; ".join(g["color"] for g in groups))
            else:
                cands = name_candidates(model, products)
                if cands:
                    kind = "yellow"
                    status = ("შესამოწმებელი: ეს მოდელი საიტზე არის, მაგრამ ამ არტიკულით არა. "
                              "მსგავსი: " + "; ".join(p["title"] for p in cands))
                    link = f"{brand.site}/products/{cands[0]['handle']}"
                else:
                    kind, status = "red", "საიტზე ვერ ვიპოვე (ძველი მოდელი / აღარ იყიდება)"

        stats[kind] += 1
        if kind in FILLS:
            for c in range(1, c_status + 3):
                ws.cell(r, c).fill = FILLS[kind]
        ws.cell(r, c_status, f"{ICONS[kind]} {status}")
        if link:
            ws.cell(r, c_status + 1, link).hyperlink = link
        ws.cell(r, c_status + 2, n_img or None)
        print(f"  [{kind:6}] {code} {art_raw} {model} {color}: {status}")

    ws.column_dimensions[get_column_letter(c_status)].width = 60
    ws.column_dimensions[get_column_letter(c_status + 1)].width = 45
    wb.save(out / f"{brand.name} - შედეგი.xlsx")
    print("\nშედეგი:", stats, "| სურათები:", len(list(img_dir.glob("*.jpg"))))


if __name__ == "__main__":
    main()
