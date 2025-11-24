# confluence_attachments_from_zip_gui_11_dnd.py
# 08版ベース：DnD(ドラッグ&ドロップ)対応 / 黒地ログ / md・Word切替  (C) Tanukida
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont
from pathlib import Path
import os, re, sys, zipfile, shutil
from datetime import datetime
from bs4 import BeautifulSoup
import unicodedata, urllib.parse
from collections import defaultdict
from functools import lru_cache

APP_TITLE = "Klefki Conflu"
ATT_DIR_NAME = "添付ファイル"
# USE_FILE_URI_FOR_ATTACHMENTS = True # 添付リンクを絶対パス (file:///C:/...) にする
USE_OFFICE_URI_SCHEME = True  # .xlsx/.docx/.pptx をアプリで直接開く (Windows + Office)
HTML_DIR_NAME = "html_pages" # HTML出力先ディレクトリ名
POKEBALL_FILE = "pokeball.png"
EXE_ICON_FILE = "exe_icon.png"
BG_FILE_1 = "BG_01.png"
EMPTY_ICON_FILE = "Empty_icon.png"
FOOTER_GORI_NAME = "footer_gori1.png"
HEADER_IMAGE = "Header_logo_1.png"
RESOURCE_DIR_CANDIDATES = ["resources", "resource", "assets", "img", "_internal"]
LOG_NOT_FOUND_ATTACHMENTS = True   # 見つからなかった添付をログ出力する
NOT_FOUND_LOG_NAME = "not_found_attachments.log"
SIDEBAR_HTML_ROOT: Path | None = None
SIDEBAR_ITEMS = [] 
SIDEBAR_EMPTY_PAGES: set[Path] = set()   # 空白ページの一覧

# --- optional deps ---
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # pip install tkinterdnd2
    HAS_DND = True
except Exception:
    HAS_DND = False
    DND_FILES = None
    TkinterDnD = None

try:
    import magic
    HAS_MAGIC = True
except Exception:
    HAS_MAGIC = False

try:
    import html2text
    HAS_H2T = True
except Exception:
    HAS_H2T = False

try:
    import docx
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False

MIME_TO_EXT = {
    "image/jpeg": ".jpg","image/png": ".png","image/gif": ".gif","image/webp": ".webp","image/tiff": ".tif","image/webp": ".webp",
    "application/pdf": ".pdf","application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.ms-excel": ".xls",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-powerpoint": ".ppt",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
    "text/plain": ".txt","application/zip": ".zip","video/mp4": ".mp4","audio/mpeg": ".mp3",
}

# ---------- utils ----------
EXT_RE = re.compile(r"\s*\.[A-Za-z0-9]{1,6}$")
def strip_any_ext(name: str) -> str:
    if not name: return name
    return EXT_RE.sub("", name.strip())

def sanitize(s: str) -> str:
    s = re.sub(r'[\\/:*?"<>|]', " ", s or ""); return re.sub(r"\s+", " ", s).strip()

def ensure_unique(dst: Path) -> Path:
    if not dst.exists(): return dst
    stem, ext = os.path.splitext(dst.name); i = 2
    while True:
        cand = dst.with_name(f"{stem} ({i}){ext}")
        if not cand.exists(): return cand
        i += 1

def mime_from_bytes(b: bytes) -> str | None:
    if not HAS_MAGIC: return None
    try: return magic.from_buffer(b, mime=True)
    except Exception: return None

def log_append(txt: tk.Text, line: str):
    # 色分け（CMD風）
    lvl = "INFO"
    s = line.lstrip()
    if s.startswith("[ERROR]") or s.startswith("ERROR:") or "エラー" in s:
        lvl = "ERROR"
    elif s.startswith("[WARN]") or s.startswith("WARNING:") or s.startswith("[SKIP]"):
        lvl = "WARN"
    elif s.startswith("[OK]") or s.startswith("[SUMMARY]") or s.startswith("DONE"):
        lvl = "OK"
    elif s.startswith("[MOVE]"):
        lvl = "MOVE"
    elif s.startswith("[INFO]"):
        lvl = "INFO"
    txt.config(state="normal")
    txt.insert("end", line + "\n", (lvl,))
    txt.see("end")
    txt.config(state="disabled")
    pump_gui(txt)

def pump_gui(widget: tk.Widget) -> None:
    # 重い処理中に GUI が真っ白で固まらないように更新するヘルパ
    try:
        # 描画キューとイベントキューを処理する
        widget.update_idletasks()
        widget.update()
    except tk.TclError:
        # ウィンドウを閉じた後などに呼ばれても落ちないように握りつぶす
        pass


def _normalize_filename(name: str) -> str:
    """比較用にファイル名文字列を正規化（URLデコード→NFKC→trim→lower）"""
    if not isinstance(name, str):
        name = str(name or "")
    name = urllib.parse.unquote(name)
    name = unicodedata.normalize("NFKC", name)
    return name.strip().lower()

def _get_attach_dir(out_root: Path) -> Path:
    """添付の実体が入っているディレクトリ"""
    return (out_root / ATT_DIR_NAME).resolve()

def build_attachment_index(out_root: Path) -> dict:
    # out_root/添付ファイル 配下を1回だけ走査し，名前→Path / stem→Paths の索引を作る．
    att_root = (out_root / ATT_DIR_NAME)
    by_name: dict[str, Path] = {}
    by_lower: dict[str, Path] = {}
    by_stem: dict[str, list[Path]] = defaultdict(list)
    if att_root.exists():
        for p in att_root.rglob("*"):
            if p.is_file():
                by_name[p.name] = p
                by_lower[p.name.lower()] = p
                by_stem[p.stem.lower()].append(p)
    return {"by_name": by_name, "by_lower": by_lower, "by_stem": dict(by_stem)}

def _log_not_found_attachment(out_root: Path, wanted: str, context: str = "") -> None:
    if not LOG_NOT_FOUND_ATTACHMENTS:
        return
    try:
        p = out_root / NOT_FOUND_LOG_NAME
        line = f"{datetime.datetime.now().isoformat(timespec='seconds')}  MISS  {wanted}"
        if context:
            line += f"  ({context})"
        with p.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def normalize_all_attachment_filenames(out_root: Path) -> None:
    # 任意：展開済みの添付ファイルを “正規化名” に一括リネーム
    attach_dir = _get_attach_dir(out_root)
    if not attach_dir.exists():
        return
    seen = {}
    for a in sorted(attach_dir.glob("**/*")):
        if not a.is_file():
            continue
        orig = a.name
        norm = _normalize_filename(orig)
        if not norm:
            continue
        # 衝突回避
        stem, ext = os.path.splitext(norm)
        idx = 0
        newname = norm
        while (a.with_name(newname)).exists():
            idx += 1
            newname = f"{stem}_{idx}{ext}"
        if newname != orig:
            try:
                a.rename(a.with_name(newname))
            except Exception:
                pass

def _is_blank_storage_html(storage_html: str) -> bool:
    # Confluence の storage HTML が実質『空』かどうかを判定する
    if not storage_html or not storage_html.strip():
        return True

    try:
        s = BeautifulSoup(storage_html or "", "lxml")
    except Exception:
        s = BeautifulSoup(storage_html or "", "html.parser")

    # タグを除いたテキストが空なら「空白ページ」とみなす
    return s.get_text(strip=True) == ""


# ---------- entities.xml 解析 ----------
def parse_entities(xml_bytes: bytes):
    import xml.etree.ElementTree as ET
    spaces, pages, att_title, att_to_page, att_filename = {}, {}, {}, {}, {}
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return spaces, pages, att_title, att_to_page, att_filename

    def pick(parent, paths):
        for p in paths:
            el = parent.find(p)
            if el is not None and (el.text or "").strip():
                return el.text.strip()
        return ""

    # Space
    for obj in root.findall(".//object[@class='Space']"):
        sid  = pick(obj, ["id[@name='id']", "property[@name='id']"])
        skey = pick(obj, ["property[@name='key']", "property[@name='spaceKey']"])
        if sid and skey:
            spaces[sid] = skey

    # Page
    for obj in root.findall(".//object[@class='Page']"):
        pid   = pick(obj, ["id[@name='id']", "property[@name='id']"])
        title = pick(obj, ["property[@name='title']"]) or (f"page_{pid}" if pid else "")
        spaceId = ""
        space_prop = obj.find("property[@name='space']")
        if space_prop is not None:
            spaceId = pick(space_prop, ["id[@name='id']", "property[@name='id']"])
        parentId = ""
        par_prop = obj.find("property[@name='parent']")
        if par_prop is not None:
            parentId = pick(par_prop, ["id[@name='id']", "property[@name='id']"])
        if pid:
            pages[pid] = {"title": title, "spaceId": spaceId, "parentId": parentId}

    # Attachment
    for obj in root.findall(".//object[@class='Attachment']"):
        aid    = pick(obj, ["id[@name='id']", "property[@name='id']"])
        atitle = pick(obj, ["property[@name='title']"]) or ""
        pageId = ""
        cont = obj.find("property[@name='container']")
        if cont is not None:
            pageId = pick(cont, ["id[@name='id']", "property[@name='id']"])
        if aid:
            if atitle: att_title[aid] = atitle
            if pageId: att_to_page[aid] = pageId

    # AttachmentVersion / AttachmentData -> fileName
    for cls in ("AttachmentVersion", "AttachmentData"):
        for obj in root.findall(f".//object[@class='{cls}']"):
            aid = pick(obj, ["property[@name='attachment']/id[@name='id']",
                            "property[@name='attachment']/property[@name='id']"])
            fname = pick(obj, ["property[@name='fileName']"])
            if aid and fname:
                att_filename[aid] = fname
    return spaces, pages, att_title, att_to_page, att_filename

def _decide_space_key(spaces: dict, pages: dict) -> str:
    if spaces:
        if len(spaces) == 1: return list(spaces.values())[0]
        freq = {}
        for p in pages.values():
            sid = p.get("spaceId", "")
            if sid: freq[sid] = freq.get(sid, 0) + 1
        if freq:
            best_sid = max(freq, key=freq.get)
            return spaces.get(best_sid, "UnknownSpace")
        return list(spaces.values())[0]
    return "UnknownSpace"

def _get_space_key_from_zip(zip_path: Path) -> str:
    try:
        with zipfile.ZipFile(str(zip_path), "r") as zf:
            names = [n for n in zf.namelist() if not n.endswith("/")]
            attach_roots = sorted({ n.split("attachments/")[0] + "attachments/"
                                    for n in names if "attachments/" in n })
            ent_candidates = [n for n in names if n.lower().endswith("entities.xml")]
            entities_name = None
            if attach_roots:
                parent = "/".join(attach_roots[0].strip("/").split("/")[:-1])
                for n in ent_candidates:
                    if n.startswith(parent + "/"):
                        entities_name = n; break
            if not entities_name and ent_candidates:
                entities_name = ent_candidates[0]
            if not entities_name: return "UnknownSpace"
            spaces, pages, *_ = parse_entities(zf.read(entities_name))
            return _decide_space_key(spaces, pages)
    except Exception:
        return "UnknownSpace"

def _get_space_key_from_folder(root: Path) -> str:
    try:
        candidates = sorted([p for p in root.rglob("attachments") if p.is_dir()],
                            key=lambda p: len(str(p)))
        ent_file = None
        if candidates:
            ent = candidates[0].parent / "entities.xml"
            if ent.exists(): ent_file = ent
        if not ent_file:
            ents = list(root.rglob("entities.xml"))
            if ents: ent_file = ents[0]
        if not ent_file: return "UnknownSpace"
        spaces, pages, *_ = parse_entities(ent_file.read_bytes())
        return _decide_space_key(spaces, pages)
    except Exception:
        return "UnknownSpace"

def _build_auto_out_root(input_path: Path, log: tk.Text) -> Path:
    if input_path.is_file() and input_path.suffix.lower() == ".zip":
        space_key = _get_space_key_from_zip(input_path); base_dir = input_path.parent
    else:
        space_key = _get_space_key_from_folder(input_path); base_dir = input_path.parent
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    folder_name = f"{ts}_{sanitize(space_key) or 'UnknownSpace'}"
    out_root = base_dir / folder_name
    log_append(log, f"[OUT] 出力先（自動）: {out_root}")
    return out_root

# 添付リンクの href を作る共通関数
def _file_uri_raw(p: Path) -> str:
    # エンコード無しの file:/// 絶対URI（バックスラッシュ→スラッシュ）
    p = Path(p).resolve()
    return "file:///" + p.as_posix()

def _find_pokeball_image(dst_root: Path) -> Path | None:
    # pokeball.png を以下の優先順で探し，見つかればその Path を返す:
    # 1) 出力先
    here = dst_root / POKEBALL_FILE
    if here.exists():
        return here

    # 2) 環境変数
    env = os.environ.get("POKEBALL_PATH")
    if env:
        p = Path(env)
        if p.exists():
            return p

    # 3) スクリプトのあるフォルダ
    try:
        script_dir = Path(__file__).resolve().parent
    except NameError:
        script_dir = Path.cwd()

    candidates = [script_dir / POKEBALL_FILE]

    # 4) exe/スクリプトの親にあるリソースフォルダ候補
    #   PyInstaller(非onefile) は sys.executable の親が exe の場所
    exe_parent = Path(getattr(sys, "executable", "")) .resolve().parent if getattr(sys, "executable", None) else script_dir
    for base in {exe_parent, script_dir}:
        for folder in RESOURCE_DIR_CANDIDATES:
            candidates.append(base / folder / POKEBALL_FILE)

    # 5) 親直下
    candidates.append(exe_parent / POKEBALL_FILE)

    for c in candidates:
        if c.exists():
            return c
    return None

def _find_bg_image(dst_root: Path) -> Path | None:
    # BG_01.png を以下の優先順で探し，見つかればその Path を返す:
    try:
        script_dir = Path(__file__).resolve().parent
    except NameError:
        script_dir = Path.cwd()

    candidates = [script_dir / BG_FILE_1]

    exe_parent = Path(getattr(sys, "executable", "")) .resolve().parent \
        if getattr(sys, "executable", None) else script_dir
    for base in {exe_parent, script_dir}:
        for folder in RESOURCE_DIR_CANDIDATES:
            candidates.append(base / folder / BG_FILE_1)

    candidates.append(exe_parent / BG_FILE_1)

    for c in candidates:
        if c.exists():
            return c
    return None

def _find_resource_file(filename: str) -> Path | None:
    """
    exe_icon.png など，任意のリソースファイルを探す共通関数．
    - スクリプトと同じフォルダ
    - exe の親フォルダ配下の resources / img / assets / resource / _internal
    などを順番に探す
    """
    try:
        script_dir = Path(__file__).resolve().parent
    except NameError:
        script_dir = Path.cwd()

    candidates: list[Path] = []

    # スクリプトと同じ場所
    candidates.append(script_dir / filename)

    # exe / スクリプトの親フォルダ + リソース候補フォルダ
    exe_parent = Path(getattr(sys, "executable", "")) .resolve().parent \
        if getattr(sys, "executable", None) else script_dir
    for base in {exe_parent, script_dir}:
        for folder in RESOURCE_DIR_CANDIDATES:
            candidates.append(base / folder / filename)

    # 親フォルダ直下
    candidates.append(exe_parent / filename)

    for c in candidates:
        if c.exists():
            return c
    return None

def _ensure_bg_image(out_root: Path) -> Path | None:
    # 背景画像を検索する
    attach_root = _get_attach_dir(out_root)          # out_root / ATT_DIR_NAME
    dst = attach_root / BG_FILE_1
    if dst.exists():
        return dst

    src = _find_bg_image(out_root)
    if not src or not src.exists():
        return None

    try:
        attach_root.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        return dst
    except Exception:
        return None
    
def _ensure_pokeball(out_root: Path) -> Path | None:
    # pokeball.png を out_root/添付ファイル/ に用意して，その Path を返す．
    attach_root = _get_attach_dir(out_root)  # out_root / ATT_DIR_NAME
    dst = attach_root / POKEBALL_FILE

    if not dst.exists():
        src = _find_pokeball_image(out_root)
        if src and src.exists():
            try:
                attach_root.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
            except Exception:
                return None

    return dst if dst.exists() else None

    
def _ensure_exe_icon(out_root: Path) -> Path | None:
    # exe_icon.png を out_root/添付ファイル/ 配下に用意して，その Path を返す．
    attach_root = _get_attach_dir(out_root)  # out_root / ATT_DIR_NAME
    dst = attach_root / EXE_ICON_FILE

    if not dst.exists():
        src = _find_resource_file(EXE_ICON_FILE)
        if src and src.exists():
            try:
                attach_root.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
            except Exception:
                return None

    return dst if dst.exists() else None

def _ensure_empty_icon(out_root: Path) -> Path | None:
    # Empty_icon.png を out_root/添付ファイル/ に用意して，その Path を返す
    attach_root = _get_attach_dir(out_root)
    dst = attach_root / EMPTY_ICON_FILE

    if not dst.exists():
        src = _find_resource_file(EMPTY_ICON_FILE)
        if src and src.exists():
            try:
                attach_root.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
            except Exception:
                return None

    return dst if dst.exists() else None

def _ensure_footer_gori(out_root: Path) -> Path | None:
    # footer_gori1.png を 出力先/添付ファイル/ にコピーし，そのパスを返す
    attach_root = _get_attach_dir(out_root)
    dst = attach_root / FOOTER_GORI_NAME

    if not dst.exists():
        src = _find_resource_file(FOOTER_GORI_NAME)
        if src and src.exists():
            try:
                attach_root.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
            except Exception:
                return None

    return dst if dst.exists() else None

def _ensure_header_logo(out_root: Path) -> Path | None:
    # Header_logo_1.png を 出力先/添付ファイル/ にコピーし，そのパスを返す

    attach_root = _get_attach_dir(out_root)
    dst = attach_root / HEADER_IMAGE

    if not dst.exists():
        src = _find_resource_file(HEADER_IMAGE)
        if src and src.exists():
            try:
                attach_root.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
            except Exception:
                return None

    return dst if dst.exists() else None


# ---------- 添付復元（出力は out_root/添付ファイル 内） ----------
def write_output(rel_under_attachments: Path, data: bytes, attach_root: Path,
                title_for_name: str | None,
                preferred_ext_from_entities: str | None,
                page_id: str | None, att_id: str | None,
                spaces: dict, pages: dict,
                dry_run: bool, log: tk.Text):

    # ==== PageId 階層 → ページ名フォルダ ====
    page_title = ""
    if page_id and pages:
        page_title = pages.get(page_id, {}).get("title", "") or ""
    safe_title = sanitize(page_title) if page_title else "その他"

    # すべて「添付ファイル/<ページ名>/」に集約
    out_parent = attach_root / safe_title
    # ===============================================

    if not dry_run:
        out_parent.mkdir(parents=True, exist_ok=True)

    orig = sanitize(rel_under_attachments.name)
    leaf_ext = Path(orig).suffix.lower()
    stem, _ = os.path.splitext(orig)

    # (A) 「.zipという拡張子のファイル」だけは除外
    if leaf_ext == ".zip" or (preferred_ext_from_entities and preferred_ext_from_entities.lower().endswith(".zip")):
        log_append(log, f"[SKIP] ZIPファイル除外: {rel_under_attachments}")
        return

    mime0 = mime_from_bytes(data)

    # (B) OOXML 判定：中身がZIPでも Office なら通す
    def _detect_ooxml_ext_from_zip_bytes(b: bytes) -> str | None:
        try:
            import io, zipfile
            z = zipfile.ZipFile(io.BytesIO(b))
            names = z.namelist()
            # OOXMLの典型構造
            has_ct = any(n.endswith("[Content_Types].xml") for n in names)
            if not has_ct:
                return None
            # サブフォルダで判定
            if any(n.startswith("word/") for n in names):
                return ".docx"
            if any(n.startswith("xl/") for n in names):
                return ".xlsx"
            if any(n.startswith("ppt/") for n in names):
                return ".pptx"
            # OOXMLの亜種はここに増やせる（.odt 等）
            return None
        except Exception:
            return None

    ooxml_hint = None
    if mime0 == "application/zip":
        ooxml_hint = _detect_ooxml_ext_from_zip_bytes(data)
        # OOXMLと判断できない素のZIPだけ弾く
        if ooxml_hint is None:
            log_append(log, f"[SKIP] ZIP(MIME) 除外: {rel_under_attachments}")
            return

    new_stem = strip_any_ext(sanitize(title_for_name or stem)) or "attachment"

    # (C) 拡張子決定を強化（entities > 実ファイル拡張子 > MIME推定 > OOXML検知）
    preferred_ext = os.path.splitext(preferred_ext_from_entities)[1].lower() if preferred_ext_from_entities else None
    final_ext = ""

    if preferred_ext:
        if not new_stem.lower().endswith(preferred_ext):
            final_ext = preferred_ext
    elif leaf_ext and not new_stem.lower().endswith(leaf_ext):
        final_ext = leaf_ext
    else:
        if mime0 and (mime0 in MIME_TO_EXT):
            guessed = MIME_TO_EXT[mime0]
            if not new_stem.lower().endswith(guessed):
                final_ext = guessed

    # OOXMLヒントが得られた場合は最優先で上書き
    if ooxml_hint and not new_stem.lower().endswith(ooxml_hint):
        final_ext = ooxml_hint

    final_name = new_stem + final_ext
    dst = ensure_unique(out_parent / final_name)

    if dry_run:
        log_append(log, f"[PLAN] {rel_under_attachments} -> {dst.relative_to(attach_root)}"); return

    with open(dst, "wb") as fw:
        fw.write(data)
    log_append(log, f"[OK] {dst.relative_to(attach_root)}")

# ---------- 再配置 ----------
def _base_from_pathtxt(ptxt: Path) -> Path:
    name = ptxt.name
    if name.endswith(".path.txt"):
        return ptxt.with_name(name[:-9])
    return ptxt.with_suffix("")

def rehome_by_page_title(base_root: Path, log: tk.Text, progress_cb=None):
    path_files = list(base_root.rglob("*.path.txt"))
    target = len(path_files)
    moved = 0
    num_count = 0
    for ptxt in path_files:
        num_count += 1
        base = _base_from_pathtxt(ptxt)
        if not base.exists():
            expected = _base_from_pathtxt(ptxt).name
            parent = ptxt.parent
            found = None
            for cand in parent.iterdir():
                if cand.is_file() and cand.name.lower() == expected.lower():
                    found = cand; break
            if found is None:
                log_append(log, f"[WARN] 本体が見つからないためスキップ: {ptxt} -> {expected}")
                continue
            base = found

        page_title = ""
        try:
            text = ptxt.read_text(encoding="utf-8", errors="ignore")
            for line in text.splitlines():
                if line.startswith("PageTitle:"):
                    page_title = line.split(":", 1)[1].strip(); break
        except Exception:
            continue
        if not page_title: page_title = "Untitled Page"

        dst_dir = base_root / sanitize(page_title)
        dst_dir.mkdir(parents=True, exist_ok=True)

        dst_file = ensure_unique(dst_dir / base.name)
        dst_path = dst_file.with_suffix(dst_file.suffix + ".path.txt")

        shutil.move(str(base), str(dst_file))
        shutil.move(str(ptxt), str(dst_path))
        moved += 1
        log_append(log, f"[MOVE] {base}  ->  {dst_file.relative_to(base_root)}")
        _step_progress(20, 40, num_count, target, progress_cb)
        pump_gui(log)

    removed = 0
    dirs = sorted((p for p in base_root.rglob("*") if p.is_dir()), key=lambda p: -len(p.parts))
    for d in dirs:
        try:
            if not any(d.iterdir()):
                d.rmdir(); removed += 1
                log_append(log, f"[RM] 空フォルダ削除: {d.relative_to(base_root)}")
        except Exception:
            pass
        pump_gui(log)

    log_append(log, f"[REHOME] PageTitle再配置: {moved} files / 空フォルダ削除: {removed}\n")

# ---------- 添付復元（ZIP/Folder） ----------
def process_zip(zip_path: Path, attach_root: Path, dry_run: bool, log: tk.Text, progress_cb=None):
    with zipfile.ZipFile(str(zip_path), "r") as zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        attach_roots = sorted({ n.split("attachments/")[0] + "attachments/"
                                for n in names if "attachments/" in n })
        if not attach_roots:
            log_append(log, "[WARN] 'attachments/' が見つかりません．"); return
        attach_root_in_zip = min(attach_roots, key=len)

        ent_candidates = [n for n in names if n.lower().endswith("entities.xml")]
        entities_name = None
        parent = "/".join(attach_root_in_zip.strip("/").split("/")[:-1])
        for n in ent_candidates:
            if n.startswith(parent + "/"):
                entities_name = n; break
        if not entities_name and ent_candidates:
            entities_name = ent_candidates[0]

        spaces = pages = att_title = att_to_page = att_filename = {}
        if entities_name:
            spaces, pages, att_title, att_to_page, att_filename = parse_entities(zf.read(entities_name))
            log_append(log, f"[INFO] entities.xml: {entities_name}  pages={len(pages)} atts={len(att_title)} filenames={len(att_filename)}")
        else:
            log_append(log, "[WARN] entities.xml が見つからないため，タイトル置換/元名出力は限定的になります．")

        targets = len([n for n in names if n.startswith(attach_root_in_zip)])
        num_count = 0
        for n in [n for n in names if n.startswith(attach_root_in_zip)]:
            num_count += 1
            rel = Path(n[len(attach_root_in_zip):])
            parts = rel.parts
            page_id = parts[0] if len(parts) >= 2 else None
            att_id  = parts[1] if len(parts) >= 2 else None

            title_for_name = att_title.get(att_id, None)
            preferred_ext_from_entities = att_filename.get(att_id, None) if att_id else None

            data = zf.read(n)
            write_output(rel, data, attach_root,
                        title_for_name, preferred_ext_from_entities,
                        page_id, att_id, spaces, pages, dry_run, log)
            pump_gui(log)
            _step_progress(0, 20, num_count, targets, progress_cb)

        log_append(log, "[SUMMARY] 添付復元完了（ZIP）\n")

def process_folder(src_root: Path, attach_root: Path, dry_run: bool, log: tk.Text, progress_cb=None):
    candidates = sorted([p for p in src_root.rglob("attachments") if p.is_dir()],
                        key=lambda p: len(str(p)))
    if not candidates:
        log_append(log, "[WARN] 'attachments' が見つかりません．"); return
    attach_root_on_disk = candidates[0]

    ent_file = next((p for p in [attach_root_on_disk.parent / "entities.xml"] if p.exists()), None)
    if not ent_file:
        ents = list(src_root.rglob("entities.xml"))
        ent_file = ents[0] if ents else None

    spaces = pages = att_title = att_to_page = att_filename = {}
    if ent_file:
        spaces, pages, att_title, att_to_page, att_filename = parse_entities(ent_file.read_bytes())
        log_append(log, f"[INFO] entities.xml: {ent_file}  pages={len(pages)} atts={len(att_title)} filenames={len(att_filename)}")
    else:
        log_append(log, "[WARN] entities.xml が見つからないため，タイトル置換/元名出力ではなくIDで出力します．")

    total_entries = len([p for p in attach_root_on_disk.rglob("*") if p.is_file()])
    num_count = 0
    for f in [p for p in attach_root_on_disk.rglob("*") if p.is_file()]:
        num_count += 1
        rel = f.relative_to(attach_root_on_disk)
        parts = rel.parts
        page_id = parts[0] if len(parts) >= 2 else None
        att_id  = parts[1] if len(parts) >= 2 else None

        title_for_name = att_title.get(att_id, None)
        preferred_ext_from_entities = att_filename.get(att_id, None) if att_id else None

        data = f.read_bytes()
        write_output(rel, data, attach_root,
                    title_for_name, preferred_ext_from_entities,
                    page_id, att_id, spaces, pages, dry_run, log)
        pump_gui(log)
        _step_progress(0, 20, num_count, total_entries, progress_cb)

    log_append(log, "[SUMMARY] 添付復元完了（Folder）\n")

@lru_cache(maxsize=8192)
def _rel_href_from_cached(from_dir_str: str, target_str: str) -> str:
    rel = os.path.relpath(target_str, start=from_dir_str)
    rel_posix = rel.replace("\\", "/")
    parts = rel_posix.split("/")
    return "/".join(urllib.parse.quote(p, safe="!$&'()*+,;=:@[]-%._") for p in parts if p)

def _rel_href_from(from_dir: Path, target: Path) -> str:
    return _rel_href_from_cached(str(from_dir), str(target))

# ----------------------------------------------------------------------------------
# チェックボックス／添付リンク／ページ内リンク
# ----------------------------------------------------------------------------------
def confluence_storage_to_html(storage_html: str, page_titles_chain: list[str],
                                html_root: Path, out_root: Path,
                                *, attach_index: dict | None = None,
                                resolved_icons: dict | None = None) -> str:
    parser = "lxml"
    try:
        soup = BeautifulSoup(storage_html or "", parser)
        
        # --- BackToTop の CSS/JS を <head> に一度だけ注入 ---
        head = soup.find("head")
        if not head:
            # まれに <head> が無い HTML もあるので生成しておく
            html_tag = soup.find("html") or soup
            head = soup.new_tag("head")
            if html_tag.contents:
                html_tag.insert(0, head)
            else:
                html_tag.append(head)

        # CSS
        if not head.find(id="backToTop-style"):
            style_tag = soup.new_tag("style", id="backToTop-style")
            style_tag.string = """
                #backToTop{
                position:fixed; right:24px; bottom:-140px; width:80px; height:80px;
                padding:6px; border-radius:12px; background:#ffffffcc; box-shadow:0 6px 18px rgba(0,0,0,.22);
                cursor:pointer; transition:transform .4s ease, bottom .4s ease, opacity .3s ease;
                z-index:10050; opacity:.0; backdrop-filter: blur(6px);
                }
                #backToTop.show{ bottom:24px; opacity:1; }
                #backToTop img{ width:100%; height:100%; object-fit:contain; display:block; }
                #scrollSentinel{ position:relative; width:100%; height:1px; }
                """
            head.append(style_tag)

        # JS
        if not head.find(id="backToTop-script"):
            script_tag = soup.new_tag("script", id="backToTop-script")
            script_tag.string = r"""
                (function(){
                // DOMReady ラッパ
                function ready(fn){
                    if (document.readyState !== "loading") fn();
                    else document.addEventListener("DOMContentLoaded", fn);
                }

                // 最寄りのスクロール可能祖先を探索（overflow-y: auto/scroll）
                function getScrollableAncestor(el){
                    var cur = el && el.parentElement;
                    while (cur) {
                    var s = getComputedStyle(cur).overflowY;
                    if (s === "auto" || s === "scroll") return cur;
                    cur = cur.parentElement;
                    }
                    return window; // 無ければウィンドウを root に
                }

                ready(function(){
                    var btn  = document.getElementById("backToTop");
                    var sent = document.getElementById("scrollSentinel");
                    if (!btn || !sent) return;

                    // 監視対象となるスクロール root を決定
                    var scrollRoot = getScrollableAncestor(sent);

                    // 末尾付近かどうかの数値判定（フォールバック & 初期表示用）
                    function toggleByBottom(){
                    if (scrollRoot === window) {
                        var doc = document.documentElement;
                        var nearBottom = (window.scrollY + window.innerHeight) >= (doc.scrollHeight - 2);
                        if (nearBottom) btn.classList.add("show");
                        else            btn.classList.remove("show");
                    } else {
                        var nearBottom = (scrollRoot.scrollTop + scrollRoot.clientHeight) >= (scrollRoot.scrollHeight - 2);
                        if (nearBottom) btn.classList.add("show");
                        else            btn.classList.remove("show");
                    }
                    }

                    // IntersectionObserver（root を scrollRoot に合わせる）
                    try {
                    var io = new IntersectionObserver(function(entries){
                        var e = entries[0];
                        if (e && e.isIntersecting) btn.classList.add("show");
                        else                       btn.classList.remove("show");
                    }, {
                        root: (scrollRoot === window ? null : scrollRoot),
                        rootMargin: "0px 0px -10% 0px",
                        threshold: 0
                    });
                    io.observe(sent);
                    } catch(_) {
                    // 古い環境など IO 未対応時は数値判定のみで対応
                    }

                    // スクロール／リサイズ／ロードで再判定（scrollRoot にもハンドラを付ける）
                    var addEvt = (scrollRoot === window) ? window : scrollRoot;
                    addEvt.addEventListener("scroll", toggleByBottom, {passive:true});
                    window.addEventListener("resize", toggleByBottom, {passive:true});
                    window.addEventListener("load",   toggleByBottom);

                    // クリックでスムーズスクロール
                    function goTop(){
                    try { (scrollRoot === window ? window : scrollRoot).scrollTo({top:0, behavior:"smooth"}); }
                    catch(_) { (scrollRoot === window ? window : scrollRoot).scrollTop = 0; }
                    }
                    btn.addEventListener("click", function(ev){ ev.preventDefault(); goTop(); });
                    btn.addEventListener("keydown", function(ev){
                    if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); goTop(); }
                    });

                    // 初期判定
                    toggleByBottom();
                });
                })();
                """
            head.append(script_tag)
        # --- /BackToTop ヘッダ注入 ここまで ---

        
    except Exception:
        soup = BeautifulSoup(storage_html or "", "html.parser")
    title = page_titles_chain[-1] if page_titles_chain else "その他"
    
    # --- ユーティリティ ---
    def _page_title() -> str:
        return page_titles_chain[-1] if page_titles_chain else "その他"

    def _attach_folder_for_page() -> Path:
        return out_root / "添付ファイル" / _page_title()

    def _html_dir_for_page() -> Path:
        # このページ（page_titles_chain）の HTML が出力されるディレクトリを返す
        chain = page_titles_chain or ["その他"]
        if len(chain) == 1:
            return html_root / "その他"
        d = html_root
        for t in chain[:-1]:
            d = d / sanitize(t)
        return d

    def _candidate_attach_paths(filename: str, attach_dir: Path) -> list[Path]:
        name_lower = (filename or "").strip().lower()
        base, ext = os.path.splitext(name_lower)
        if attach_index:
            cands: list[Path] = []
            # 1) 完全一致（大小無視）
            hit = attach_index["by_lower"].get(name_lower)
            if hit:
                cands.append(hit)
            # 2) 拡張子ゆらぎ
            if not cands:
                alt_map = {".jpg": [".jpeg"], ".jpeg": [".jpg"], ".png": [".jpg", ".jpeg"]}
                for alt in alt_map.get(ext, []):
                    altname = base + alt
                    hit2 = attach_index["by_lower"].get(altname)
                    if hit2:
                        cands.append(hit2)
                        break
            # 3) stem 一致
            if not cands:
                cands.extend(attach_index["by_stem"].get(base, []))
            return cands

        # フォールバック（従来の全探索．なるべく通らないようにする）
        candidates = []
        for a in attach_dir.glob("**/*"):
            if a.name.lower() == name_lower:
                candidates.append(a)
        if not candidates:
            alt_exts = {".jpg": [".jpeg"], ".jpeg": [".jpg"], ".png": [".jpg", ".jpeg"]}.get(ext, [])
            for alt in alt_exts:
                alt_name = base + alt
                for a in attach_dir.glob("**/*"):
                    if a.name.lower() == alt_name:
                        candidates.append(a)
        if not candidates:
            for a in attach_dir.glob("**/*"):
                if a.name.lower().startswith(base):
                    candidates.append(a)
        return candidates

    def _href_to_attachment(filename: str) -> tuple[str, str]:
        # 添付の href と label を返す（相対パス版）．
        # 見つからない時はログして，最後の希望として “とりあえず期待パス” を返す
        attach_dir = _get_attach_dir(out_root)
        page_dir = _html_dir_for_page()  # ← このページのHTMLが出力されるフォルダ

        # 候補検索
        cands = _candidate_attach_paths(filename, attach_dir)
        if cands:
            target = cands[0]
        else:
            _log_not_found_attachment(out_root, filename, "href_to_attachment")
            target = attach_dir / filename  # 存在しない可能性あり

        # 相対パスでリンクを作る
        href = _rel_href_from(page_dir, target)

        label = Path(filename).name
        return href, label

    def _folder_href_for_attachment(filename: str) -> str:
        # 添付ファイルが置かれているフォルダへの “相対パス” を返す
        attach_dir = _get_attach_dir(out_root)
        candidates = _candidate_attach_paths(filename, attach_dir)
        target = next((c for c in candidates if c.exists()), attach_dir / filename)
        parent = target.parent if target.exists() else attach_dir

        # このページのHTMLが出力されるフォルダを基準に相対パス化
        page_dir = _html_dir_for_page()
        return _rel_href_from(page_dir, parent)

    # --- view-file マクロ（画像/ファイル） ---
    for macro in soup.find_all(lambda t: t.name and t.name.endswith("structured-macro")):
        if macro.get("ac:name") == "view-file":
            param_name = macro.find(lambda t: t.name and t.name.endswith("parameter") and t.get("ac:name")=="name")
            if param_name:
                attach = param_name.find(lambda t: t.name and t.name.endswith("attachment"))
                if attach and attach.has_attr("ri:filename"):
                    filename = (attach.get("ri:filename") or "").strip()
                    if filename:
                        href, label = _href_to_attachment(filename)
                        ext = os.path.splitext(filename)[1].lower()
                        
                        # 動画ファイルが含まれていた際の処理 -------------------------------------------------
                        video_exts = {".mp4", ".webm", ".ogv", ".ogg", ".m4v"}
                        audio_exts = {".mp3", ".wav", ".m4a", ".ogg"}

                        # サイズ指定（任意／無ければ自動）
                        wparam = macro.find(lambda t: t.name and t.name.endswith("parameter") and t.get("ac:name")=="width")
                        hparam = macro.find(lambda t: t.name and t.name.endswith("parameter") and t.get("ac:name")=="height")
                        w = (wparam.get_text(strip=True) if wparam else "") or ""
                        h = (hparam.get_text(strip=True) if hparam else "") or ""
                        size_attr = (f' width="{w}"' if w.isdigit() else "") + (f' height="{h}"' if h.isdigit() else "")

                        if ext in video_exts:
                            macro.replace_with(BeautifulSoup(
                                f'<figure class="confluence-video">'
                                f'  <video controls preload="metadata"{size_attr}>'
                                f'    <source src="{href}" type="video/{ext.lstrip(".")}">'
                                f'    <a href="{href}" target="_blank" rel="noopener">{label}</a>'
                                f'  </video>'
                                f'</figure>',
                                "html.parser"
                            ))
                            continue

                        if ext in audio_exts:
                            macro.replace_with(BeautifulSoup(
                                f'<p class="confluence-audio">'
                                f'  <audio controls preload="metadata">'
                                f'    <source src="{href}" type="audio/{ext.lstrip(".")}">'
                                f'    <a href="{href}" target="_blank" rel="noopener">{label}</a>'
                                f'  </audio>'
                                f'</p>',
                                "html.parser"
                            ))
                            continue
                        # ---------------------------------------------------
                        # 画像拡張子なら本文インライン（サムネ）＋クリックで拡大
                        if ext.lower() in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg"}:
                            macro.replace_with(BeautifulSoup(
                                f'<figure class="confluence-image">'
                                f'  <a href="{href}" class="zoom" aria-label="画像を拡大">'
                                f'    <img src="{href}" class="thumb" alt="{label}">'
                                f'  </a>'
                                f'</figure>',
                                "html.parser"
                            ))
                        else:
                            folder = _folder_href_for_attachment(filename)
                            macro.replace_with(BeautifulSoup(
                                f'<p><a href="{href}" target="_blank" rel="noopener">{label}</a>'
                                f' <a class="open-folder" href="{folder}" target="_blank" title="フォルダを開く" aria-label="フォルダを開く">📁</a></p>',
                                "html.parser"))
                        continue
            macro.unwrap()
    
    # --- ac:multimedia / ac:structured-macro name="multimedia" を動画として扱う ---
    for mm in soup.find_all(lambda t: (
        (t.name and t.name.endswith("structured-macro") and t.get("ac:name") == "multimedia")
        or (t.name and t.name.endswith("multimedia"))
    )):
        # 添付のファイル名を拾う
        ri = mm.find(lambda t: t.name and t.name.endswith("attachment"))
        url = None
        filename = None
        if ri and ri.has_attr("ri:filename"):
            filename = (ri.get("ri:filename") or "").strip()
        else:
            uri = mm.find(lambda t: t.name and t.name.endswith("url"))
            if uri and uri.has_attr("ri:value"):
                url = (uri.get("ri:value") or "").strip()

        if filename:
            href, label = _href_to_attachment(filename)
            ext = os.path.splitext(filename)[1].lower()
            if ext in {".mp4", ".webm", ".ogv", ".ogg", ".m4v"}:
                mm.replace_with(BeautifulSoup(
                    f'<figure class="confluence-video"><video controls preload="metadata">'
                    f'  <source src="{href}" type="video/{ext.lstrip(".")}">'
                    f'  <a href="{href}" target="_blank" rel="noopener">{label}</a>'
                    f'</video></figure>', "html.parser"))
                continue
            if ext in {".mp3", ".wav", ".m4a", ".ogg"}:
                mm.replace_with(BeautifulSoup(
                    f'<p class="confluence-audio"><audio controls preload="metadata">'
                    f'  <source src="{href}" type="audio/{ext.lstrip(".")}">'
                    f'  <a href="{href}" target="_blank" rel="noopener">{label}</a>'
                    f'</audio></p>', "html.parser"))
                continue
        elif url and (url.lower().endswith(".mp4") or url.lower().endswith(".webm")):
            mm.replace_with(BeautifulSoup(
                f'<figure class="confluence-video"><video controls preload="metadata" src="{url}"></video></figure>',
                "html.parser"))
            continue

    # --- 画像マクロ <ac:image> を <img>（サムネ）＋リンク（フル）に変換 ---
    for aimg in soup.find_all(lambda t: t.name and t.name.endswith("image")):
        ri_att = aimg.find(lambda t: t.name and t.name.endswith("attachment"))
        if ri_att and ri_att.has_attr("ri:filename"):
            filename = (ri_att.get("ri:filename") or "").strip()
            if filename:
                src, _ = _href_to_attachment(filename)   # フル画像のURL/パス
                html = (
                    f'<figure class="confluence-image">'
                    f'  <a href="{src}" class="zoom" aria-label="画像を拡大">'
                    f'    <img src="{src}" class="thumb" alt="{filename}">'
                    f'  </a>'
                    f'</figure>'
                )
                aimg.replace_with(BeautifulSoup(html, "html.parser"))
                continue
        ri_url = aimg.find(lambda t: t.name and t.name.endswith("url"))
        if ri_url and ri_url.has_attr("ri:value"):
            url = (ri_url.get("ri:value") or "").strip()
            if url:
                html = (
                    f'<figure class="confluence-image">'
                    f'  <a href="{url}" class="zoom" aria-label="画像を拡大">'
                    f'    <img src="{url}" class="thumb" alt="">'
                    f'  </a>'
                    f'</figure>'
                )
                aimg.replace_with(BeautifulSoup(html, "html.parser"))
                continue
        aimg.replace_with(BeautifulSoup('<p>[画像]</p>', "html.parser"))

    # --- ac:link + ri:attachment / ri:page ---
    for alink in soup.find_all(lambda t: t.name and t.name.endswith("link")):
        # 添付
        ri_attach = alink.find(lambda t: t.name and t.name.endswith("attachment"))
        if ri_attach and ri_attach.has_attr("ri:filename"):
            filename = (ri_attach.get("ri:filename") or "").strip()
            if filename:
                href, _label = _href_to_attachment(filename)
                label = alink.get_text(strip=True) or _label
                folder = _folder_href_for_attachment(filename)
                alink.replace_with(BeautifulSoup(
                    f'<p><a href="{href}" target="_blank" rel="noopener">{label}</a>'
                    f' <a class="open-folder" href="{folder}" target="_blank" title="フォルダを開く" aria-label="フォルダを開く">📁</a></p>',
                    "html.parser"))
                continue
        # ページ
        ri_page = alink.find(lambda t: t.name and t.name.endswith("page"))
        if ri_page and ri_page.has_attr("ri:content-title"):
            title = (ri_page.get("ri:content-title") or "").strip()
            if title:
                safe = sanitize(title)
                label = alink.get_text(strip=True) or title
                # 同じディレクトリ内の HTML にリンク（存在チェックはしない/後で作る）
                alink.replace_with(BeautifulSoup(f'<p><a href="{safe}.html">{label}</a></p>', "html.parser"))
                continue

    # Confluence名前空間タグは中身だけ残す
    for tag in list(soup.find_all()):
        if ":" in tag.name:
            tag.unwrap()

    # --- 本文HTML（段落間隔の体裁をちょっと整える） ---
    body_html = str(soup)

    # --- 背景画像 BG_01.png への相対パスを計算 ---
    page_dir = _html_dir_for_page()
    
    # --- 末尾フッター：添付格納先（相対リンク化） ---
    attach_dir = _attach_folder_for_page().resolve()
    footer_rel = _rel_href_from(page_dir, attach_dir)
    if not footer_rel.endswith("/"):
        footer_rel += "/"

    # 事前解決（resolved_icons）があればそれを使う
    bg_url = (resolved_icons.get("bg") if resolved_icons else "") or ""
    logo_url = (resolved_icons.get("logo") if resolved_icons else "") or ""
    exe_icon_rel = (resolved_icons.get("exe_icon") if resolved_icons else "") or ""
    pokeball_rel = (resolved_icons.get("pokeball") if resolved_icons else "") or ""
    empty_icon_rel = (resolved_icons.get("empty_icon") if resolved_icons else "") or ""
    back_top_img = (resolved_icons.get("back_top_img") if resolved_icons else "") or ""

    # 事前解決が無いときだけ従来の確保・相対化を行う
    if not bg_url:
        try:
            bg_candidate = _get_attach_dir(out_root) / "BG_01.png"
            if bg_candidate.exists():
                bg_url = _rel_href_from(page_dir, bg_candidate)
        except Exception:
            bg_url = ""

    if not logo_url:
        try:
            logo_candidate = _ensure_header_logo(out_root)
            if logo_candidate.exists():
                logo_url = _rel_href_from(out_root, logo_candidate)
        except Exception:
            logo_url = ""

    if not exe_icon_rel:
        try:
            exe_icon_path = _ensure_exe_icon(out_root)
            if exe_icon_path is not None:
                exe_icon_rel = _rel_href_from(page_dir, exe_icon_path)
        except Exception:
            exe_icon_rel = ""

    if not pokeball_rel:
        try:
            pokeball_path = _ensure_pokeball(out_root)
            if pokeball_path is not None:
                pokeball_rel = _rel_href_from(page_dir, pokeball_path)
        except Exception:
            pokeball_rel = ""

    if not empty_icon_rel:
        try:
            empty_icon_path = _ensure_empty_icon(out_root)
            if empty_icon_path is not None:
                empty_icon_rel = _rel_href_from(page_dir, empty_icon_path)
        except Exception:
            empty_icon_rel = ""
            
    if not back_top_img:
        try:
            gori_path = _ensure_footer_gori(out_root)   # 添付/ footer_gori1.png を保証
            if gori_path is not None:
                back_top_img = _rel_href_from(page_dir, gori_path)  # ページ基準の相対パスへ
        except Exception:
            back_top_img = ""

    # フッター用画像
    back_to_top_html = ""
    if back_top_img:
        back_to_top_html = f"""
        <div id="backToTop"
            role="button"
            tabindex="0"
            aria-label="ページの先頭へ戻る"
            title="ページの先頭へ戻る">
        <img src="{back_top_img}" alt="トップに戻る">
        </div>
        <div id="scrollSentinel" aria-hidden="true"></div>
        """
    # フッター本体
    footer_html = f"""
    <hr>
    <div class="footer" style="margin-top:20px;">
        <a href="{footer_rel}"
        class="open-folder"
        title="添付ファイル格納先を開く"
        aria-label="添付ファイル格納先を開く"
        style="font-size:14px; text-decoration:none;">
        📁 添付ファイル格納先を開く
        </a>
    </div>
    """

    # このページ位置から見たサイドバー用リンクを作る
    def _generate_sidebar_links_for_current_page() -> str:
        parts: list[str] = []

        # --- Top に戻るリンク ---
        index_path = out_root / "index.html"
        top_href = _rel_href_from(page_dir, index_path)
        parts.append(
            f'<div class="sidebar-toplink"><a href="{top_href}">Topに戻る</a></div>'
        )

        if not SIDEBAR_ITEMS:
            return "\n".join(parts)

        # ================================
        # ①：同名タイトルの重複除去
        #     → (大項目, タイトル) 単位で「最新だけ」残す
        # ================================
        latest: dict[tuple[str, str], tuple[list[str], Path]] = {}
        for chain, pth in SIDEBAR_ITEMS:
            if len(chain) == 1:
                group_key = "その他"
            else:
                group_key = chain[0]        # 例：ポケモンディレクションレッジ

            title = chain[-1]
            latest[(group_key, title)] = (chain, pth)

        # ================================
        # ②：大項目ごとに分類
        # ================================
        groups: dict[str, list[tuple[str, str]]] = {}

        for (group_key, title), (_chain, pth) in latest.items():
            href = _rel_href_from(page_dir, pth)
            is_empty = pth in SIDEBAR_EMPTY_PAGES
            groups.setdefault(group_key, []).append((title, href, is_empty))

        # ================================
        # ②.5：このページ自身の href を計算
        # ================================
        current_title = _page_title()
        current_page_dir = _html_dir_for_page()
        current_page_path = current_page_dir / f"{sanitize(current_title)}.html"
        current_href = _rel_href_from(page_dir, current_page_path)

        # ================================
        # ③：サイドバー HTML 出力
        # ================================
        # 「その他」が一番下に来るように並び替え
        ordered_groups = sorted(
            groups.keys(),
            key=lambda g: (g == "その他", g) 
        )

        for group in ordered_groups:
            parts.append('<div class="sidebar-section">')
            if pokeball_rel:
                title_html = (
                    f'<div class="sidebar-section-title">'
                    f'<img src="{pokeball_rel}" alt="•" class="sidebar-section-icon">'
                    f'<span>{group}</span></div>'
                )
            else:
                title_html = (
                    f'<div class="sidebar-section-title"><span>{group}</span></div>'
                )

            parts.append(title_html)

            for label, href, is_empty in sorted(groups[group], key=lambda x: x[0]):
                cls = "sidebar-link-current" if href == current_href else "sidebar-link"

                icon_html = ""
                if is_empty and empty_icon_rel:
                    icon_html = (
                        f'<img src="{empty_icon_rel}" class="sidebar-empty-icon" '
                        f'alt="Empty page">'
                    )
                parts.append(f'<a href="{href}" class="{cls}">{icon_html}{label}</a>')

            parts.append("</div>")

        return "\n".join(parts)

    sidebar_links_html = _generate_sidebar_links_for_current_page()

    # ライトボックス本体
    lightbox_html = """
    <div id="lightbox">
        <div class="lb-inner">
        <img src="" alt="">
        </div>
    </div>
    """

    # ライトボックス用スクリプト＋言語案内モーダル
    script_html = """
    <script>
    (function () {
        const lb  = document.getElementById("lightbox");
        const img = lb ? lb.querySelector("img") : null;

        function openLightbox(src, alt) {
            if (!lb || !img) return;
            img.src = src;
            img.alt = alt || "";
            lb.classList.add("open");
        }

        function closeLightbox() {
            if (!lb) return;
            lb.classList.remove("open");
        }

        // 画像クリックでライトボックス表示
        if (lb && img) {
            document.addEventListener("click", function (ev) {
                // a.zoom を起点にライトボックスを開く
                const a = ev.target.closest("a.zoom");
                if (a) {
                    ev.preventDefault();
                    openLightbox(
                        a.getAttribute("href"),
                        a.getAttribute("aria-label") || ""
                    );
                    return;
                }

                // オーバーレイやクローズボタンをクリックした場合は閉じる
                const t = ev.target;
                if (t.id === "lightbox" || t.classList.contains("lightbox__close")) {
                    closeLightbox();
                }
            });

            document.addEventListener("keyup", function (ev) {
                if (ev.key === "Escape") {
                    closeLightbox();
                }
            });
        }

        // --- サイドバーの幅変更 -------------------------------
        (function initSidebarResize() {
            const sidebar = document.querySelector(".sidebar");
            const resizer = document.getElementById("sidebar-resizer");
            if (!sidebar || !resizer) return;

            let dragging = false;
            let startX   = 0;
            let startW   = 0;

            resizer.addEventListener("mousedown", function (ev) {
                dragging = true;
                startX   = ev.clientX;
                startW   = sidebar.getBoundingClientRect().width;
                document.body.classList.add("resizing-sidebar");
                ev.preventDefault();
            });

            document.addEventListener("mousemove", function (ev) {
                if (!dragging) return;
                const dx = ev.clientX - startX;
                let newW = startW + dx;

                // 最小 / 最大幅の制限
                if (newW < 180) newW = 180;
                if (newW > 520) newW = 520;

                sidebar.style.width = newW + "px";
            });

            document.addEventListener("mouseup", function () {
                if (!dragging) return;
                dragging = false;
                document.body.classList.remove("resizing-sidebar");
            });
        })();

        // --- 言語案内モーダル -------------------------------
        (function initLangHelp() {
            const select  = document.getElementById("lang-select");
            const overlay = document.getElementById("lang-help-overlay");
            if (!select || !overlay) return;

            const modals = overlay.querySelectorAll(".lang-help-modal");

            function showModal(lang) {
                overlay.classList.add("is-open");
                overlay.setAttribute("aria-hidden", "false");

                modals.forEach(function (m) {
                    if (m.getAttribute("data-lang") === lang) {
                        m.style.display = "block";
                    } else {
                        m.style.display = "none";
                    }
                });

                // ブラウザ翻訳の検知用に lang 属性も合わせておく
                document.documentElement.lang = lang;
            }

            function hideModal() {
                overlay.classList.remove("is-open");
                overlay.setAttribute("aria-hidden", "true");
            }

            // プルダウン変更時にモーダル表示
            select.addEventListener("change", function () {
                const v = select.value || "ja";
                showModal(v);
            });

            // 背景クリック or × ボタンで閉じる
            overlay.addEventListener("click", function (ev) {
                if (ev.target === overlay || ev.target.hasAttribute("data-lang-help-close")) {
                    hideModal();
                }
            });

            // Esc キーでも閉じる
            document.addEventListener("keyup", function (ev) {
                if (ev.key === "Escape") {
                    hideModal();
                }
            });            
        })();

        // --- 現在ページの位置までサイドバーを自動スクロール ----------------
        (function scrollSidebarToCurrent() {
            const sidebar = document.querySelector(".sidebar");
            if (!sidebar) return;
            const current = sidebar.querySelector(".sidebar-link-current");
            if (!current) return;

            const sidebarRect = sidebar.getBoundingClientRect();
            const currentRect = current.getBoundingClientRect();
            const offsetTop   = currentRect.top - sidebarRect.top;
            const targetScroll =
                offsetTop - (sidebar.clientHeight / 2) + (current.offsetHeight / 2);

            sidebar.scrollTop = targetScroll;
        })();

    })();
    </script>
    """


    # 背景画像（BG_01.png）が見つからなかったときは単色背景にする
    bg_style = f"background: #f5f5f5 url('{bg_url}') repeat;" if bg_url else "background: #f5f5f5;"

    full = f"""<!DOCTYPE html>
    <html lang="ja">
    <head>
    <meta charset="utf-8">
    <title>{title}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body {{
        margin: 0;
        padding: 0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif;
        {bg_style}
        }}

        /* 右上の言語選択バー */
        .topbar {{
        position: fixed;
        top: 0;
        right: 0;
        padding: 8px 16px;
        font-size: 12px;
        z-index: 2100;
        background: rgba(255, 255, 255, 0.9);
        border-bottom-left-radius: 8px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
        }}
        .topbar label {{
        margin-right: 4px;
        }}
        .lang-select {{
        padding: 2px 4px;
        font-size: 12px;
        }}

        /* 言語案内モーダル */
        .lang-help-overlay {{
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.45);
        display: none;
        align-items: center;
        justify-content: center;
        z-index: 2200;
        }}
        .lang-help-overlay.is-open {{
        display: flex;
        }}
        .lang-help-modal {{
        position: relative;
        background: #ffffff;
        max-width: 520px;
        width: 90%;
        padding: 1.25rem 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 18px 45px rgba(0,0,0,0.35);
        font-size: 14px;
        line-height: 1.5;
        }}
        .lang-help-modal h2 {{
        margin-top: 0;
        margin-bottom: 0.5rem;
        font-size: 16px;
        }}
        .lang-help-modal p {{
        margin: 0.4rem 0;
        }}
        .lang-help-modal ol {{
        margin: 0.4rem 0 0.2rem;
        padding-left: 1.4rem;
        }}
        .lang-help-close {{
        position: absolute;
        top: 0.35rem;
        right: 0.5rem;
        border: none;
        background: transparent;
        font-size: 18px;
        cursor: pointer;
        }}

        .layout {{
        display: flex;
        min-height: 100vh;
        }}

        /* サイドバー */
        .sidebar {{
        width: 260px;
        box-sizing: border-box;
        padding: 16px 12px;
        background: rgba(255, 255, 255, 0.92);
        border-right: 1px solid #e5e7eb;
        overflow-y: auto;
        position: sticky;
        top: 0;
        max-height: 100vh;
        }}
        .sidebar-resizer {{
        width: 5px;
        cursor: col-resize;
        background: transparent;
        }}
        .sidebar-resizer:hover {{
        background: rgba(148, 163, 184, 0.4);
        }}
        body.resizing-sidebar {{
        cursor: col-resize;
        user-select: none;
        }}
        .sidebar-title {{
        display: flex;
        align-items: center;
        gap: 6px;
        font-size: 18px;
        font-weight: 600;
        margin-bottom: 8px;
        }}
        .sidebar-title-icon {{
        width: 20px;
        height: 20px;
        flex-shrink: 0;
        }}
        .sidebar-empty-icon {{
        width: 24px;
        height: 24px;
        margin-right: 4px;
        vertical-align: text-bottom;
        }}
        .sidebar-links a {{
        display: block;
        font-size: 14px;
        padding: 4px 2px;
        color: #2563eb;
        text-decoration: none;
        border-radius: 4px;
        }}
        .sidebar-links a:hover {{
        background: #e5f0ff;
        }}
        .sidebar-links a.sidebar-link-current {{
        background: rgba(248, 113, 113, 0.25);  /* 薄い赤背景 */
        color: #b91c1c;                          /* 文字も少し濃い赤に */
        font-weight: 600;
        }}

        /*Topリンクとセクション区切り */
        .sidebar-toplink {{
        margin-bottom: 12px;
        padding-bottom: 8px;
        border-bottom: 1px solid #e5e7eb;
        }}
        .sidebar-toplink a {{
        font-weight: 600;
        color: #111827;
        text-decoration: none;
        }}
        .sidebar-toplink a:hover {{
        text-decoration: underline;
        }}
        .sidebar-section {{
        margin-bottom: 16px;
        padding-bottom: 8px;
        border-bottom: 1px solid #e5e7eb;
        }}
        .sidebar-section-title {{
        display: flex;
        align-items: center;
        gap: 6px;
        font-size: 13px;
        font-weight: 700;
        color: #6b7280;
        margin-bottom: 4px;
        }}
        .sidebar-section-icon {{
        width: 16px;
        height: 16px;
        flex-shrink: 0;
        }}
        .sidebar-section a {{
        display: block;
        font-size: 13px;
        padding: 2px 2px;
        color: #2563eb;
        text-decoration: none;
        border-radius: 3px;
        }}
        .sidebar-section a:hover {{
        background: #e5f0ff;
        }}

        /*ページタイトル用 */
        .page-header {{
        margin-bottom: 1.5rem;
        border-bottom: 1px solid #e5e7eb;
        padding-bottom: .75rem;
        }}
        .page-title {{
        margin: 0;
        font-size: 1.6rem;
        font-weight: 700;
        }}

        /* 本文側 */
        .page-container {{
        flex: 1;
        display: flex;
        justify-content: center;
        box-sizing: border-box;
        padding: 24px;
        }}
        .content-box {{
        background: #ffffff;
        max-width: 960px;
        width: 100%;
        box-shadow: 0 10px 30px rgba(15,23,42,0.15);
        border-radius: 8px;
        padding: 32px 40px 40px;
        box-sizing: border-box;
        }}
        .content-box h1,
        .content-box h2,
        .content-box h3 {{
        margin-top: 1.6em;
        }}
        .content-box p {{
        line-height: 1.8;
        margin: 0.5em 0;
        }}
        .content-box ul,
        .content-box ol {{
        padding-left: 1.6em;
        }}

        /* テーブル（Confluence の表） */
        .content-box table,
        .content-box table.confluenceTable {{
        border-collapse: collapse;
        border-spacing: 0;
        width: 100%;
        margin: 0.75rem 0;
        font-size: 14px;
        }}

        .content-box th,
        .content-box td,
        .content-box table.confluenceTable th,
        .content-box table.confluenceTable td {{
        border: 1px solid #e5e7eb;
        padding: 4px 8px;
        vertical-align: top;
        }}

        .content-box thead th,
        .content-box table.confluenceTable thead th {{
        background: #f9fafb;
        font-weight: 600;
        }}

        .footer {{
        margin-top: 2rem;
        font-size: 13px;
        color: #4b5563;
        border-top: 1px solid #e5e7eb;
        padding-top: 0.75rem;
        }}

        .confluence-image img {{
        max-width: 100%;
        height: auto;
        }}

        /* 画像クリック時のライトボックス */
        #lightbox {{
            position: fixed;
            inset: 0;
            background: rgba(0,0,0,0.75);
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 3000;
        }}

        /* open クラスが付与されたときだけ表示 */
        #lightbox.open {{
            display: flex;
        }}

        /* 画像の最大サイズを画面内に収める（重要） */
        #lightbox img {{
            max-width: 90vw;
            max-height: 90vh;
            object-fit: contain;
            border-radius: 8px;
            box-shadow: 0 0 20px rgba(0,0,0,0.4);
        }}
        
        /* 動画/音声のサイズ調整 */
        .content-box video,
        .content-box audio {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: .5rem 0;
        }}
        
        /* === ページ最下部 トップに戻る === */
        #backToTop{{
            position: fixed;
            right: max(16px, env(safe-area-inset-right));
            bottom: max(16px, env(safe-area-inset-bottom));
            width: 88px;        /* 画像サイズに合わせて調整 */
            height: 88px;
            z-index: 10030;     /* ヘッダやモーダルより手前/奥は環境に合わせ微調整 */
            transform: translateY(140%);
            opacity: 0;
            transition: transform .35s ease, opacity .35s ease;
            cursor: pointer;
            user-select: none;
            -webkit-tap-highlight-color: transparent;
            }}
        #backToTop.show{{
            transform: translateY(0);
            opacity: 1;
            }}
        #backToTop img{{
            display:block;
            width:100%;
            height:auto;
            filter: drop-shadow(0 2px 4px rgba(0,0,0,.35));
            }}
        #backToTop:focus-visible{{
            outline: 3px solid #3b82f6; /* アクセシビリティ */
            border-radius: 10px;
            }}
        /* ページ末尾監視用のダミー要素（高さ0でOK） */
            #scrollSentinel{{ width:1px; height:1px; }}


    </style>
    </head>

<body>
    <!-- 右上の言語選択プルダウン -->
    <div class="topbar">
        <label for="lang-select">Language:</label>
        <select id="lang-select" class="lang-select">
            <option value="ja">日本語</option>
            <option value="en">English</option>
        </select>
    </div>

    <!-- 言語ごとの翻訳案内モーダル -->
    <div id="lang-help-overlay" class="lang-help-overlay" aria-hidden="true">
        <!-- 日本語案内 -->
        <div class="lang-help-modal" data-lang="ja">
            <button type="button" class="lang-help-close" data-lang-help-close>&times;</button>
            <h2>ブラウザ翻訳の使い方（日本語）</h2>
            <p>このページは Confluence から復元した HTML です．ブラウザの翻訳機能を使うと，他の言語でも内容を確認できます．</p>
            <ol>
                <li>ブラウザのアドレスバー付近に表示される「翻訳」アイコンをクリックします．</li>
                <li>アイコンが表示されない場合は，ページ上で右クリックし「日本語に翻訳」などのメニューを選択します．</li>
                <li>「翻訳先の言語」で <strong>日本語</strong> を選択し，必要に応じて「このサイトは常に翻訳する」をオンにします．</li>
            </ol>
            <p>※ 翻訳処理はブラウザ側で行われ，このページのデータは変更されません．</p>
        </div>

        <!-- 英語案内 -->
        <div class="lang-help-modal" data-lang="en">
            <button type="button" class="lang-help-close" data-lang-help-close>&times;</button>
            <h2>How to use the browser&apos;s translation (English)</h2>
            <p>This page was restored from Confluence HTML. You can use your browser&apos;s built-in translation feature to read it in your preferred language.</p>
            <ol>
                <li>Click the translation icon near the address bar.</li>
                <li>If you don&apos;t see the icon, right-click on the page and choose “Translate to English” or similar.</li>
                <li>Select <strong>English</strong> as the target language and optionally enable “Always translate this site”.</li>
            </ol>
            <p>The translation itself is handled by the browser, not by this page.</p>
        </div>
    </div>

    <div class="layout">
        <aside class="sidebar">
            <div class="sidebar-title">
                {f'<img src="{exe_icon_rel}" alt="Index" class="sidebar-title-icon">' if exe_icon_rel else ''}
                <span>index</span>
            </div>
            <div class="sidebar-links">
                {sidebar_links_html}
            </div>
        </aside>
        <div class="sidebar-resizer" id="sidebar-resizer"></div>
        <div class="page-container">
        
        <article class="content-box">
            <header class="page-header">
            <h1 class="page-title">{title}</h1>
            </header>
            {body_html}
            {footer_html}
            {back_to_top_html}
        </article>
        </main>
    </div>

    {lightbox_html}
    {script_html}
    </body>
    </html>
    """
    return full

# ----------------------------------------------------------------------------------
# XML → Markdown / Word 生成（階層出力／親なし=その他／添付なし）
# ----------------------------------------------------------------------------------
def _html_to_plaintext(html: str) -> str:
    text = re.sub(r"<\s*br\s*/?\s*>", "\n", html or "", flags=re.I)
    text = re.sub(r"</p\s*>", "\n\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def _parse_pages_parent_from_entities_xml_bytes(ent_bytes: bytes):
    import xml.etree.ElementTree as ET
    pages = {}; body_html = {}
    root = ET.fromstring(ent_bytes)

    def pick(parent, paths):
        for p in paths:
            el = parent.find(p)
            if el is not None and (el.text or "").strip():
                return el.text.strip()
        return ""

    for obj in root.findall(".//object[@class='Page']"):
        pid   = pick(obj, ["id[@name='id']", "property[@name='id']"])
        title = pick(obj, ["property[@name='title']"])
        par   = ""
        par_prop = obj.find("property[@name='parent']")
        if par_prop is not None:
            par = pick(par_prop, ["id[@name='id']", "property[@name='id']"])
        if pid:
            pages[pid] = {"title": title or f"page_{pid}", "parentId": par}

    for obj in root.findall(".//object[@class='BodyContent']"):
        page_id = pick(obj, ["property[@name='content']/id[@name='id']"])
        html = pick(obj, ["property[@name='body']"])
        if page_id and html:
            body_html[page_id] = html

    return pages, body_html

def _collect_pages_parent_from_pages_dir(pages_dir: Path):
    import xml.etree.ElementTree as ET
    result = {}
    if not pages_dir.exists(): return result
    for xml in pages_dir.rglob("*.xml"):
        try:
            tree = ET.parse(xml); root = tree.getroot()
            page_obj = root.find(".//object[@class='Page']") or root
            pid   = (page_obj.findtext("id[@name='id']", "") or "").strip()
            title = (page_obj.findtext("property[@name='title']", "") or "").strip()
            parentId = ""
            par_prop = page_obj.find("property[@name='parent']")
            if par_prop is not None:
                parentId = (par_prop.findtext("id[@name='id']", "") or par_prop.findtext("property[@name='id']", "") or "").strip()
            html = (page_obj.findtext("property[@name='body']/property[@name='storage']", "") or "").strip()
            if pid or title or html:
                result[pid or f"(unknown)_{xml.stem}"] = {"title": title or xml.stem, "parentId": parentId, "html": html}
        except Exception:
            pass
    return result

def _build_dir_chain_for_page(pid: str, pages_map: dict) -> list[str]:
    chain = []; seen = set(); cur = pid
    while cur and cur in pages_map and cur not in seen:
        seen.add(cur)
        t = pages_map[cur].get("title") or f"page_{cur}"
        chain.append(sanitize(t))
        cur = pages_map[cur].get("parentId") or ""
    chain.reverse()
    return chain if chain else ["その他"]

def _dir_for_titles(dst_root: Path, titles: list[str]) -> tuple[Path, str]:
    page_title = titles[-1]
    if len(titles) == 1:
        page_dir = dst_root / "その他"
    else:
        page_dir = dst_root
        for t in titles[:-1]:
            page_dir = page_dir / t
    page_dir.mkdir(parents=True, exist_ok=True)
    return page_dir, page_title

def _parse_pages_and_bodies_from_entities_bytes(ent_bytes: bytes):
    """entities.xml から {pid: {title,parentId}}, {pid: body_html} を返す"""
    import xml.etree.ElementTree as ET
    pages, bodies = {}, {}
    root = ET.fromstring(ent_bytes)

    def pick(parent, paths):
        for p in paths:
            el = parent.find(p)
            if el is not None and (el.text or "").strip():
                return el.text.strip()
        return ""

    for obj in root.findall(".//object[@class='Page']"):
        pid   = pick(obj, ["id[@name='id']", "property[@name='id']"])
        title = pick(obj, ["property[@name='title']"])
        parentId = ""
        par = obj.find("property[@name='parent']")
        if par is not None:
            parentId = pick(par, ["id[@name='id']", "property[@name='id']"])
        if pid:
            pages[pid] = {"title": title or f"page_{pid}", "parentId": parentId}

    for obj in root.findall(".//object[@class='BodyContent']"):
        pid  = pick(obj, ["property[@name='content']/id[@name='id']"])
        html = pick(obj, ["property[@name='body']"])
        if pid and html:
            bodies[pid] = html
    return pages, bodies

def _build_chain(pid: str, pages_map: dict) -> list[str]:
    chain, seen, cur = [], set(), pid
    while cur and cur in pages_map and cur not in seen:
        seen.add(cur)
        t = sanitize(pages_map[cur].get("title") or f"page_{cur}")
        chain.append(t); cur = pages_map[cur].get("parentId") or ""
    chain.reverse()
    return chain if chain else ["その他"]

def _dir_for_chain(dst_root: Path, chain: list[str]) -> tuple[Path, str]:
    title = chain[-1]
    if len(chain) == 1:
        page_dir = dst_root / "その他"
    else:
        page_dir = dst_root
        for t in chain[:-1]:
            page_dir = page_dir / t
    page_dir.mkdir(parents=True, exist_ok=True)
    return page_dir, title

def _write_index_html(index_root: Path, html_root: Path, pages_map: dict, link_prefix: str = ""):
    # インデックス用ページの処理
    # --- 親→子 ---
    children = {}
    for pid, meta in pages_map.items():
        par = meta.get("parentId") or ""
        children.setdefault(par, []).append(pid)
    for v in children.values():
        v.sort(key=lambda p: (pages_map[p].get("title") or f"page_{p}").lower())

    # --- 重複タイトル除去 ---
    children_dedup = {}
    for par, pids in children.items():
        seen = set(); uniq = []
        for p in pids:
            t = sanitize(pages_map[p].get("title") or f"page_{p}")
            if t in seen: continue
            seen.add(t); uniq.append(p)
        children_dedup[par] = uniq

    def make_href(chain: list[str]) -> str:
        href_parts = ["その他"] if len(chain) == 1 else chain[:-1]
        href_dir = "/".join(href_parts)
        return f"{href_dir + '/' if href_dir else ''}{chain[-1]}.html"

    def render_children(pid: str, chain: list[str]) -> str:
        html = "<ul>"
        for ch in children_dedup.get(pid, []):
            title = sanitize(pages_map[ch].get("title") or f"page_{ch}")
            ch_chain = chain + [title]
            href = make_href(ch_chain)
            html += f"<li><a href='{href}' target='_blank' rel='noopener'>{title}</a>"
            html += render_children(ch, ch_chain)
            html += "</li>"
        html += "</ul>"
        return html

    top_pids = children_dedup.get("", [])
    sections_html = []

    def make_href(chain: list[str]) -> str:
        href_parts = ["その他"] if len(chain) == 1 else chain[:-1]
        href_dir = "/".join(href_parts)
        path_part = f"{href_dir + '/' if href_dir else ''}{chain[-1]}.html"
        return f"{link_prefix}{path_part}"

    # --- pokeball.png を「添付ファイル」フォルダに用意してから使う ---
    dst_ball = _ensure_pokeball(index_root)
    pokeball_exists = dst_ball is not None and dst_ball.exists()
    # index.html から見た相対パスは常に「添付ファイル/pokeball.png」
    pokeball_rel = f"{ATT_DIR_NAME}/{POKEBALL_FILE}" if pokeball_exists else POKEBALL_FILE

    # 子を持つ最上位 → セクション
    for top in top_pids:
        top_title = sanitize(pages_map[top].get("title") or f"page_{top}")
        has_children = len(children_dedup.get(top, [])) > 0
        if not has_children:
            continue

        icon_html = (f"<img class='poke' src='{pokeball_rel}' alt='•'/>"
                    if pokeball_exists else "<span class='dot'></span>")

        href_top = make_href([top_title])
        sec = []
        sec.append("<details class='sec' open>")
        sec.append(f"<summary>{icon_html}<span class='ttl'>{top_title}</span></summary>")
        sec.append("<ul class='root'>")
        sec.append(f"<li><a href='{href_top}' target='_blank' rel='noopener'>{top_title}</a>")
        sec.append(render_children(top, [top_title]))
        sec.append("</li></ul></details>")
        sections_html.append("".join(sec))

    # 子を持たない最上位は「その他」へ
    others = []
    for top in top_pids:
        if len(children_dedup.get(top, [])) == 0:
            t = sanitize(pages_map[top].get("title") or f"page_{top}")
            others.append(f"<li><a href='{make_href([t])}' target='_blank' rel='noopener'>{t}</a></li>")
    if others:
        icon_html = (f"<img class='poke' src='{pokeball_rel}' alt='•'/>"
                    if pokeball_exists else "<span class='dot'></span>")
        sections_html.append(
            "<details class='sec' open>"
            f"<summary>{icon_html}<span class='ttl'>その他</span></summary>"
            "<ul class='root'>" + "".join(others) + "</ul></details>"
        )
    
    # 背景画像を取得してタイル指定
    bg_url = ""
    try:
        bg_candidate = _get_attach_dir(index_root) / "BG_01.png"  # 添付フォルダ直下（既存の配置と同じ想定）
        if bg_candidate.exists():
            bg_url = _rel_href_from(index_root, bg_candidate)
    except Exception:
        bg_url = ""
        
    # 背景スタイルを決定した直後あたりに追加
    logo_url = ""
    try:
        p = _ensure_header_logo(index_root)
        if p:
            logo_url = _rel_href_from(index_root, p)   # index.html から見た相対パス
    except Exception:
        logo_url = ""
        
    # 背景スタイル（見つからなければ単色）
    bg_style = f"background: #f5f5f5 url('{bg_url}') repeat;" if bg_url else "background: #f5f5f5;"

    # --- index.html 本体 ---
    index = f"""<!doctype html>
    <html lang="ja"><head>
    <meta charset="utf-8"><title>Index</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    :root {{ --line:#ff7b7b; }}
    body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Noto Sans JP",Helvetica,Arial,sans-serif;line-height:1.7;margin:24px; {bg_style} }}
    
    /* 固定ヘッダー */
    #appHeader{{
    position:fixed; inset:0 0 auto 0; height:84px;
    display:flex; align-items:center; gap:16px;
    padding:50px 20px; box-sizing:border-box;
    background:rgba(15,18,22,.92); backdrop-filter:saturate(160%) blur(6px);
    border-bottom:1px solid rgba(255,255,255,.08); z-index:2000;
    }}
    #appHeader .brand{{display:flex; align-items:center; gap:12px;}}
    #appHeader .brand img{{height:100px; width:auto; display:block}}
    #appHeader .search{{flex:1; display:flex; justify-content:center;}}
    #appHeader .search input{{
    width:100%; max-width:760px; padding:.6rem .9rem; border:3px solid #e3edff;
    border-radius:10px; background:#3b4046; color:#e5e7eb; outline:none;
    }}
    #appHeader .search input::placeholder{{color:#9aa0a6}}
    #appHeader .cta{{margin-left:auto;}}
    #appHeader .cta a{{
    display:inline-flex; align-items:center; justify-content:center;
    height:40px; padding:0 16px; border-radius:10px;
    background:#16a34a; color:#fff; text-decoration:none; font-weight:700;
    box-shadow:0 2px 10px rgba(22,163,74,.35);
    }}
    #appHeader .cta a:hover{{filter:brightness(1.07)}}
    
    h1{{margin:0 0 12px}}
    /* details/summary 見出し */
    details.sec{{margin:18px 0}}
    summary{{list-style:none;cursor:pointer;user-select:none;display:flex;align-items:center;
            gap:.6rem;font-size:1.6rem;font-weight:800;margin:40px 0 4px}}
    summary::-webkit-details-marker{{display:none}}
    /* 赤ライン */
    summary {{border-bottom: 4px solid var(--line);padding-bottom: 6px;margin-bottom: 6px;}}
    /* 旧・赤丸（pokeball.png がないときに表示） */
    summary .dot{{display:inline-block;width:18px;height:18px;border-radius:50%;border:3px solid var(--line)}}
    /* ボール画像 */
    summary .poke{{width:22px;height:22px;display:inline-block;vertical-align:middle}}
    ul{{list-style:circle}}
    ul.root>li{{margin:.25rem 0}}
    #q{{width:100%;max-width:520px;padding:.5rem .7rem;border:1px solid #ccc;border-radius:8px;margin-bottom:16px}}
    li.hidden{{display:none}}
    </style>
    <script>
    function filter(){{
    const q = document.getElementById('q').value.trim().toLowerCase();
    const anyQuery = !!q;
    document.querySelectorAll('details.sec').forEach(sec => {{
        let any=false;
        sec.querySelectorAll('li').forEach(li => {{
        const t = li.textContent.toLowerCase();
        const hide = anyQuery && !t.includes(q);
        li.classList.toggle('hidden', hide);
        if(!hide) any=true;
        }});
        // ヒットがあるセクションだけ表示 検索中は自動で開く
        sec.style.display = (anyQuery && !any) ? 'none' : '';
        if(anyQuery && any) sec.open = true;
    }});
    }}
    </script>
    </head>
    
    <body>
    <header id="appHeader" role="banner" aria-label="Klefki Conflu Header">
    <div class="brand">
        <img src="{logo_url}" alt="Klefki Conflu">
    </div>
    <div class="search">
        <input id="q" type="search" placeholder="検索キーワード…" oninput="filter()" aria-label="ページを検索">
    </div>
    <div class="cta">
        <a href="https://github.com/Sadc2h4/Klefki-Conflu" target="_blank" rel="noopener">About this Application</a>
    </div>
    </header>
    <h1>ページ目次</h1>
    {''.join(sections_html)}
    </body></html>"""
    index = index.replace("{logo_url}", logo_url)
    (index_root / "index.html").write_text(index, encoding="utf-8")

# ループ index/total に応じて start〜end の間で進捗を更新
def _step_progress(start: float, end: float, index: int, total: int, progress_cb) -> None:
    if progress_cb is None or total <= 0:
        return
    ratio = index / total
    value = start + (end - start) * ratio
    progress_cb(value)   

def generate_html_from_xml_root(input_path: Path, out_html_root: Path, out_root: Path, log_box, progress_cb=None):
    
    # Zip/フォルダどちらでも HTML を生成
    if input_path.is_file() and input_path.suffix.lower()==".zip":
        log_append(log_box, "=== XML→HTML 生成 (ZIPダイレクト読込) ===")
        with zipfile.ZipFile(input_path) as z:
            ent = next((n for n in z.namelist() if n.endswith("/entities.xml") or n=="entities.xml"), None)
            if not ent:
                raise RuntimeError("Zip内に entities.xml が見つかりません．")
            pages_map, bodies = _parse_pages_and_bodies_from_entities_bytes(z.read(ent))
    else:
        log_append(log_box, "=== XML→HTML 生成 (フォルダ) ===")
        entities_list = list(Path(input_path if input_path.is_dir() else input_path.parent).rglob("entities.xml"))
        if not entities_list:
            raise RuntimeError("entities.xml が見つかりません．")
        pages_map, bodies = {}, {}
        for ent in entities_list:
            p, b = _parse_pages_and_bodies_from_entities_bytes(Path(ent).read_bytes())
            pages_map.update(p)
            bodies.update(b)

    global SIDEBAR_HTML_ROOT, SIDEBAR_ITEMS, SIDEBAR_EMPTY_PAGES
    SIDEBAR_HTML_ROOT = out_html_root
    SIDEBAR_ITEMS = []  # ★毎回リセットさせる
    SIDEBAR_EMPTY_PAGES = set()

    # 1周目：全ページのチェインとパスを収集
    page_entries: list[tuple[str, list[str], Path, str]] = []
    total_pages = len(pages_map)
    total_entries = len(page_entries)
    log_append(log_box, f"[Sub process 1] Collect all page chains and paths  total：{total_pages}")
    for i, (pid, meta) in enumerate(pages_map.items(), start=1):
        chain = _build_chain(pid, pages_map)
        page_dir, page_title = _dir_for_chain(out_html_root, chain)
        page_entries.append((pid, chain, page_dir, page_title))
        _step_progress(40, 45, i, total_pages, progress_cb) 
    log_append(log_box, f"=== DONE === Sub process 1 complete")

    # 2周目：サイドバー用のアイテムを先に全て登録
    log_append(log_box, f"[Sub process 2] Sidebar Item Registration  total：{total_pages}")
    for i, (pid, chain, page_dir, page_title) in enumerate(page_entries, start=1):
        page_path = page_dir / f"{page_title}.html"
        SIDEBAR_ITEMS.append((chain, page_path))
        _step_progress(45, 50, i, total_pages, progress_cb)
    log_append(log_box, f"=== DONE === Sub process 2 complete")

    # 3周目：全ページを確認して空白ページの事前判定
    log_append(log_box, f"[Sub process 3] Search for blank pages  total：{total_pages}")
    page_empty_map: dict[Path, bool] = {}
    for i, (pid, chain, page_dir, page_title) in enumerate(page_entries, start=1):
        page_path = page_dir / f"{page_title}.html"
        storage_html = bodies.get(pid, "")

        is_empty = _is_blank_storage_html(storage_html)

        prev = page_empty_map.get(page_path)
        if prev is None:
            page_empty_map[page_path] = is_empty
        else:
            page_empty_map[page_path] = prev and is_empty
        _step_progress(50, 55, i, total_pages, progress_cb)
    log_append(log_box, f"=== DONE === Sub process 3 complete")

    SIDEBAR_EMPTY_PAGES = {p for p, empty in page_empty_map.items() if empty}

    # 4周目：各ページのHTMLを書き出し（サイドバーは全ページ分を見て生成）
    attach_index = build_attachment_index(out_root) # 添付インデックスを1回だけ構築させる
    made = 0
    for i, (pid, chain, page_dir, page_title) in enumerate(page_entries, start=1):
        page_path = page_dir / f"{page_title}.html"
        storage_html = bodies.get(pid, "")
        # このページのフォルダを基準に共通アイコンの相対パスを事前解決
        icons = {}
        try:
            p = _ensure_bg_image(out_root)
            if p:
                icons["bg"] = _rel_href_from(page_dir, p)
        except Exception:
            pass
        try:
            p = _ensure_exe_icon(out_root)
            if p: icons["exe_icon"] = _rel_href_from(page_dir, p)
        except Exception:
            pass
        try:
            p = _ensure_pokeball(out_root)
            if p: icons["pokeball"] = _rel_href_from(page_dir, p)
        except Exception:
            pass
        try:
            p = _ensure_empty_icon(out_root)
            if p: icons["empty_icon"] = _rel_href_from(page_dir, p)
        except Exception:
            pass
        try:
            p = _ensure_footer_gori(out_root)
            if p:icons["back_top_img"] = _rel_href_from(page_dir, p)
        except Exception:
            pass

        html = confluence_storage_to_html(
            storage_html,
            chain,
            out_html_root,
            out_root,
            attach_index=attach_index,
            resolved_icons=icons,
        )
        page_dir.mkdir(parents=True, exist_ok=True)
        page_path.write_text(html, encoding="utf-8")

        log_append(log_box, f"[HTML] {'/'.join(chain)}/{page_title}.html")
        pump_gui(log_box)
        made += 1

        _step_progress(55, 99, i, total_pages, progress_cb)  

    _write_index_html(out_root, out_html_root, pages_map, link_prefix=f"{HTML_DIR_NAME}/")
    log_append(log_box, f"=== XML→HTML 生成 完了（ページ {made} 件／index.html 生成） ===")



def _write_docx(page_dir: Path, page_title: str, html: str):
    doc = docx.Document()
    text = _html_to_plaintext(html)
    for line in text.splitlines():
        doc.add_paragraph(line if line.strip() else "")
    page_dir.mkdir(parents=True, exist_ok=True)
    doc.save(str(page_dir / f"{page_title}.docx"))

def _generate_docx_from_zip(zip_path: Path, dst_root: Path, log_box: tk.Text):
    if not HAS_DOCX:
        raise RuntimeError("python-docx が未導入のため Word 出力できません．`pip install python-docx` を実行してください．")
    with zipfile.ZipFile(zip_path) as z:
        ent_name = None
        for name in z.namelist():
            if name.endswith("/entities.xml") or name == "entities.xml":
                ent_name = name; break
        if not ent_name:
            raise RuntimeError("Zip内に entities.xml が見つかりません．")
        ent_bytes = z.read(ent_name)
        pages_map, body_html = _parse_pages_parent_from_entities_xml_bytes(ent_bytes)
        made = 0
        for pid, meta in pages_map.items():
            titles = _build_dir_chain_for_page(pid, pages_map)
            page_dir, page_title = _dir_for_titles(dst_root, titles)
            html = body_html.get(pid, "")
            _write_docx(page_dir, page_title, html)
            log_append(log_box, f"[DOCX] {'/'.join(titles)}/{page_title}.docx"); made += 1
        log_append(log_box, f"=== XML→Word 生成 完了（Zip, ページ {made} 件）===")

def _generate_docx_from_folder(src_root: Path, dst_root: Path, log_box: tk.Text):
    if not HAS_DOCX:
        raise RuntimeError("python-docx が未導入のため Word 出力できません．`pip install python-docx` を実行してください．")
    entities_list = list(src_root.rglob("entities.xml"))
    if not entities_list:
        raise RuntimeError("entities.xml が見つかりません．XML/Zip展開ルートを確認してください．")
    pages_dir = src_root / "pages"
    pages_from_pages = _collect_pages_parent_from_pages_dir(pages_dir)
    pages_map_all = {}; body_html_all = {}
    for ent in entities_list:
        with open(ent, "rb") as fp:
            pages_map, body_html = _parse_pages_parent_from_entities_xml_bytes(fp.read())
        pages_map_all.update(pages_map); body_html_all.update(body_html)
    for k, v in pages_from_pages.items():
        base = pages_map_all.get(k, {"title": v.get("title",""), "parentId": v.get("parentId","")})
        if v.get("title"): base["title"] = v["title"]
        if v.get("parentId"): base["parentId"] = v["parentId"]
        pages_map_all[k] = base
        if v.get("html"): body_html_all[k] = v["html"]
    made = 0
    for pid, meta in pages_map_all.items():
        titles = _build_dir_chain_for_page(pid, pages_map_all)
        page_dir, page_title = _dir_for_titles(dst_root, titles)
        html = body_html_all.get(pid, "")
        _write_docx(page_dir, page_title, html)
        log_append(log_box, f"[DOCX] {'/'.join(titles)}/{page_title}.docx"); made += 1
    log_append(log_box, f"=== XML→Word 生成 完了（Folder, ページ {made} 件） ===")

def generate_docx_from_xml_root(input_path: Path, out_docx_root: Path, log_box: tk.Text):
    if input_path.is_file() and input_path.suffix.lower() == ".zip":
        log_append(log_box, f"=== XML→Word 生成 (ZIP直読) ===")
        _generate_docx_from_zip(input_path, out_docx_root, log_box)
    else:
        src_root = input_path if input_path.is_dir() else input_path.parent
        log_append(log_box, f"=== XML→Word 生成 (フォルダ) ===")
        _generate_docx_from_folder(src_root, out_docx_root, log_box)

# ----------------------------------------------------------------------------------
# ロゴ表示
# ----------------------------------------------------------------------------------
    # コマンド画面にロゴをプリントする処理
    # https://patorjk.com/software/taag/#p=display&v=0&f=Ogre&t=Otoware%20Spamton
def ComandView_Logo_Print(log_box):
    log_append(log_box ," _  ___   ___ ___ _  ___    ___ __  __  _ ___ _  _  _    .--.")
    log_append(log_box ,"| |/ / | | __| __| |/ / |  / _//__\|  \| | __| || || |  / .-.'---------.")
    log_append(log_box ,"|   <| |_| _|| _||   <| | | \_| \/ | | ' | _|| || \/ |  \ '-'.-A--AA-A'")
    log_append(log_box ,"|_|\_\___|___|_| |_|\_\_|  \__/\__/|_|\__|_| |___\__/    '--'")
    log_append(log_box ,"")
    log_append(log_box ,"Created by               : Sad (Twitter : @Tower_16_C2H4)")
    log_append(log_box ,"Version                  : 1.40")
    log_append(log_box ,"Development environment  : Python3.10.8 ")
    log_append(log_box ,"Operating environment    : Windows10 , Windows11")
    log_append(log_box ,"")
    log_append(log_box ,"=== Choose Confluence Buckup File (Zip file) ===")
    log_append(log_box ,"")
    pass

# ----------------------------------------------------------------------------------
# DnDユーティリティ
# ----------------------------------------------------------------------------------
def _split_dropped_files(data: str) -> list[str]:
    # tkinterdnd2 の <<Drop>> event.data を Windows/複数選択にも対応して分解
    # 例: '{C:\\Program Files\\a b.zip}' '{D:\\x y\\z.zip}'
    if not data:
        return []
    items = []
    buf = ""
    brace = 0
    for ch in data:
        if ch == "{":
            brace += 1
            if brace == 1: 
                buf = ""
                continue
        if ch == "}":
            brace -= 1
            if brace == 0:
                items.append(buf)
                buf = ""
                continue
        if brace == 0 and ch == " ":
            if buf:
                items.append(buf); buf = ""
            continue
        buf += ch
    if buf:
        items.append(buf)
    # バックスラッシュ正規化
    return [s.strip().strip('"') for s in items if s.strip()]

# ----------------------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------------------
BaseTk = TkinterDnD.Tk if HAS_DND else tk.Tk

class App(BaseTk):    
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE); self.minsize(740, 440)

        icon_png = _find_resource_file(EXE_ICON_FILE)  # 例: exe_icon.png を探す
        ico_file = _find_resource_file("exe_logo.ico") # .ico が別にあるならそれを探す

        if ico_file and Path(ico_file).exists():
            try:
                self.iconbitmap(default=str(ico_file))   # ← BaseTk ではなく self を使う！
            except Exception:
                pass  # 失敗しても致命的ではないので握りつぶす
            

        frm = ttk.Frame(self, padding=10); frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Select Confluence Zip").grid(row=0, column=0, sticky="w")
        self.in_var = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=self.in_var, width=78)
        ent.grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="参照", command=self.pick_input).grid(row=0, column=2, sticky="e")

        info = ttk.Label(
            frm,
            text=(
                "Zip を選ぶか,  ウィンドウ/テキスト欄へドラッグ＆ドロップしてください. 出力先は自動作成（YYYYMMDDhhmmss_SpaceKey）．\n"
                "Select zip or drag and drop into the window/text field."
            ),
            justify="left"
        )
        info.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6,0))

        opt = ttk.Frame(frm); opt.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4,0))
        self.rehome_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="復元後，ページタイトルで再配置（添付ファイル配下）", variable=self.rehome_var).pack(side="left")

        # 出力形式トグル（排他）
        self.md_var = tk.BooleanVar(value=True)
        self.docx_var = tk.BooleanVar(value=False)
        def on_md_toggle():
            if self.md_var.get(): self.docx_var.set(False)
        def on_docx_toggle():
            if self.docx_var.get(): self.md_var.set(False)
            if self.docx_var.get() and not HAS_DOCX:
                log_append(self.log, "※ Word出力には python-docx が必要です．`pip install python-docx`\n")
        # ttk.Checkbutton(opt, text="mdファイルとして復元（md_pages）", variable=self.md_var, command=on_md_toggle).pack(side="left", padx=(12,0))
        # ttk.Checkbutton(opt, text="Wordファイルとして復元（word_pages）", variable=self.docx_var, command=on_docx_toggle).pack(side="left", padx=(12,0))

        # allow_dir は内部的にfalseのまま所持（validateで参照するため）
        self.allow_dir_var = tk.BooleanVar(value=False)

        btns = ttk.Frame(frm); btns.grid(row=3, column=0, columnspan=3, sticky="w", pady=(8,0))
        self.btn_dry = ttk.Button(btns, text="①ドライラン（計画表示）", command=lambda: self.run(True))
        self.btn_run = ttk.Button(btns, text="②実行（変更を反映）", command=lambda: self.run(False))
        self.btn_dry.pack(side="left"); self.btn_run.pack(side="left", padx=(8,0))

        # ログ（黒地・等幅）＋ 縦スクロールバー
        self.log = tk.Text(frm, height=28, wrap="none")
        self.log.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(8, 0))

        # 縦スクロールバーを右側に配置
        self.log_scroll = ttk.Scrollbar(frm, orient="vertical", command=self.log.yview)
        self.log_scroll.grid(row=4, column=2, sticky="ns", pady=(8, 0))

        # Text と Scrollbar を連動
        self.log.config(state="disabled", yscrollcommand=self.log_scroll.set)

        # ログ用フォントを固定（TkFixedFont ベース）
        try:
            cmd_font = tkfont.nametofont("TkFixedFont")
            cmd_font.configure(size=11)
        except Exception:
            # 念のためのフォールバック
            cmd_font = ("Courier New", 11)

        self.log.configure(
            bg="#111111", fg="#EEEEEE", insertbackground="#FFFFFF",
            selectbackground="#444444", selectforeground="#FFFFFF", font=cmd_font
        )
        self.log.tag_configure("OK",   foreground="#9AE66E")
        self.log.tag_configure("WARN", foreground="#FFD166")
        self.log.tag_configure("ERROR",foreground="#FF6B6B")
        self.log.tag_configure("MOVE", foreground="#6EC1FF")
        self.log.tag_configure("INFO", foreground="#EEEEEE")

        # ログ行・中央列はリサイズで広がるように
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(4, weight=1)


        if not HAS_MAGIC:
            log_append(self.log, "※ python-magic 未導入．Windowsは `pip install python-magic-bin` 推奨．\n")
        if not HAS_H2T:
            log_append(self.log, "※ html2text 未導入．`pip install html2text` でMarkdown精度UP．\n")
        if not HAS_DOCX:
            log_append(self.log, "※ Word出力は python-docx が必要．`pip install python-docx`\n")
        if not HAS_DND:
            log_append(self.log, "※ tkinterdnd2 未導入のため DnD は無効です．`pip install tkinterdnd2`\n")

        # 進捗バー
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress = ttk.Progressbar(frm, variable=self.progress_var, maximum=100)
        self.progress.grid(row=5, column=0, columnspan=3, sticky="we", pady=(4, 0))

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(4, weight=1)

        # DnD登録
        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop_files)
            ent.drop_target_register(DND_FILES)
            ent.dnd_bind("<<Drop>>", self._on_drop_files)
            self.log.drop_target_register(DND_FILES)
            self.log.dnd_bind("<<Drop>>", self._on_drop_files)

        self.in_var.trace_add("write", lambda *args: self.validate_input())
        self.validate_input()
        ComandView_Logo_Print(self.log)

    # --- DnD handler ---
    def _on_drop_files(self, event):
        paths = _split_dropped_files(event.data)
        if not paths:
            return
        # ZIP優先で一つ拾う（複数投下時）
        zips = [p for p in paths if p.lower().endswith(".zip")]
        target = zips[0] if zips else paths[0]
        self.in_var.set(target)  # set -> validate_input() が呼ばれる

        # 入力が有効なら，そのまま ①ドライラン → ②本実行 を自動で実施
        if self.validate_input():
            # self.run(True)   # ①ドライラン
            # log_append(self.log, "\n--- ドライラン完了 本実行を開始 ---\n")
            self.run(False)  # ②実行

    def validate_input(self):
        p = Path(self.in_var.get().strip()) if self.in_var.get().strip() else None
        ok = False
        if p and p.exists():
            if p.is_file() and p.suffix.lower() == ".zip":
                ok = True
            elif self.allow_dir_var.get() and p.is_dir():
                ok = True
        state = "normal" if ok else "disabled"
        self.btn_dry.config(state=state); self.btn_run.config(state=state)
        return ok

    def pick_input(self):
        f = filedialog.askopenfilename(filetypes=[("ZIP files","*.zip"), ("All files","*.*")])
        if f:
            self.in_var.set(f); return
        if self.allow_dir_var.get():
            d = filedialog.askdirectory()
            if d: self.in_var.set(d)

    def _set_progress(self, value: float):
        # 進捗バーを更新（0〜100）
        try:
            self.progress_var.set(float(value))
            self.progress.update_idletasks()
        except Exception:
            pass

    def run(self, dry_run: bool):
        if not self.validate_input():
            messagebox.showwarning("注意", "入力が未指定です．")
            return

        # 実行中はボタンを無効化し，進捗リセット
        self.btn_dry.config(state="disabled")
        self.btn_run.config(state="disabled")
        self._set_progress(0)

        in_p = Path(self.in_var.get().strip())
        out_root = _build_auto_out_root(in_p, self.log)
        attach_root = out_root / ATT_DIR_NAME
        html_root = out_root / HTML_DIR_NAME

        if not dry_run:
            try:
                attach_root.mkdir(parents=True, exist_ok=True)
                html_root.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("エラー", f"出力先作成に失敗: {e}")
                # ボタン状態を元に戻して終了
                ok = self.validate_input()
                self.btn_dry.config(state="normal" if ok else "disabled")
                self.btn_run.config(state="normal" if ok else "disabled")
                return

        try:
            # 添付復元（ステップ1）0% ⇒ 20%
            if in_p.is_file() and in_p.suffix.lower() == ".zip":
                log_append(self.log, f"=== ZIP入力: {in_p.name} (dry_run={dry_run}) ===")
                process_zip(in_p, attach_root, dry_run, self.log, progress_cb=self._set_progress)
            else:
                log_append(self.log, f"=== フォルダ入力: {in_p} (dry_run={dry_run}) ===")
                process_folder(in_p, attach_root, dry_run, self.log, progress_cb=self._set_progress)
            self._set_progress(20)

            # ドライラン
            if dry_run:
                self._set_progress(100)
                log_append(self.log, f"=== DONE (Dry-Run) === 出力予定: {out_root}")
                return

            # 再配置（ステップ2）20% ⇒ 40%
            if self.rehome_var.get():
                log_append(self.log, f"=== Re-home by PageTitle 開始（{ATT_DIR_NAME} 配下） ===")
                rehome_by_page_title(attach_root, self.log, progress_cb=self._set_progress)
            self._set_progress(40)

            # HTML 出力（ステップ3）40% ⇒ 99%
            generate_html_from_xml_root(in_p, html_root, out_root, self.log, progress_cb=self._set_progress)

            # HTML完了0% ⇒ 100%
            self._set_progress(100)

            log_append(self.log, f"=== DONE === 出力: {out_root}")
            messagebox.showinfo("Process successful","出力が完了しました")

        except Exception as e:
            exc_type, exc_value, exc_tb = e.__traceback__.tb_frame, e, e.__traceback__
            line_number = exc_tb.tb_lineno
            messagebox.showerror("エラった", f"エラーが発生した行番号: {line_number}\n" + str(e))

        finally:
            # 実行完了後にボタン状態を戻す
            ok = self.validate_input()
            self.btn_dry.config(state="normal" if ok else "disabled")
            self.btn_run.config(state="normal" if ok else "disabled")

if __name__ == "__main__":
    App().mainloop()