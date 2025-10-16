# -*- coding: utf-8 -*-
# Version: 2025-09-06 Clean stable build — brand: Ψηφιακή Κατανομή Μαθητών Α' Δημοτικού
import re, os, json, importlib.util, datetime as dt, math, base64, unicodedata
from pathlib import Path
from io import BytesIO

# Project root directory
ROOT = Path(__file__).parent.resolve()

import streamlit as st
import pandas as pd

# --- Embedded logo fallback (base64) ---
import base64
from io import BytesIO
LOGO_B64 = "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"
LOGO_MIME = "image/png"

def _get_logo_bytes():
    """Return logo bytes: from file path if available, else from embedded base64."""
    path = None
    try:
        path = _find_logo_path()
    except Exception:
        path = None
    if path:
        try:
            return Path(path).read_bytes()
        except Exception:
            pass
    if LOGO_B64:
        try:
            return base64.b64decode(LOGO_B64)
        except Exception:
            return None
    return None

def _inject_floating_logo(width_px=62):
    """Render a floating logo at bottom-right that stays on screen while scrolling."""
    
    # Extra guard: if already authenticated and terms accepted, don't render
    try:
        if st.session_state.get("auth_ok") and st.session_state.get("accepted_terms"):
            return
    except Exception:
        pass
# Get bytes from file or embedded
    data = _get_logo_bytes()
    if not data:
        return
    b64 = base64.b64encode(data).decode('utf-8')
    mime = LOGO_MIME

    import streamlit as st
    st.markdown(f"""
<style>
#floating-logo {{
  position: fixed;
  left: 285px;
  bottom: 16px;
  z-index: 9999;
  opacity: 0.95;
  pointer-events: none;  /* don't block clicks */
}}
#floating-logo img {{
  width: {width_px}px;
  height: auto;
  filter: drop-shadow(0 1px 2px rgba(0,0,0,0.20));
  opacity: 0.92;
}}
@media (max-width: 768px) {{
  #floating-logo img {{ width: {max(72, int(0.85*width_px))}px; }}
  #floating-logo {{ left: 285px; bottom: 12px; }}
}}
</style>
<div id="floating-logo">
  <img src="data:{mime};base64,{b64}" alt="logo" />
</div>
""", unsafe_allow_html=True)

from PIL import Image, ImageDraw, ImageFont

# ---- Logo helpers ----
def _find_logo_path():
    from pathlib import Path as _P
    here = _P(__file__).parent
    candidates = [
        "logo_sidebar_preview_selected.png",
        "logo_lotus_lilac_sidebar100.png",
        "logo_lotus_lilac_original.png",
        "logo_lotus_lilac_header180.png",
        "logo_violet_white.png",
        "logo.png",
        "assets/logo.png",
        "lotus_appicon_white_1024.png",
        ]
    search_bases = [here, here / "assets", _P("/mnt/data")]
    # direct candidates
    for base in search_bases:
        for c in candidates:
            p = base / c
            if p.exists():
                return str(p)
    # wildcard fallback
    for base in search_bases:
        for p in base.glob("lotus*.png"):
            return str(p)
    return None

# ---- Compose logo with bottom caption overlay ----

def _make_logo_with_overlay(img_path, width=140, text="“No man is an island”"):
    try:
        im = Image.open(img_path).convert("RGBA")
    except Exception:
        return None
    # Resize keeping aspect
    scale = width / im.width
    target_h = int(im.height * scale)
    im = im.resize((width, target_h), Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS)

    draw = ImageDraw.Draw(im, "RGBA")

    # Choose a calligraphic/bold-italic style if available; fallback to bold sans
    font_candidates = [
        ("DejaVuSans.ttf", 20),
        ("Arial.ttf", 20),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20),
    ]
    font = None
    for fname, fsize in font_candidates:
        try:
            font = ImageFont.truetype(fname, fsize)
            break
        except Exception:
            continue
    if font is None:
        font = ImageFont.load_default()

    # Auto-fit font size to width with margins
    max_w = int(width * 0.92)
    fsize = getattr(font, "size", 18)
    # try decreasing size until it fits
    while True:
        bbox = draw.textbbox((0,0), text, font=font, stroke_width=2)
        tw, th = bbox[2], bbox[3]
        if tw <= max_w or fsize <= 11:
            break
        fsize -= 1
        try:
            font = ImageFont.truetype(font.path, fsize) if hasattr(font, "path") else ImageFont.truetype("DejaVuSans.ttf", fsize)
        except Exception:
            font = ImageFont.truetype("DejaVuSans.ttf", fsize)

    # Position near bottom
    bbox = draw.textbbox((0,0), text, font=font, stroke_width=2)
    tw, th = bbox[2], bbox[3]
    tx = max(0, (width - tw)//2)
    ty = target_h - th - 6

    # Draw with strong black stroke for contrast (no background strip)
    draw.text((tx, ty), text, font=font, fill=(255,255,255,255), stroke_width=2, stroke_fill=(0,0,0,220))

    return im

# ---------------------------
# Ρυθμίσεις σελίδας / Branding
# ---------------------------

# Load logo once (if available)
_logo_path = _find_logo_path()
_logo_img = None
if _logo_path:
    try:
        _logo_img = Image.open(_logo_path)
    except Exception:
        _logo_img = None

# Load logo image for page icon (from bytes or path)
_logo_bytes = _get_logo_bytes()
_logo_img = None
if _logo_bytes:
    try:
        _logo_img = Image.open(BytesIO(_logo_bytes))
    except Exception:
        _logo_img = None
st.set_page_config(page_title="Ψηφιακή Κατανομή Μαθητών Α' Δημοτικού", page_icon=_logo_img if _logo_img else "🧩", layout="wide")

# --- Κεφαλίδα σελίδας ---
st.title("Ψηφιακή Κατανομή Μαθητών Α' Δημοτικού")

# --- Υπότιτλος με εικονίδιο λωτού ---
try:
    import base64
    _logo_inline_bytes = _get_logo_bytes()
    _logo_inline_b64 = base64.b64encode(_logo_inline_bytes).decode("ascii") if _logo_inline_bytes else ""
except Exception:
    _logo_inline_b64 = ""
st.markdown(f"""
<div style="display:flex; align-items:center; gap:8px; opacity:0.85;">
  <span>«Για μια παιδεία που βλέπει το φώς σε όλα τα παιδιά»</span>
  <img src="data:image/png;base64,{_logo_inline_b64}" alt="lotus" style="width:18px; height:auto; margin-top:-2px; " />
</div>
""" , unsafe_allow_html=True)

# Show floating logo only on the initial screen (before auth + terms)
try:
    _auth = bool(st.session_state.get("auth_ok", False))
    _terms = bool(st.session_state.get("accepted_terms", False))
except Exception:
    _auth, _terms = (False, False)
if not (_auth and _terms):
    _inject_floating_logo(width_px=62)

def _load_module(name: str, file_path: Path):
    spec = importlib.util.spec_from_file_location(name, str(file_path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore
    return mod

def _read_file_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()

def _timestamped(base: str, ext: str) -> str:
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    import re as _re
    safe = _re.sub(r"[^A-Za-z0-9_\-\.]+", "_", base)
    return f"{safe}_{ts}{ext}"

def _find_latest_step6():
    """Εντοπίζει το πιο πρόσφατο αρχείο STEP1_6_PER_SCENARIO_*.xlsx στον φάκελο της εφαρμογής."""
    try:
        candidates = sorted((p for p in ROOT.glob("STEP1_6_PER_SCENARIO*.xlsx") if p.is_file()),
                            key=lambda p: p.stat().st_mtime,
                            reverse=True)
        return candidates[0] if candidates else None
    except Exception:
        return None

def _check_required_files(paths):
    missing = [str(p) for p in paths if not p.exists()]
    return missing

def _inject_logo(logo_bytes: bytes, width_px: int = 140, mime: str = "image/png"):
    b64 = base64.b64encode(logo_bytes).decode("ascii")
    html = f"""
    <div style="position: fixed; bottom: 38px; right: 38px; z-index: 1000;">
        <img src="data:{mime};base64,{b64}" style="width:{width_px}px; height:auto; opacity:0.95; border-radius:12px;" />
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def _restart_app():
    # Καθάρισε session_state κλειδιά
    for k in list(st.session_state.keys()):
        if k.startswith("uploader_") or k in ("auth_ok","accepted_terms","app_enabled","last_final_path"):
            try:
                del st.session_state[k]
            except Exception:
                pass
    # Καθάρισε caches
    try:
        st.cache_data.clear()
    except Exception:
        pass
    try:
        st.cache_resource.clear()
    except Exception:
        pass
    # ΔΙΑΓΡΑΦΗ παραγόμενων αρχείων για πλήρη καθαρισμό
    try:
        for pat in ("STEP7_FINAL_SCENARIO*.xlsx", "STEP1_6_PER_SCENARIO*.xlsx", "INPUT_STEP1*.xlsx"):
            for f in ROOT.glob(pat):
                try:
                    f.unlink()
                except Exception:
                    pass
    except Exception:
        pass
    st.rerun()

def _terms_md():
    return """
**Υποχρεωτική Αποδοχή Όρων Χρήσης**  
Χρησιμοποιώντας την εφαρμογή δηλώνετε ότι:  
- Δεν τροποποιείτε τη λογική των αλγορίθμων και δεν αναδιανέμετε τα αρχεία χωρίς άδεια.  
- Αναλαμβάνετε την ευθύνη για την ορθότητα των εισαγόμενων δεδομένων.  
- Η εφαρμογή παρέχεται «ως έχει», χωρίς εγγύηση για οποιαδήποτε χρήση.  

**Πνευματικά Δικαιώματα & Νομική Προστασία**  
© 2025 Γιαννίτσαρου Παναγιώτα — Όλα τα δικαιώματα διατηρούνται.  
Για άδεια χρήσης/συνεργασίες: *panayiotayiannitsarou@gmail.com*.
"""

def _story_md():
    return """
**Η εφαρμογή αυτή γεννήθηκε από μια εσωτερική ανάγκη:** να θυμίσει ότι **κανένα παιδί δεν πρέπει να μένει στο περιθώριο**. Το παιδί δεν είναι απλώς ένα όνομα σε λίστα. Είναι παρουσία, ψυχή, μέλος μιας ομάδας. Μια απερίσκεπτη κατανομή ή ένας λανθασμένος παιδαγωγικός χειρισμός μπορεί να ταράξει την εύθραυστη ψυχική ισορροπία ενός παιδιού — και μαζί της, την ηρεμία μιας οικογένειας.

Όπως έγραψε ο John Donne, «Κανένας άνθρωπος δεν είναι νησί» (“No man is an island”)¹: κανείς δεν υπάρχει απομονωμένος· ό,τι συμβαίνει σε έναν, αφορά όλους. Είμαστε μέρος ενός ευρύτερου συνόλου· η μοίρα, η χαρά ή ο πόνος του άλλου μας αγγίζουν, γιατί είμαστε συνδεδεμένοι.

Στο σχολείο αυτό γίνεται πράξη: κάθε απόφαση είναι πράξη παιδαγωγικής ευθύνης. Ένα πρόγραμμα κατανομής δεν είναι ποτέ απλώς ένα τεχνικό εργαλείο. Είναι **έκφραση παιδαγωγικής ευθύνης** και **κοινωνικής ευαισθησίας**. Δεν είναι μόνο αλγόριθμος· είναι έκφραση κοινωνικής ευαισθησίας και εμπιστοσύνης στο μέλλον — των παιδιών και της κοινωνίας.

*¹ Η φράση υπογραμμίζει ότι κανείς δεν είναι πλήρως ανεξάρτητος.*

— Με εκτίμηση,  
**Γιαννίτσαρου Παναγιώτα**

**Απόσπασμα από τον John Donne**
> No man is an island,
> Entire of itself;
> Every man is a piece of the continent,
> A part of the main.
> If a clod be washed away by the sea,
> Europe is the less,
> As well as if a promontory were,
> As well as if a manor of thy friend’s
> Or of thine own were.
> Any man's death diminishes me,
> Because I am involved in mankind.
> And therefore never send to know for whom the bell tolls;
> It tolls for thee.

— *John Donne*
"""

# ---------------------------
# Αρχεία που δεν αλλάζουμε (modules 1→7)
# ---------------------------
REQUIRED = [
    ROOT / "export_step1_6_per_scenario.py",
    ROOT / "step1_immutable_ALLINONE.py",
    ROOT / "step_2_helpers_FIXED.py",
    ROOT / "step_2_zoiroi_idiaterotites_FIXED_v3_PATCHED.py",
    ROOT / "step3_amivaia_filia_FIXED.py",
    ROOT / "step4_corrected.py",
    ROOT / "step5_enhanced.py",
    ROOT / "step6_compliant.py",
    ROOT / "step7_fixed_final.py",
]

# ---------------------------
# Sidebar: πρόσβαση, όροι, λογότυπο
# ---------------------------
with st.sidebar:

    # (Logo in sidebar removed — using floating bottom-right logo)

    st.header("🔐 Πρόσβαση & Ρυθμίσεις")

    # 1) ΠΑΝΤΑ ΠΡΩΤΟ: πεδίο κωδικού
    if "auth_ok" not in st.session_state:
        st.session_state.auth_ok = False
    pwd = st.text_input("Κωδικός πρόσβασης", type="password")
    if pwd:
        st.session_state.auth_ok = (pwd.strip() == "katanomi2025")
        if not st.session_state.auth_ok:
            st.error("Λανθασμένος κωδικός.")

    # 2) Έπειτα: αποδοχή όρων (ορατή χωρίς προϋπόθεση)
    if "accepted_terms" not in st.session_state:
        st.session_state.accepted_terms = False
    st.session_state.accepted_terms = st.checkbox(
        "✅ Αποδέχομαι τους Όρους Χρήσης",
        value=st.session_state.get("accepted_terms", False)
    )

    # 3) Κάτω: expander με Όρους & Πνευματικά
    with st.expander("📄 Όροι Χρήσης & Πνευματικά Δικαιώματα", expanded=False):
        st.markdown(_terms_md())
    
    # — Κουμπί: Ιστορία Δημιουργίας & Πηγή Έμπνευσης (χωρίς δεύτερη見见见 επικεφαλίδα)
    if "show_story" not in st.session_state:
        st.session_state.show_story = False
    if st.button("🧭 Η Ιστορία της Δημιουργίας & Πηγή Έμπνευσης", use_container_width=True, key="btn_story"):
        st.session_state.show_story = not st.session_state.show_story
    if st.session_state.show_story:
        st.markdown(_story_md())
st.divider()
    # Λογότυπο: απενεργοποιημένο κατ' απαίτηση

# ---------------------------
# Πύλες προστασίας
# ---------------------------
if not st.session_state.auth_ok:
    st.warning("🔐 Εισάγετε τον σωστό κωδικό για πρόσβαση (αριστερά).")
    st.stop()

if not st.session_state.accepted_terms:
    st.warning("✅ Για να συνεχίσετε, αποδεχθείτε τους Όρους Χρήσης (αριστερά).")
    st.stop()

# ---------------------------
st.subheader("📦 Έλεγχος αρχείων")
missing = _check_required_files(REQUIRED)
if missing:
    st.error("❌ Λείπουν αρχεία:\n" + "\n".join(f"- {m}" for m in missing))
else:
    st.success("✅ Όλα τα απαραίτητα αρχεία βρέθηκαν.")

st.divider()

# ---------------------------
# 🚀 Εκτέλεση Κατανομής
# ---------------------------
st.header("🚀 Εκτέλεση Κατανομής")

up_all = st.file_uploader("Ανέβασε αρχικό Excel (για 1→7)", type=["xlsx"], key="uploader_all")
colA, colB, colC = st.columns([1,1,1])
with colA:
    pick_step4_all = st.selectbox("Κανόνας επιλογής στο Βήμα 4", ["best", "first", "strict"], index=0, key="pick_all")
with colB:
    final_name_all = st.text_input("Όνομα αρχείου Τελικού Αποτελέσματος", value=_timestamped("STEP7_FINAL_SCENARIO", ".xlsx"))
with colC:
    if up_all is not None:
        try:
            df_preview = pd.read_excel(up_all, sheet_name=0)
            N = df_preview.shape[0]
            min_classes = max(2, math.ceil(N/25)) if N else 0
            st.metric("Μαθητές / Ελάχιστα τμήματα", f"{N} / {min_classes}")
        except Exception:
            st.caption("Δεν ήταν δυνατή η ανάγνωση για προεπισκόπηση.")

if st.button("🚀 ΕΚΤΕΛΕΣΗ ΚΑΤΑΝΟΜΗΣ", type="primary", use_container_width=True):
    if missing:
        st.error("Δεν είναι δυνατή η εκτέλεση: λείπουν modules.")
    elif up_all is None:
        st.warning("Πρώτα ανέβασε ένα Excel.")
    else:
        try:
            input_path = ROOT / _timestamped("INPUT_STEP1", ".xlsx")
            with open(input_path, "wb") as f:
                f.write(up_all.getbuffer())

            m = _load_module("export_step1_6_per_scenario", ROOT / "export_step1_6_per_scenario.py")
            s7 = _load_module("step7_fixed_final", ROOT / "step7_fixed_final.py")

            step6_path = ROOT / _timestamped("STEP1_6_PER_SCENARIO", ".xlsx")
            with st.spinner("Τρέχουν τα Βήματα 1→6..."):
                m.build_step1_6_per_scenario(str(input_path), str(step6_path), pick_step4=pick_step4_all)

            with st.spinner("Τρέχει το Βήμα 7..."):
                xls = pd.ExcelFile(step6_path)
                sheet_names = [s for s in xls.sheet_names if s != "Σύνοψη"]
                if not sheet_names:
                    st.error("Δεν βρέθηκαν sheets σεναρίων (εκτός από 'Σύνοψη').")
                else:
                    # Συγκεντρώνουμε ΟΛΑ τα σενάρια από ΟΛΑ τα φύλλα
                    candidates = []
                    import random as _rnd
                    for sheet in sheet_names:
                        df_sheet = pd.read_excel(step6_path, sheet_name=sheet)
                        scen_cols = [c for c in df_sheet.columns if re.match(r"^ΒΗΜΑ6_ΣΕΝΑΡΙΟ_\d+$", str(c))]
                        for col in scen_cols:
                            s = s7.score_one_scenario(df_sheet, col)
                            s["sheet"] = sheet
                            candidates.append(s)

                    if not candidates:
                        st.error("Δεν βρέθηκαν σενάρια Βήματος 6 σε κανένα φύλλο.")
                    else:
                        # 1) Κανόνας: ΠΡΩΤΑ min total_score, tie→ λιγότερα broken, μετά diff_population→diff_gender_total→diff_greek
                        pool_sorted = sorted(
                            candidates,
                            key=lambda s: (
                                int(s["total_score"]),
                                int(s.get("broken_friendships", 0)),
                                int(s["diff_population"]),
                                int(s["diff_gender_total"]),
                                int(s["diff_greek"]),
                                str(s["scenario_col"]),
                            )
                        )

                        head = pool_sorted[0]
                        ties = [s for s in pool_sorted if (
                            int(s["total_score"]) == int(head["total_score"]) and
                            int(s.get("broken_friendships", 0)) == int(head.get("broken_friendships", 0)) and
                            int(s["diff_population"]) == int(head["diff_population"]) and
                            int(s["diff_gender_total"]) == int(head["diff_gender_total"]) and
                            int(s["diff_greek"]) == int(head["diff_greek"])
                        )]

                        _rnd.seed(42)
                        best = _rnd.choice(ties) if len(ties) > 1 else head

                        winning_sheet = best["sheet"]
                        winning_col = best["scenario_col"]
                        final_out = ROOT / final_name_all

                        full_df = pd.read_excel(step6_path, sheet_name=winning_sheet).copy()
                        with pd.ExcelWriter(final_out, engine="xlsxwriter") as w:
                            labels = sorted(
                                [str(v) for v in full_df[winning_col].dropna().unique() if re.match(r"^Α\d+$", str(v))],
                                key=lambda x: int(re.search(r"\d+", x).group(0))
                            )
                            for lab in labels:
                                sub = full_df.loc[full_df[winning_col] == lab, ["ΟΝΟΜΑ", winning_col]].copy()
                                sub = sub.rename(columns={winning_col: "ΤΜΗΜΑ"})
                                sub.to_excel(w, index=False, sheet_name=str(lab))

                        st.session_state["last_final_path"] = str(final_out.resolve())
                        st.session_state["last_step6_path"] = str(step6_path)
                        st.session_state["last_winning_sheet"] = str(winning_sheet)
                        st.session_state["last_winning_col"] = str(winning_col)

                        st.success(f"✅ Ολοκληρώθηκε. Νικητής: φύλλο {winning_sheet} — στήλη {winning_col}")
                        st.download_button(
                            "⬇️ Κατέβασε Τελικό Αποτέλεσμα (1→7)",
                            data=_read_file_bytes(final_out),
                            file_name=final_out.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.caption("ℹ️ Το αρχείο αποθηκεύτηκε και θα χρησιμοποιηθεί **αυτόματα** από τα «📊 Στατιστικά».")
# >>> INLINE_STEP8_BLOCK_START <<<
# Step 8 UI block: requires step8_fixed_final.UnifiedProcessor and uses BOTH:
#   (A) Αρχικό Μητρώο (source master)
#   (B) Βήμα 7 — Τελικό Αποτέλεσμα (multi-sheet allocation)
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

try:
    from step8_fixed_final import UnifiedProcessor
except Exception as _e:
    UnifiedProcessor = None
    _IMPORT_ERROR = _e
else:
    _IMPORT_ERROR = None

def _save_upload(uploaded, prefix):
    path = Path("/mnt/data") / f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    path.write_bytes(uploaded.getvalue())
    return str(path)

def _validate_step7_excel(path: str):
    """Basic validation: must have at least one sheet with ΟΝΟΜΑ column."""
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        return False, f"Δεν μπορώ να διαβάσω το αρχείο Βήμα 7: {e}", {}
    sheet_names = xls.sheet_names
    if not sheet_names:
        return False, "Το αρχείο Βήμα 7 δεν έχει φύλλα.", {}
    summary = {}
    for s in sheet_names:
        try:
            df = pd.read_excel(path, sheet_name=s)
            cols = [str(c).strip() for c in df.columns]
            has_name = "ΟΝΟΜΑ" in cols
            summary[s] = {"rows": int(df.shape[0]), "has_ΟΝΟΜΑ": bool(has_name)}
        except Exception as e:
            summary[s] = {"rows": 0, "has_ΟΝΟΜΑ": False, "error": str(e)}
    if not any(v.get("has_ΟΝΟΜΑ") for v in summary.values()):
        return False, "Κανένα φύλλο στο αρχείο Βήμα 7 δεν περιέχει στήλη 'ΟΝΟΜΑ'.", summary
    return True, "", summary

def _find_step7_from_session() -> str | None:
    p = st.session_state.get("last_final_path")
    if p and Path(p).exists():
        return p
    return None

def render_step8():
    st.header("🧠 Βήμα 8 — Fill + Optimize (με Βήμα 7)")
    if _IMPORT_ERROR:
        st.error(f"Δεν φορτώθηκε το module `step8_fixed_final`: {_IMPORT_ERROR}")
        return
    col1, col2 = st.columns(2)
    with col1:
        src_file = st.file_uploader("📥 Αρχικό Μητρώο (Source.xlsx)", type=["xlsx"], key="step8_src")
    with col2:
        default_step7 = _find_step7_from_session()
        tmpl_file = st.file_uploader(
            "📄 Βήμα 7 — Τελικό Αποτέλεσμα (Final Step7.xlsx)",
            type=["xlsx"], key="step8_tmpl",
            help="Αν έχεις ήδη τρέξει τα Βήματα 1→7 από αυτή την εφαρμογή, θα χρησιμοποιηθεί αυτόματα το τελευταίο παραγόμενο αρχείο.",
        )
        if (not tmpl_file) and default_step7:
            st.info(f"Χρήση αυτόματου Βήμα 7: **{Path(default_step7).name}**")
    max_iters = st.sidebar.slider("Βήμα 8 — Μέγιστες επαναλήψεις", 10, 500, 120, step=10, key="step8_iters")
    run = st.button("🚀 Fill + Optimize + Τελικό (Βήμα 8)", type="primary", use_container_width=True, key="btn_step8_run",
                    disabled=not (src_file and (tmpl_file or _find_step7_from_session())))
    if not run:
        return
    # Resolve paths
    src_path = _save_upload(src_file, "source") if src_file else None
    if tmpl_file:
        tmpl_path = _save_upload(tmpl_file, "step7_final")
    else:
        tmpl_path = _find_step7_from_session()
    if not (src_path and tmpl_path):
        st.error("Λείπουν αρχεία για το Βήμα 8.")
        return
    # Validate Step7
    ok, msg, summary = _validate_step7_excel(tmpl_path)
    with st.expander("Έλεγχος Βήμα 7 — Τελικού Αποτελέσματος", expanded=True):
        if summary:
            prev = pd.DataFrame.from_dict(summary, orient="index").reset_index().rename(columns={"index": "Φύλλο"})
            st.dataframe(prev, use_container_width=True, height=min(400, 60 + 28*len(prev)))
        if not ok:
            st.error(f"❌ {msg}")
            return
    filled_path = str(Path("/mnt/data") / f"filled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    final_path  = str(Path("/mnt/data") / f"final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    proc = UnifiedProcessor()
    with st.status("✍️ **Fill**: Συμπλήρωση φύλλων Βήμα 7 με δεδομένα από Αρχικό Μητρώο...", expanded=True) as s:
        proc.read_source_data(src_path)
        proc.fill_target_excel(tmpl_path, filled_path)
        s.update(label="✅ Fill ολοκληρώθηκε", state="complete")
    with open(filled_path, "rb") as f:
        st.download_button("⬇️ Κατέβασμα Filled (πριν το Optimize)", f.read(),
                           file_name=Path(filled_path).name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_filled_step8")
    with st.status("🧠 **Optimize**: Στοχευμένες ανταλλαγές...", expanded=True) as s:
        proc.load_filled_data(filled_path)
        swaps, spreads = proc.optimize(max_iterations=max_iters)
        proc.export_optimized_excel(swaps, spreads, final_path)
        s.update(label="✅ Optimize ολοκληρώθηκε", state="complete")
    st.subheader("📊 Τελικό — Διασπορές (spreads)")
    met_cols = st.columns(4)
    met_cols[0].metric("EP3 spread", spreads.get("ep3", 0))
    met_cols[1].metric("Boys spread", spreads.get("boys", 0))
    met_cols[2].metric("Girls spread", spreads.get("girls", 0))
    met_cols[3].metric("Greek (Ν) spread", spreads.get("greek_yes", 0))
    try:
        stats = proc._get_team_stats()
    except Exception:
        stats = None
    if stats:
        df_stats = (
            pd.DataFrame(stats).T
            .rename(columns={
                "boys": "Αγόρια", "girls": "Κορίτσια",
                "greek_yes": "Καλή Γνώση (Ν)", "greek_no": "Όχι Καλή (Ο)",
                "ep1": "Επίδοση 1", "ep2": "Επίδοση 2", "ep3": "Επίδοση 3"
            })
            .reset_index().rename(columns={"index": "Τμήμα"})
        )
        st.dataframe(df_stats, use_container_width=True, height=min(420, 60 + 35*len(df_stats)))
    st.subheader("🔁 Εφαρμοσθέντα Swaps")
    if swaps:
        df_swaps = pd.DataFrame(swaps)
        basic_cols = [c for c in ["type", "from_team", "students_out", "to_team", "students_in"] if c in df_swaps.columns]
        extra_cols = [c for c in df_swaps.columns if c not in basic_cols]
        st.dataframe(df_swaps[basic_cols + extra_cols], use_container_width=True, height=min(420, 60 + 28*len(df_swaps)))
    else:
        st.info("Δεν εφαρμόστηκαν swaps (ή ήδη ήσουν εντός στόχων).")
    st.subheader("⬇️ Λήψη Τελικού Excel")
    with open(final_path, "rb") as f:
        st.download_button("Κατέβασμα Τελικού (optimized)", f.read(),
                           file_name=Path(final_path).name,
                           type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_final_step8")

# >>> INLINE_STEP8_BLOCK_END <<<

# >>> INLINE_STEP8_BLOCK_CALL <<<
try:
    render_step8()
except Exception as _e_step8:
    import streamlit as st
    st.info("⚠️ Βήμα 8: Δεν κατέστη δυνατό να εκτελεστεί το inline block. Λόγος: " + str(_e_step8))

