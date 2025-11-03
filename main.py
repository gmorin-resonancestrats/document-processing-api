"""
API FastAPI pour Extraction et Reconstruction de Documents Word
Déployable sur Railway, Render, ou tout serveur Python

Endpoints:
- POST /extract-document: Extrait et analyse un document Word
- POST /reconstruct-document: Reconstruit un document Word à partir de segments traduits
- GET /health: Health check
"""

from fastapi import FastAPI, HTTPException, Header
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
import base64
import io
import zipfile
import json
import datetime
import re
import xml.etree.ElementTree as ET
from collections import Counter
from xml.sax.saxutils import escape as xml_escape
import logging
import os

# Configuration du logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuration
API_KEY = os.getenv("API_KEY", "votre-cle-secrete-a-changer")  # À définir dans les variables d'environnement
MAX_FILE_SIZE_MB = 100  # Limite de taille de fichier

app = FastAPI(
    title="Document Processing API",
    description="API pour extraction et reconstruction de documents Word",
    version="1.0.0"
)

# ==================== MODELS ====================

class ExtractDocumentRequest(BaseModel):
    docx_base64: str = Field(..., description="Document Word encodé en base64")
    file_name: str = Field(default="document.docx", description="Nom du fichier")
    client: Optional[str] = Field(default="", description="Nom du client")
    project: Optional[str] = Field(default="", description="Nom du projet")
    source_lang: Optional[str] = Field(default="", description="Langue source")
    target_lang: Optional[str] = Field(default="", description="Langue cible")

class ReconstructDocumentRequest(BaseModel):
    segments: List[Dict[str, Any]] = Field(..., description="Liste des segments traduits")
    document_metadata: Dict[str, Any] = Field(..., description="Métadonnées du document")
    file_name: str = Field(default="translated.docx", description="Nom du fichier de sortie")

# ==================== NAMESPACES XML ====================

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}

# ==================== UTILITY FUNCTIONS ====================

def verify_api_key(x_api_key: str = Header(None)):
    """Vérifie la clé API"""
    if x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid API Key")
    return x_api_key

def decode_docx(b64: str) -> io.BytesIO:
    """Décode un document Word depuis base64"""
    try:
        # Vérifier la taille
        estimated_size_mb = len(b64) * 0.75 / (1024 * 1024)  # Approximation
        if estimated_size_mb > MAX_FILE_SIZE_MB:
            raise ValueError(f"File too large: {estimated_size_mb:.1f}MB (max: {MAX_FILE_SIZE_MB}MB)")
        
        raw = base64.b64decode(b64)
        return io.BytesIO(raw)
    except Exception as e:
        logger.error(f"Error decoding docx: {e}")
        raise ValueError(f"Invalid base64 data: {e}")

def load_xml_from_zip(zf: zipfile.ZipFile, path: str):
    """Charge un fichier XML depuis le ZIP"""
    try:
        with zf.open(path) as f:
            return ET.parse(f).getroot()
    except KeyError:
        return None

def gen_id(prefix: str) -> str:
    """Génère un ID unique"""
    ts = datetime.datetime.utcnow().strftime("%Y%m%d%H%M%S%f")[:-3]
    return f"{prefix}_{ts}"

# ==================== EXTRACTION FUNCTIONS ====================

def map_style_ids_to_names(zf: zipfile.ZipFile) -> dict:
    """Mappe les IDs de style vers leurs noms"""
    w_ns = "{" + NS['w'] + "}"
    style_map = {}
    styles = load_xml_from_zip(zf, "word/styles.xml")
    if styles is None:
        return style_map
    
    for st in styles.findall(".//w:style", NS):
        style_id = st.attrib.get(w_ns + "styleId")
        name_el = st.find("w:name", NS)
        style_name = name_el.attrib.get(w_ns + "val") if name_el is not None else None
        if style_id:
            style_map[style_id] = style_name or style_id
    
    return style_map

def extract_doc_defaults(zf: zipfile.ZipFile) -> dict:
    """Extrait les propriétés par défaut du document"""
    w_ns = "{" + NS['w'] + "}"
    defaults = {
        "font_name": "",
        "font_size": "",
        "color": "",
        "bold": False,
        "italic": False,
        "underline": False
    }
    
    styles = load_xml_from_zip(zf, "word/styles.xml")
    if styles is None:
        return defaults
    
    doc_defaults = styles.find("w:docDefaults", NS)
    if doc_defaults is None:
        return defaults
    
    rPrDefault = doc_defaults.find(".//w:rPrDefault/w:rPr", NS)
    if rPrDefault is not None:
        # Police
        rFonts = rPrDefault.find("w:rFonts", NS)
        if rFonts is not None:
            for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                font = rFonts.attrib.get(w_ns + attr)
                if font:
                    defaults["font_name"] = font
                    break
        
        # Taille
        sz = rPrDefault.find("w:sz", NS)
        if sz is not None:
            val = sz.attrib.get(w_ns + "val")
            if val and val.isdigit():
                try:
                    defaults["font_size"] = round(int(val) / 2.0, 2)
                except:
                    pass
        
        # Couleur
        color = rPrDefault.find("w:color", NS)
        if color is not None:
            defaults["color"] = color.attrib.get(w_ns + "val", "")
        
        # Formatage
        defaults["bold"] = rPrDefault.find("w:b", NS) is not None
        defaults["italic"] = rPrDefault.find("w:i", NS) is not None
        defaults["underline"] = rPrDefault.find("w:u", NS) is not None
    
    return defaults

def extract_style_properties(zf: zipfile.ZipFile) -> dict:
    """Extrait les propriétés de formatage de chaque style avec héritage"""
    w_ns = "{" + NS['w'] + "}"
    style_props = {}
    style_inheritance = {}
    
    styles = load_xml_from_zip(zf, "word/styles.xml")
    if styles is None:
        return style_props
    
    # Première passe : extraire toutes les propriétés brutes
    for style in styles.findall(".//w:style", NS):
        style_id = style.attrib.get(w_ns + "styleId")
        if not style_id:
            continue
        
        # Vérifier l'héritage
        based_on = style.find("w:basedOn", NS)
        parent_id = based_on.attrib.get(w_ns + "val") if based_on is not None else None
        style_inheritance[style_id] = parent_id
        
        props = {
            "font_name": "",
            "font_size": "",
            "color": "",
            "bold": False,
            "italic": False,
            "underline": False,
            "underline_type": ""
        }
        
        # Propriétés de run dans le style
        rPr = style.find(".//w:rPr", NS)
        if rPr is not None:
            # Police
            rFonts = rPr.find("w:rFonts", NS)
            if rFonts is not None:
                for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                    font = rFonts.attrib.get(w_ns + attr)
                    if font:
                        props["font_name"] = font
                        break
            
            # Taille
            sz = rPr.find("w:sz", NS)
            if sz is not None:
                val = sz.attrib.get(w_ns + "val")
                if val and val.isdigit():
                    try:
                        props["font_size"] = round(int(val) / 2.0, 2)
                    except:
                        pass
            
            # Couleur
            color = rPr.find("w:color", NS)
            if color is not None:
                props["color"] = color.attrib.get(w_ns + "val", "")
            
            # Gras
            b = rPr.find("w:b", NS)
            if b is not None:
                b_val = b.attrib.get(w_ns + "val", "1")
                props["bold"] = b_val not in ("0", "false", "False")
            
            # Italique
            i = rPr.find("w:i", NS)
            if i is not None:
                i_val = i.attrib.get(w_ns + "val", "1")
                props["italic"] = i_val not in ("0", "false", "False")
            
            # Soulignement
            u = rPr.find("w:u", NS)
            if u is not None:
                u_val = u.attrib.get(w_ns + "val", "single")
                if u_val != "none":
                    props["underline"] = True
                    props["underline_type"] = u_val
        
        style_props[style_id] = props
    
    # Deuxième passe : appliquer l'héritage
    doc_defaults = extract_doc_defaults(zf)
    
    def resolve_style(style_id: str, visited: set = None, depth: int = 0):
        """Résout récursivement les propriétés d'un style"""
        MAX_DEPTH = 20
        if visited is None:
            visited = set()
        
        if style_id in visited or depth > MAX_DEPTH:
            return {}
        visited.add(style_id)
        
        if style_id not in style_props:
            return {}
        
        resolved = dict(doc_defaults)
        
        parent_id = style_inheritance.get(style_id)
        if parent_id:
            parent_props = resolve_style(parent_id, visited, depth + 1)
            for key, value in parent_props.items():
                if value:
                    resolved[key] = value
        
        current_props = style_props[style_id]
        for key, value in current_props.items():
            if value or isinstance(value, bool):
                resolved[key] = value
        
        return resolved
    
    resolved_styles = {}
    for style_id in style_props.keys():
        resolved_styles[style_id] = resolve_style(style_id)
    
    return resolved_styles

def get_color_value(color_elem):
    """Extrait la valeur de couleur (hex)"""
    if color_elem is None:
        return ""
    w_ns = "{" + NS['w'] + "}"
    return color_elem.attrib.get(w_ns + "val", "")

def get_spacing_values(spacing_elem):
    """Extrait les valeurs d'espacement"""
    if spacing_elem is None:
        return {"before": "", "after": "", "line": "", "lineRule": ""}
    w_ns = "{" + NS['w'] + "}"
    return {
        "before": spacing_elem.attrib.get(w_ns + "before", ""),
        "after": spacing_elem.attrib.get(w_ns + "after", ""),
        "line": spacing_elem.attrib.get(w_ns + "line", ""),
        "lineRule": spacing_elem.attrib.get(w_ns + "lineRule", "")
    }

def get_indent_values(ind_elem):
    """Extrait les valeurs d'indentation"""
    if ind_elem is None:
        return {"left": "", "right": "", "firstLine": "", "hanging": ""}
    w_ns = "{" + NS['w'] + "}"
    return {
        "left": ind_elem.attrib.get(w_ns + "left", ""),
        "right": ind_elem.attrib.get(w_ns + "right", ""),
        "firstLine": ind_elem.attrib.get(w_ns + "firstLine", ""),
        "hanging": ind_elem.attrib.get(w_ns + "hanging", "")
    }

def get_numbering_props(pPr):
    """Extrait les propriétés de numérotation (listes)"""
    if pPr is None:
        return {"numId": "", "ilvl": ""}
    numPr = pPr.find("w:numPr", NS)
    if numPr is None:
        return {"numId": "", "ilvl": ""}
    w_ns = "{" + NS['w'] + "}"
    numId_elem = numPr.find("w:numId", NS)
    ilvl_elem = numPr.find("w:ilvl", NS)
    return {
        "numId": numId_elem.attrib.get(w_ns + "val", "") if numId_elem is not None else "",
        "ilvl": ilvl_elem.attrib.get(w_ns + "val", "") if ilvl_elem is not None else ""
    }

def get_shading_values(shd_elem):
    """Extrait les valeurs de fond/surlignage"""
    if shd_elem is None:
        return {"fill": "", "color": ""}
    w_ns = "{" + NS['w'] + "}"
    return {
        "fill": shd_elem.attrib.get(w_ns + "fill", ""),
        "color": shd_elem.attrib.get(w_ns + "color", "")
    }

def get_paragraph_props(p, style_props: dict, doc_defaults: dict) -> dict:
    """Extrait TOUTES les propriétés de paragraphe et de run"""
    w_ns = "{" + NS['w'] + "}"
    pPr = p.find("w:pPr", NS)
    
    # Identifier le style du paragraphe
    style_id = None
    if pPr is not None:
        pStyle = pPr.find("w:pStyle", NS)
        if pStyle is not None:
            style_id = pStyle.attrib.get(w_ns + "val")
    
    # Propriétés par défaut : doc defaults + style
    style_defaults = dict(doc_defaults)
    if style_id and style_id in style_props:
        for key, value in style_props[style_id].items():
            if value or isinstance(value, bool):
                style_defaults[key] = value
    
    # Propriétés de paragraphe
    align = None
    spacing = {"before": "", "after": "", "line": "", "lineRule": ""}
    indent = {"left": "", "right": "", "firstLine": "", "hanging": ""}
    numbering = {"numId": "", "ilvl": ""}
    keep_next = False
    keep_lines = False
    page_break_before = False
    shading_para = {"fill": "", "color": ""}
    
    if pPr is not None:
        jc = pPr.find("w:jc", NS)
        if jc is not None:
            align = jc.attrib.get(w_ns + "val")
        
        spacing = get_spacing_values(pPr.find("w:spacing", NS))
        indent = get_indent_values(pPr.find("w:ind", NS))
        numbering = get_numbering_props(pPr)
        
        keep_next = pPr.find("w:keepNext", NS) is not None
        keep_lines = pPr.find("w:keepLines", NS) is not None
        page_break_before = pPr.find("w:pageBreakBefore", NS) is not None
        
        shading_para = get_shading_values(pPr.find("w:shd", NS))

    # Propriétés de run (agrégées)
    bold_any = False
    italic_any = False
    underline_any = False
    underline_type = ""
    strike_any = False
    double_strike_any = False
    small_caps_any = False
    all_caps_any = False
    
    font_names = []
    font_sizes = []
    colors = []
    highlights = []
    shading_runs = []
    
    has_explicit_run_props = False
    
    for r in p.findall("w:r", NS):
        rPr = r.find("w:rPr", NS)
        if rPr is not None:
            has_explicit_run_props = True
            
            # Gras
            b = rPr.find("w:b", NS)
            if b is not None:
                b_val = b.attrib.get(w_ns + "val", "1")
                if b_val not in ("0", "false", "False"):
                    bold_any = True
            
            # Italique
            i = rPr.find("w:i", NS)
            if i is not None:
                i_val = i.attrib.get(w_ns + "val", "1")
                if i_val not in ("0", "false", "False"):
                    italic_any = True
            
            # Soulignement
            u = rPr.find("w:u", NS)
            if u is not None:
                u_val = u.attrib.get(w_ns + "val", "single")
                if u_val != "none":
                    underline_any = True
                    underline_type = u_val
            
            if rPr.find("w:strike", NS) is not None:
                strike_any = True
            if rPr.find("w:dstrike", NS) is not None:
                double_strike_any = True
            if rPr.find("w:smallCaps", NS) is not None:
                small_caps_any = True
            if rPr.find("w:caps", NS) is not None:
                all_caps_any = True
            
            # Police
            rFonts = rPr.find("w:rFonts", NS)
            if rFonts is not None:
                for k in ("ascii", "hAnsi", "cs", "eastAsia"):
                    v = rFonts.attrib.get(w_ns + k)
                    if v:
                        font_names.append(v)
                        break
            
            # Taille
            sz = rPr.find("w:sz", NS)
            if sz is not None:
                val = sz.attrib.get(w_ns + "val")
                if val and val.isdigit():
                    try:
                        font_sizes.append(round(int(val) / 2.0, 2))
                    except:
                        pass
            
            # Couleur
            color = rPr.find("w:color", NS)
            if color is not None:
                colors.append(get_color_value(color))
            
            # Surlignage
            highlight = rPr.find("w:highlight", NS)
            if highlight is not None:
                highlights.append(highlight.attrib.get(w_ns + "val", ""))
            
            # Fond
            shd = rPr.find("w:shd", NS)
            if shd is not None:
                shading_runs.append(get_shading_values(shd))

    # Agrégation des valeurs
    font_name = Counter(font_names).most_common(1)[0][0] if font_names else ""
    font_size = Counter(font_sizes).most_common(1)[0][0] if font_sizes else ""
    color = Counter(colors).most_common(1)[0][0] if colors else ""
    highlight = Counter(highlights).most_common(1)[0][0] if highlights else ""
    
    shading_fill = ""
    shading_color = ""
    if shading_runs:
        fills = [s["fill"] for s in shading_runs if s["fill"]]
        colors_shd = [s["color"] for s in shading_runs if s["color"]]
        shading_fill = Counter(fills).most_common(1)[0][0] if fills else ""
        shading_color = Counter(colors_shd).most_common(1)[0][0] if colors_shd else ""

    # Utiliser les valeurs du style si aucune propriété explicite
    if not font_name:
        font_name = style_defaults.get("font_name", "")
    
    if not font_size:
        font_size = style_defaults.get("font_size", "")
    
    if not color:
        color = style_defaults.get("color", "")
    
    # Pour bold/italic/underline, utiliser le style SEULEMENT si aucun run n'a de propriétés explicites
    if not has_explicit_run_props:
        if not bold_any and style_defaults.get("bold"):
            bold_any = True
        
        if not italic_any and style_defaults.get("italic"):
            italic_any = True
        
        if not underline_any and style_defaults.get("underline"):
            underline_any = True
            if not underline_type:
                underline_type = style_defaults.get("underline_type", "single")

    return {
        "style_id": style_id,
        "align": align or "left",
        "spacing_before": spacing["before"],
        "spacing_after": spacing["after"],
        "spacing_line": spacing["line"],
        "spacing_line_rule": spacing["lineRule"],
        "indent_left": indent["left"],
        "indent_right": indent["right"],
        "indent_first_line": indent["firstLine"],
        "indent_hanging": indent["hanging"],
        "numId": numbering["numId"],
        "ilvl": numbering["ilvl"],
        "keep_next": keep_next,
        "keep_lines": keep_lines,
        "page_break_before": page_break_before,
        "shading_para_fill": shading_para["fill"],
        "shading_para_color": shading_para["color"],
        "bold_any": bold_any,
        "italic_any": italic_any,
        "underline_any": underline_any,
        "underline_type": underline_type,
        "strike_any": strike_any,
        "double_strike_any": double_strike_any,
        "small_caps_any": small_caps_any,
        "all_caps_any": all_caps_any,
        "font_name_major": font_name,
        "font_size_pt_major": font_size,
        "color": color,
        "highlight": highlight,
        "shading_fill": shading_fill,
        "shading_color": shading_color
    }


def extract_runs_with_formatting(p):
    """Extrait les runs avec leur formatage individuel ET les caractères spéciaux"""
    w_ns = "{" + NS['w'] + "}"
    runs_data = []
    
    for r in p.findall("w:r", NS):
        run_info = {"texts": [], "props": {}}
        
        rPr = r.find("w:rPr", NS)
        if rPr is not None:
            b = rPr.find("w:b", NS)
            if b is not None:
                b_val = b.attrib.get(w_ns + "val", "1")
                run_info["props"]["bold"] = b_val not in ("0", "false", "False")
            
            i_elem = rPr.find("w:i", NS)
            if i_elem is not None:
                i_val = i_elem.attrib.get(w_ns + "val", "1")
                run_info["props"]["italic"] = i_val not in ("0", "false", "False")
            
            u = rPr.find("w:u", NS)
            if u is not None:
                u_val = u.attrib.get(w_ns + "val", "single")
                if u_val != "none":
                    run_info["props"]["underline"] = True
                    run_info["props"]["underline_type"] = u_val
            
            if rPr.find("w:strike", NS) is not None:
                run_info["props"]["strike"] = True
            
            rFonts = rPr.find("w:rFonts", NS)
            if rFonts is not None:
                for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                    font = rFonts.attrib.get(w_ns + attr)
                    if font:
                        run_info["props"]["font_name"] = font
                        break
            
            sz = rPr.find("w:sz", NS)
            if sz is not None:
                val = sz.attrib.get(w_ns + "val")
                if val and val.isdigit():
                    try:
                        run_info["props"]["font_size"] = round(int(val) / 2.0, 2)
                    except:
                        pass
            
            color = rPr.find("w:color", NS)
            if color is not None:
                run_info["props"]["color"] = color.attrib.get(w_ns + "val", "")
            
            highlight = rPr.find("w:highlight", NS)
            if highlight is not None:
                run_info["props"]["highlight"] = highlight.attrib.get(w_ns + "val", "")
        
        for child in r:
            if child.tag == w_ns + "t":
                if child.text:
                    run_info["texts"].append({"type": "text", "value": child.text})
            elif child.tag == w_ns + "tab":
                run_info["texts"].append({"type": "tab", "value": "\t"})
            elif child.tag == w_ns + "br":
                br_type = child.attrib.get(w_ns + "type", "textWrapping")
                if br_type == "page":
                    run_info["texts"].append({"type": "page_break", "value": "<<PAGE_BREAK>>"})
                else:
                    run_info["texts"].append({"type": "line_break", "value": "\n"})
            elif child.tag == w_ns + "noBreakHyphen":
                run_info["texts"].append({"type": "text", "value": "‑"})
        
        if run_info["texts"]:
            runs_data.append(run_info)
    
    return runs_data

def get_paragraph_text_by_sentences(p):
    """Extrait le texte avec caractères spéciaux ET la structure des runs"""
    w_ns = "{" + NS['w'] + "}"
    
    runs_data = extract_runs_with_formatting(p)
    
    full_text_parts = []
    for run in runs_data:
        for text_segment in run["texts"]:
            full_text_parts.append(text_segment["value"])
    
    full_text = "".join(full_text_parts).strip()
    
    if not full_text:
        return [""], []
    
    sentence_endings = re.compile(r'([.!?]+)(\s+|$)')
    sentences = []
    current_pos = 0
    
    for match in sentence_endings.finditer(full_text):
        end_pos = match.end()
        sentence = full_text[current_pos:end_pos].strip()
        if sentence:
            sentences.append(sentence)
        current_pos = end_pos
    
    if current_pos < len(full_text):
        remaining = full_text[current_pos:].strip()
        if remaining:
            sentences.append(remaining)
    
    return (sentences if sentences else [full_text]), runs_data


def get_table_cell_props(tc):
    """Extrait les propriétés de cellule de tableau"""
    w_ns = "{" + NS['w'] + "}"
    tcPr = tc.find("w:tcPr", NS)
    
    props = {
        "gridSpan": "",
        "vMerge": "",
        "vAlign": "",
        "tcW": "",
        "shading_fill": "",
        "shading_color": ""
    }
    
    if tcPr is not None:
        gridSpan = tcPr.find("w:gridSpan", NS)
        if gridSpan is not None:
            props["gridSpan"] = gridSpan.attrib.get(w_ns + "val", "")
        
        vMerge = tcPr.find("w:vMerge", NS)
        if vMerge is not None:
            props["vMerge"] = vMerge.attrib.get(w_ns + "val", "restart")
        
        vAlign = tcPr.find("w:vAlign", NS)
        if vAlign is not None:
            props["vAlign"] = vAlign.attrib.get(w_ns + "val", "")
        
        tcW = tcPr.find("w:tcW", NS)
        if tcW is not None:
            props["tcW"] = tcW.attrib.get(w_ns + "w", "")
        
        shd = tcPr.find("w:shd", NS)
        if shd is not None:
            shading = get_shading_values(shd)
            props["shading_fill"] = shading["fill"]
            props["shading_color"] = shading["color"]
    
    return props

def extract_paragraph_segment(p, style_map: dict, style_props: dict, doc_defaults: dict, 
                              file_name: str, paragraph_id: str, 
                              table_index=None, row_index=None, col_index=None, cell_props=None):
    """Extrait UN segment par phrase du paragraphe avec paragraph_id"""
    props = get_paragraph_props(p, style_props, doc_defaults)
    style_name = style_map.get(props["style_id"], props["style_id"] or "Normal")
    
    # Extraire le texte par phrases
    sentences, runs_data = get_paragraph_text_by_sentences(p)
    runs_json = json.dumps(runs_data, ensure_ascii=False)
    
    # Créer un segment par phrase
    segments = []
    for sentence_text in sentences:
        seg = {
            "paragraph_id": paragraph_id,
            "style": style_name,
            "content_source": sentence_text,
            "runs_data": runs_json,
            "alignment": props["align"],
            "spacing_before": props["spacing_before"],
            "spacing_after": props["spacing_after"],
            "spacing_line": props["spacing_line"],
            "spacing_line_rule": props["spacing_line_rule"],
            "indent_left": props["indent_left"],
            "indent_right": props["indent_right"],
            "indent_first_line": props["indent_first_line"],
            "indent_hanging": props["indent_hanging"],
            "numId": props["numId"],
            "ilvl": props["ilvl"],
            "keep_next": props["keep_next"],
            "keep_lines": props["keep_lines"],
            "page_break_before": props["page_break_before"],
            "shading_para_fill": props["shading_para_fill"],
            "shading_para_color": props["shading_para_color"],
            "bold_any": props["bold_any"],
            "italic_any": props["italic_any"],
            "underline_any": props["underline_any"],
            "underline_type": props["underline_type"],
            "strike_any": props["strike_any"],
            "double_strike_any": props["double_strike_any"],
            "small_caps_any": props["small_caps_any"],
            "all_caps_any": props["all_caps_any"],
            "font_name_major": props["font_name_major"],
            "font_size_pt_major": props["font_size_pt_major"],
            "color": props["color"],
            "highlight": props["highlight"],
            "shading_fill": props["shading_fill"],
            "shading_color": props["shading_color"],
            "file_name": file_name,
            "table_index": table_index if table_index is not None else "",
            "row_index": row_index if row_index is not None else "",
            "col_index": col_index if col_index is not None else ""
        }
        
        # Ajouter les propriétés de cellule si applicable
        if cell_props:
            seg.update({
                "cell_gridSpan": cell_props["gridSpan"],
                "cell_vMerge": cell_props["vMerge"],
                "cell_vAlign": cell_props["vAlign"],
                "cell_width": cell_props["tcW"],
                "cell_shading_fill": cell_props["shading_fill"],
                "cell_shading_color": cell_props["shading_color"]
            })
        else:
            seg.update({
                "cell_gridSpan": "",
                "cell_vMerge": "",
                "cell_vAlign": "",
                "cell_width": "",
                "cell_shading_fill": "",
                "cell_shading_color": ""
            })
        
        segments.append(seg)
    
    return segments

def extract_from_parent(parent, style_map: dict, style_props: dict, doc_defaults: dict, 
                       file_name: str, paragraph_counter: list, 
                       table_index=None, row_index=None, col_index=None, cell_props=None):
    """Extrait les segments d'un parent avec compteur de paragraphes"""
    segments = []
    for p in parent.findall("w:p", NS):
        paragraph_counter[0] += 1
        paragraph_id = f"P{paragraph_counter[0]:06d}"
        segs = extract_paragraph_segment(p, style_map, style_props, doc_defaults, 
                                        file_name, paragraph_id, table_index, 
                                        row_index, col_index, cell_props)
        segments.extend(segs)
    return segments

def extract_document(zf: zipfile.ZipFile, file_name: str) -> list:
    """Extrait le document complet avec toutes les propriétés"""
    w_ns = "{" + NS['w'] + "}"
    
    logger.info(f"Extracting document: {file_name}")
    
    style_map = map_style_ids_to_names(zf)
    doc_defaults = extract_doc_defaults(zf)
    style_props = extract_style_properties(zf)
    
    doc = load_xml_from_zip(zf, "word/document.xml")
    if doc is None:
        raise ValueError("Invalid Word document: document.xml not found")
    
    body = doc.find("w:body", NS)
    if body is None:
        raise ValueError("Invalid Word document: body not found")

    segments = []
    t_idx = 0
    paragraph_counter = [0]

    # Parcourir le body
    for child in list(body):
        if child.tag == w_ns + "p":
            paragraph_counter[0] += 1
            paragraph_id = f"P{paragraph_counter[0]:06d}"
            segs = extract_paragraph_segment(child, style_map, style_props, 
                                            doc_defaults, file_name, paragraph_id)
            segments.extend(segs)
            
        elif child.tag == w_ns + "tbl":
            t_idx += 1
            r_idx = 0
            for tr in child.findall("w:tr", NS):
                r_idx += 1
                c_idx = 0
                for tc in tr.findall("w:tc", NS):
                    c_idx += 1
                    cell_props = get_table_cell_props(tc)
                    segs = extract_from_parent(tc, style_map, style_props, 
                                              doc_defaults, file_name, 
                                              paragraph_counter, t_idx, r_idx, c_idx, cell_props)
                    for s in segs:
                        s["style"] = "Table Cell / " + s["style"]
                    segments.extend(segs)

    # En-têtes et pieds de page
    for name in zf.namelist():
        if name.startswith("word/header") and name.endswith(".xml"):
            root = load_xml_from_zip(zf, name)
            if root is not None:
                for p in root.findall(".//w:p", NS):
                    paragraph_counter[0] += 1
                    paragraph_id = f"P{paragraph_counter[0]:06d}"
                    segs = extract_paragraph_segment(p, style_map, style_props, 
                                                    doc_defaults, file_name, paragraph_id)
                    for s in segs:
                        s["style"] = "Header / " + s["style"]
                    segments.extend(segs)
        
        if name.startswith("word/footer") and name.endswith(".xml"):
            root = load_xml_from_zip(zf, name)
            if root is not None:
                for p in root.findall(".//w:p", NS):
                    paragraph_counter[0] += 1
                    paragraph_id = f"P{paragraph_counter[0]:06d}"
                    segs = extract_paragraph_segment(p, style_map, style_props, 
                                                    doc_defaults, file_name, paragraph_id)
                    for s in segs:
                        s["style"] = "Footer / " + s["style"]
                    segments.extend(segs)
    
    logger.info(f"Extracted {len(segments)} segments from {paragraph_counter[0]} paragraphs")
    return segments

# ==================== API ENDPOINTS ====================

@app.get("/")
async def root():
    """Root endpoint"""
    return {
        "service": "Document Processing API",
        "version": "1.0.0",
        "endpoints": {
            "health": "/health",
            "extract": "POST /extract-document",
            "reconstruct": "POST /reconstruct-document"
        }
    }

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "timestamp": datetime.datetime.utcnow().isoformat()
    }

@app.post("/extract-document")
async def extract_document_endpoint(
    request: ExtractDocumentRequest,
    x_api_key: str = Header(None)
):
    """
    Extrait et analyse un document Word
    
    Retourne un JSON avec:
    - document: métadonnées du document
    - segments: liste de tous les segments avec leurs propriétés complètes
    """
    try:
        # Vérifier l'API key
        verify_api_key(x_api_key)
        
        logger.info(f"Starting extraction for: {request.file_name}")
        start_time = datetime.datetime.utcnow()
        
        # Décoder et valider le document
        bio = decode_docx(request.docx_base64)
        
        # Extraire le document
        with zipfile.ZipFile(bio) as zf:
            segs = extract_document(zf, request.file_name)
        
        # Générer l'output
        document_id = gen_id("DOC")
        out_segments = []
        
        for i, s in enumerate(segs, start=1):
            seg_id = f"SEG_{str(i).zfill(6)}"
            out_segments.append({
                "segment_id": seg_id,
                "document_id": document_id,
                "order": i,
                **s
            })
        
        end_time = datetime.datetime.utcnow()
        processing_time = (end_time - start_time).total_seconds()
        
        logger.info(f"Extraction completed in {processing_time:.2f}s - {len(out_segments)} segments")
        
        return {
            "success": True,
            "document": {
                "document_id": document_id,
                "file_name": request.file_name,
                "client": request.client,
                "project": request.project,
                "source_lang": request.source_lang,
                "target_lang": request.target_lang,
                "status": "extracted",
                "total_segments": len(out_segments),
                "processing_time_seconds": round(processing_time, 2)
            },
            "segments": out_segments
        }
        
    except ValueError as e:
        logger.error(f"Validation error: {e}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error(f"Extraction error: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Extraction failed: {str(e)}")

@app.post("/reconstruct-document")
async def reconstruct_document_endpoint(
    request: ReconstructDocumentRequest,
    x_api_key: str = Header(None)
):
    """
    Reconstruit un document Word à partir de segments traduits
    
    Cette fonction sera implémentée avec le code du module 32
    Pour l'instant, retourne un placeholder
    """
    try:
        # Vérifier l'API key
        verify_api_key(x_api_key)
        
        logger.info(f"Starting reconstruction for: {request.file_name}")
        
        # TODO: Implémenter la reconstruction du document
        # Ce sera le code du module 32 adapté
        
        return {
            "success": True,
            "message": "Reconstruction endpoint - to be implemented",
            "file_name": request.file_name,
            "segments_count": len(request.segments)
        }
        
    except Exception as e:
        logger.error(f"Reconstruction error: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Reconstruction failed: {str(e)}")

# ==================== ERROR HANDLERS ====================

@app.exception_handler(HTTPException)
async def http_exception_handler(request, exc):
    return JSONResponse(
        status_code=exc.status_code,
        content={
            "success": False,
            "error": exc.detail,
            "timestamp": datetime.datetime.utcnow().isoformat()
        }
    )

@app.exception_handler(Exception)
async def general_exception_handler(request, exc):
    logger.error(f"Unhandled exception: {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={
            "success": False,
            "error": "Internal server error",
            "detail": str(exc),
            "timestamp": datetime.datetime.utcnow().isoformat()
        }
    )

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)