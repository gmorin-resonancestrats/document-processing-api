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

def get_paragraph_text_by_sentences(p):
    """Extrait le texte du paragraphe divisé par phrases"""
    w_ns = "{" + NS['w'] + "}"
    
    # Construire le texte complet du paragraphe
    text_parts = []
    for node in p.iter():
        if node.tag == w_ns + "t":
            text_parts.append(node.text or "")
    
    full_text = "".join(text_parts).strip()
    
    if not full_text:
        return [""]
    
    # Diviser par phrases
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
    
    return sentences if sentences else [full_text]

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
    sentences = get_paragraph_text_by_sentences(p)
    
    # Créer un segment par phrase
    segments = []
    for sentence_text in sentences:
        seg = {
            "paragraph_id": paragraph_id,
            "style": style_name,
            "content_source": sentence_text,
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

# ==================== RECONSTRUCTION FUNCTIONS ====================

def as_bool(v):
    """Convertit une valeur en booléen"""
    if v in (True, 1, "1", "true", "True", "TRUE", "yes"):
        return True
    return False

def to_int(v):
    """Convertit une valeur en entier"""
    try:
        if v in (None, "", "null"):
            return None
        return int(v)
    except:
        return None

def to_str(v):
    """Convertit une valeur en string, retourne vide si None"""
    if v in (None, "null", "None"):
        return ""
    return str(v)

def norm_align(a):
    """Normalise l'alignement"""
    a = (a or "").lower()
    return {"left": "left", "center": "center", "right": "right", "both": "both", "justify": "both"}.get(a, "left")

def norm_style(s):
    """Normalise le style"""
    s = (s or "").lower()
    if "table cell" in s:
        return s
    if "heading 1" in s:
        return "heading1"
    if "heading 2" in s:
        return "heading2"
    if "heading 3" in s:
        return "heading3"
    if "heading 4" in s:
        return "heading4"
    if "list paragraph" in s:
        return "listparagraph"
    return "normal"

def xml_header():
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

def content_types_xml():
    return f'''{xml_header()}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

def rels_root_xml():
    return f'''{xml_header()}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

def document_rels_xml():
    return f'''{xml_header()}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

def app_xml():
    return f'''{xml_header()}
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Make DOCX Builder Enhanced v2.0</Application>
</Properties>'''

def core_xml(creator="Make", title="Rebuilt", subject="", description=""):
    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return f'''{xml_header()}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{xml_escape(title)}</dc:title>
  <dc:subject>{xml_escape(subject)}</dc:subject>
  <dc:creator>{xml_escape(creator)}</dc:creator>
  <cp:description>{xml_escape(description)}</cp:description>
  <cp:lastModifiedBy>{xml_escape(creator)}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

def styles_xml():
    return f'''{xml_header()}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="200" w:after="40"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="160" w:after="20"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="120" w:after="20"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/>
    <w:pPr><w:ind w:left="720"/></w:pPr>
  </w:style>
</w:styles>'''

def numbering_xml():
    """Génère le fichier de numérotation pour les listes"""
    return f'''{xml_header()}
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>'''

def to_w_style(style_key):
    """Convertit le style normalisé en ID Word"""
    m = {
        "heading1": "Heading1", "heading2": "Heading2", "heading3": "Heading3", "heading4": "Heading4",
        "listparagraph": "ListParagraph", "normal": "Normal"
    }
    return m.get(style_key, "Normal")

def align_to_wjc(al):
    """Convertit l'alignement en valeur Word"""
    return {"left": "left", "center": "center", "right": "right", "both": "both"}.get(al, "left")

# ==================== NOUVELLES FONCTIONS POUR FORMATAGE MIXTE ====================

def build_run_xml_from_data(run_data, default_props):
    """
    ✨ NOUVELLE FONCTION ✨
    Construit un run XML à partir de la structure détaillée.
    Supporte le formatage mixte et les caractères spéciaux.
    """
    parts = []
    
    # Propriétés du run (fusionner avec les valeurs par défaut)
    props = {**default_props, **run_data.get("props", {})}
    
    # Construire rPr
    rPr_parts = []
    
    if props.get("bold"):
        rPr_parts.append('<w:b/>')
    
    if props.get("italic"):
        rPr_parts.append('<w:i/>')
    
    if props.get("underline"):
        u_type = props.get("underline_type", "single")
        rPr_parts.append(f'<w:u w:val="{u_type}"/>')
    
    if props.get("strike"):
        rPr_parts.append('<w:strike/>')
    
    font_name = props.get("font_name", "")
    if font_name:
        rPr_parts.append(f'<w:rFonts w:ascii="{xml_escape(font_name)}" w:hAnsi="{xml_escape(font_name)}"/>')
    
    font_size = props.get("font_size")
    if font_size:
        try:
            sz = int(float(font_size) * 2)
            rPr_parts.append(f'<w:sz w:val="{sz}"/>')
            rPr_parts.append(f'<w:szCs w:val="{sz}"/>')
        except:
            pass
    
    color = props.get("color", "")
    if color and color != "auto":
        rPr_parts.append(f'<w:color w:val="{color}"/>')
    
    highlight = props.get("highlight", "")
    if highlight:
        rPr_parts.append(f'<w:highlight w:val="{highlight}"/>')
    
    rPr_xml = "".join(rPr_parts)
    
    # Construire le contenu du run avec caractères spéciaux
    for text_segment in run_data.get("texts", []):
        seg_type = text_segment.get("type", "text")
        value = text_segment.get("value", "")
        
        if seg_type == "text":
            # Texte normal
            escaped_text = xml_escape(value)
            if rPr_xml:
                parts.append(f'<w:r><w:rPr>{rPr_xml}</w:rPr><w:t xml:space="preserve">{escaped_text}</w:t></w:r>')
            else:
                parts.append(f'<w:r><w:t xml:space="preserve">{escaped_text}</w:t></w:r>')
        
        elif seg_type == "tab":
            # Tabulation
            if rPr_xml:
                parts.append(f'<w:r><w:rPr>{rPr_xml}</w:rPr><w:tab/></w:r>')
            else:
                parts.append('<w:r><w:tab/></w:r>')
        
        elif seg_type == "line_break":
            # Saut de ligne manuel
            if rPr_xml:
                parts.append(f'<w:r><w:rPr>{rPr_xml}</w:rPr><w:br/></w:r>')
            else:
                parts.append('<w:r><w:br/></w:r>')
        
        elif seg_type == "page_break":
            # Saut de page manuel
            if rPr_xml:
                parts.append(f'<w:r><w:rPr>{rPr_xml}</w:rPr><w:br w:type="page"/></w:r>')
            else:
                parts.append('<w:r><w:br w:type="page"/></w:r>')
    
    return "".join(parts)

def build_run_props(seg):
    """Construit les propriétés de run (rPr) à partir du segment - VERSION SIMPLE"""
    parts = []
    
    if as_bool(seg.get("bold_any")):
        parts.append('<w:b/>')
    
    if as_bool(seg.get("italic_any")):
        parts.append('<w:i/>')
    
    if as_bool(seg.get("underline_any")):
        u_type = to_str(seg.get("underline_type")) or "single"
        parts.append(f'<w:u w:val="{u_type}"/>')
    
    if as_bool(seg.get("strike_any")):
        parts.append('<w:strike/>')
    
    if as_bool(seg.get("double_strike_any")):
        parts.append('<w:dstrike/>')
    
    if as_bool(seg.get("small_caps_any")):
        parts.append('<w:smallCaps/>')
    
    if as_bool(seg.get("all_caps_any")):
        parts.append('<w:caps/>')
    
    font_name = to_str(seg.get("font_name_major"))
    if font_name:
        parts.append(f'<w:rFonts w:ascii="{xml_escape(font_name)}" w:hAnsi="{xml_escape(font_name)}"/>')
    
    font_size = seg.get("font_size_pt_major")
    if font_size:
        try:
            sz = int(float(font_size) * 2)
            parts.append(f'<w:sz w:val="{sz}"/>')
            parts.append(f'<w:szCs w:val="{sz}"/>')
        except:
            pass
    
    color = to_str(seg.get("color"))
    if color and color != "auto":
        parts.append(f'<w:color w:val="{color}"/>')
    
    highlight = to_str(seg.get("highlight"))
    if highlight:
        parts.append(f'<w:highlight w:val="{highlight}"/>')
    
    shading_fill = to_str(seg.get("shading_fill"))
    if shading_fill:
        parts.append(f'<w:shd w:val="clear" w:fill="{shading_fill}"/>')
    
    return "".join(parts)

def build_paragraph_props(seg):
    """Construit les propriétés de paragraphe (pPr) à partir du segment"""
    parts = []
    
    style_key = norm_style(seg.get("style", ""))
    if not style_key.startswith("table cell"):
        style_id = to_w_style(style_key)
        parts.append(f'<w:pStyle w:val="{style_id}"/>')
    
    align = norm_align(seg.get("alignment"))
    jc = align_to_wjc(align)
    parts.append(f'<w:jc w:val="{jc}"/>')
    
    spacing_parts = []
    spacing_before = to_str(seg.get("spacing_before"))
    spacing_after = to_str(seg.get("spacing_after"))
    spacing_line = to_str(seg.get("spacing_line"))
    spacing_line_rule = to_str(seg.get("spacing_line_rule"))
    
    if spacing_before:
        spacing_parts.append(f'w:before="{spacing_before}"')
    if spacing_after:
        spacing_parts.append(f'w:after="{spacing_after}"')
    if spacing_line:
        spacing_parts.append(f'w:line="{spacing_line}"')
    if spacing_line_rule:
        spacing_parts.append(f'w:lineRule="{spacing_line_rule}"')
    
    if spacing_parts:
        parts.append(f'<w:spacing {" ".join(spacing_parts)}/>')
    
    indent_parts = []
    indent_left = to_str(seg.get("indent_left"))
    indent_right = to_str(seg.get("indent_right"))
    indent_first = to_str(seg.get("indent_first_line"))
    indent_hanging = to_str(seg.get("indent_hanging"))
    
    if indent_left:
        indent_parts.append(f'w:left="{indent_left}"')
    if indent_right:
        indent_parts.append(f'w:right="{indent_right}"')
    if indent_first:
        indent_parts.append(f'w:firstLine="{indent_first}"')
    if indent_hanging:
        indent_parts.append(f'w:hanging="{indent_hanging}"')
    
    if indent_parts:
        parts.append(f'<w:ind {" ".join(indent_parts)}/>')
    
    num_id = to_str(seg.get("numId"))
    ilvl = to_str(seg.get("ilvl"))
    if num_id:
        parts.append(f'<w:numPr><w:ilvl w:val="{ilvl or "0"}"/><w:numId w:val="{num_id}"/></w:numPr>')
    
    if as_bool(seg.get("keep_next")):
        parts.append('<w:keepNext/>')
    if as_bool(seg.get("keep_lines")):
        parts.append('<w:keepLines/>')
    if as_bool(seg.get("page_break_before")):
        parts.append('<w:pageBreakBefore/>')
    
    shading_para_fill = to_str(seg.get("shading_para_fill"))
    if shading_para_fill:
        parts.append(f'<w:shd w:val="clear" w:fill="{shading_para_fill}"/>')
    
    return "".join(parts)

def run_xml(text, seg):
    """Génère un run XML simple - VERSION FALLBACK"""
    t = xml_escape(text or "")
    rPr = build_run_props(seg)
    
    if rPr:
        return f'<w:r><w:rPr>{rPr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
    else:
        return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'

def paragraph_xml_enhanced(seg):
    """
    ✨ VERSION AMÉLIORÉE ✨
    Génère un paragraphe XML avec formatage mixte et caractères spéciaux.
    """
    pPr = build_paragraph_props(seg)
    
    # Vérifier si on a des données de runs détaillées
    runs_json = seg.get("runs_data", "")
    
    if runs_json and runs_json != "":
        try:
            # Utiliser la structure détaillée des runs
            runs_data = json.loads(runs_json)
            
            # Extraire les propriétés par défaut du segment
            default_props = {
                "bold": as_bool(seg.get("bold_any")),
                "italic": as_bool(seg.get("italic_any")),
                "underline": as_bool(seg.get("underline_any")),
                "underline_type": to_str(seg.get("underline_type")) or "single",
                "strike": as_bool(seg.get("strike_any")),
                "font_name": to_str(seg.get("font_name_major")),
                "font_size": seg.get("font_size_pt_major"),
                "color": to_str(seg.get("color")),
                "highlight": to_str(seg.get("highlight"))
            }
            
            # Construire tous les runs
            runs_xml_parts = []
            for run_data in runs_data:
                run_xml_str = build_run_xml_from_data(run_data, default_props)
                if run_xml_str:
                    runs_xml_parts.append(run_xml_str)
            
            runs_xml = "".join(runs_xml_parts)
            
            return f'<w:p><w:pPr>{pPr}</w:pPr>{runs_xml}</w:p>'
        
        except (json.JSONDecodeError, Exception):
            # Fallback sur l'ancienne méthode si parsing échoue
            pass
    
    # Méthode de fallback: un seul run avec tout le texte
    text = seg.get("content_source", "")
    r = run_xml(text, seg)
    return f'<w:p><w:pPr>{pPr}</w:pPr>{r}</w:p>'

def build_table_cell_props(seg):
    """Construit les propriétés de cellule (tcPr)"""
    parts = []
    
    grid_span = to_str(seg.get("cell_gridSpan"))
    if grid_span and grid_span != "1":
        parts.append(f'<w:gridSpan w:val="{grid_span}"/>')
    
    v_merge = to_str(seg.get("cell_vMerge"))
    if v_merge:
        if v_merge == "restart":
            parts.append('<w:vMerge w:val="restart"/>')
        else:
            parts.append('<w:vMerge/>')
    
    v_align = to_str(seg.get("cell_vAlign"))
    if v_align:
        parts.append(f'<w:vAlign w:val="{v_align}"/>')
    
    cell_width = to_str(seg.get("cell_width"))
    if cell_width:
        parts.append(f'<w:tcW w:w="{cell_width}" w:type="dxa"/>')
    else:
        parts.append('<w:tcW w:w="2390" w:type="dxa"/>')
    
    cell_shading = to_str(seg.get("cell_shading_fill"))
    if cell_shading:
        parts.append(f'<w:shd w:val="clear" w:fill="{cell_shading}"/>')
    
    return "".join(parts)

def table_xml_enhanced(cells, rows, cols):
    """
    ✨ VERSION AMÉLIORÉE ✨
    Génère un tableau avec support du formatage mixte.
    """
    parts = []
    parts.append('<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">')
    parts.append('<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>')
    parts.append('<w:tblGrid>' + ''.join('<w:gridCol w:w="2390"/>' for _ in range(cols)) + '</w:tblGrid>')
    
    for r in range(1, rows + 1):
        parts.append('<w:tr>')
        for c in range(1, cols + 1):
            parts.append('<w:tc>')
            
            cell_segments = cells.get((r, c), [])
            if cell_segments:
                first_seg = cell_segments[0]
                tcPr = build_table_cell_props(first_seg)
                parts.append(f'<w:tcPr>{tcPr}</w:tcPr>')
            else:
                parts.append('<w:tcPr><w:tcW w:w="2390" w:type="dxa"/></w:tcPr>')
            
            if cell_segments:
                # Regrouper par paragraph_id
                paragraphs = {}
                for seg in cell_segments:
                    para_id = seg.get("paragraph_id", seg.get("order", ""))
                    if para_id not in paragraphs:
                        paragraphs[para_id] = []
                    paragraphs[para_id].append(seg)
                
                # Générer un paragraphe par paragraph_id
                for para_id in sorted(paragraphs.keys()):
                    para_segments = sorted(paragraphs[para_id], key=lambda x: x.get("order", 0))
                    
                    # Combiner intelligemment avec runs_data
                    has_runs_data = all(seg.get("runs_data") for seg in para_segments)
                    
                    if has_runs_data and len(para_segments) > 1:
                        combined_runs = []
                        for seg in para_segments:
                            try:
                                runs = json.loads(seg.get("runs_data", "[]"))
                                combined_runs.extend(runs)
                                # Ajouter un espace entre les phrases
                                combined_runs.append({
                                    "texts": [{"type": "text", "value": " "}],
                                    "props": {}
                                })
                            except:
                                pass
                        
                        first_seg = para_segments[0]
                        combined_seg = {**first_seg, "runs_data": json.dumps(combined_runs)}
                        parts.append(paragraph_xml_enhanced(combined_seg))
                    else:
                        # Traiter chaque segment individuellement
                        for seg in para_segments:
                            parts.append(paragraph_xml_enhanced(seg))
            else:
                parts.append('<w:p><w:pPr/></w:p>')
            
            parts.append('</w:tc>')
        parts.append('</w:tr>')
    
    parts.append('</w:tbl>')
    return ''.join(parts)

def build_document_xml_enhanced(segments):
    """
    ✨ VERSION AMÉLIORÉE ✨
    Construit le document.xml avec support du formatage mixte.
    """
    body_paragraphs = {}
    tables = {}
    
    for s in segments:
        if isinstance(s, str):
            try:
                s = json.loads(s)
            except:
                continue
        if not isinstance(s, dict):
            continue
        
        seg = {**s}
        seg["order"] = to_int(seg.get("order")) or 0
        seg["table_index"] = to_int(seg.get("table_index"))
        seg["row_index"] = to_int(seg.get("row_index"))
        seg["col_index"] = to_int(seg.get("col_index"))
        seg["paragraph_id"] = to_str(seg.get("paragraph_id")) or str(seg["order"])
        
        if seg["table_index"]:
            tables.setdefault(seg["table_index"], []).append(seg)
        else:
            para_id = seg["paragraph_id"]
            if para_id not in body_paragraphs:
                body_paragraphs[para_id] = []
            body_paragraphs[para_id].append(seg)

    parts = []
    parts.append(xml_header())
    parts.append('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ')
    parts.append('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    parts.append('<w:body>')

    # Trier les paragraphes
    sorted_para_ids = sorted(
        body_paragraphs.keys(),
        key=lambda pid: min(seg["order"] for seg in body_paragraphs[pid])
    )
    
    for para_id in sorted_para_ids:
        para_segments = sorted(body_paragraphs[para_id], key=lambda x: x["order"])
        
        # Vérifier si on doit combiner les segments
        has_runs_data = all(seg.get("runs_data") for seg in para_segments)
        
        if has_runs_data and len(para_segments) > 1:
            # Combiner les runs_data de tous les segments
            combined_runs = []
            for seg in para_segments:
                try:
                    runs = json.loads(seg.get("runs_data", "[]"))
                    combined_runs.extend(runs)
                    # Ajouter un espace entre les phrases
                    combined_runs.append({
                        "texts": [{"type": "text", "value": " "}],
                        "props": {}
                    })
                except:
                    pass
            
            # Retirer le dernier espace ajouté
            if combined_runs and combined_runs[-1].get("texts", [{}])[0].get("value") == " ":
                combined_runs.pop()
            
            first_seg = para_segments[0]
            combined_seg = {**first_seg, "runs_data": json.dumps(combined_runs)}
            parts.append(paragraph_xml_enhanced(combined_seg))
        else:
            # Fallback: utiliser content_source combiné
            text_parts = []
            for seg in para_segments:
                content = seg.get("content_source")
                if content is not None and content != "":
                    text_parts.append(str(content))
            
            combined_text = " ".join(text_parts) if text_parts else ""
            first_seg = para_segments[0]
            combined_seg = {**first_seg, "content_source": combined_text}
            parts.append(paragraph_xml_enhanced(combined_seg))

    # Tableaux
    for t_idx in sorted(tables.keys()):
        tseg = tables[t_idx]
        max_r = max(s["row_index"] or 1 for s in tseg)
        max_c = max(s["col_index"] or 1 for s in tseg)
        cells = {}
        for s in tseg:
            r = s["row_index"] or 1
            c = s["col_index"] or 1
            cells.setdefault((r, c), []).append(s)
        for k in list(cells.keys()):
            cells[k] = sorted(cells[k], key=lambda seg: seg["order"])
        parts.append(table_xml_enhanced(cells, max_r, max_c))

    parts.append('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>')
    parts.append('</w:body></w:document>')
    return ''.join(parts)

def build_docx_zip(document_meta, segments):
    """Construit le fichier DOCX complet"""
    files = {
        "[Content_Types].xml": content_types_xml().encode("utf-8"),
        "_rels/.rels": rels_root_xml().encode("utf-8"),
        "docProps/core.xml": core_xml(
            creator=(document_meta.get("client") or "Make"),
            title=(document_meta.get("file_name") or "Rebuilt"),
            subject=document_meta.get("project") or "",
            description="Rebuilt with full formatting preservation v2.0"
        ).encode("utf-8"),
        "docProps/app.xml": app_xml().encode("utf-8"),
        "word/_rels/document.xml.rels": document_rels_xml().encode("utf-8"),
        "word/styles.xml": styles_xml().encode("utf-8"),
        "word/numbering.xml": numbering_xml().encode("utf-8"),
        "word/document.xml": build_document_xml_enhanced(segments).encode("utf-8"),
    }
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
        for path, data in files.items():
            z.writestr(path, data)
    bio.seek(0)
    return bio.read()


# ==================== API ENDPOINTS ====================

@app.post("/reconstruct-document")
async def reconstruct_document_endpoint(
    request: ReconstructDocumentRequest,
    x_api_key: str = Header(None)
):
    """
    Reconstruit un document Word à partir de segments traduits
    Retourne le document encodé en base64
    """
    try:
        # Vérifier l'API key
        verify_api_key(x_api_key)
        
        logger.info(f"Starting reconstruction for: {request.file_name}")
        start_time = datetime.datetime.utcnow()
        
        # Extraire les données
        document_meta = request.document_metadata or {}
        segments = request.segments
        
        if not segments:
            raise ValueError("No segments provided")
        
        # Construire le document
        docx_bytes = build_docx_zip(document_meta, segments)
        file_base64 = base64.b64encode(docx_bytes).decode("utf-8")
        
        # Générer le nom du fichier
        original_filename = request.file_name or "output.docx"
        base_name = original_filename.replace(".docx", "").replace(".DOCX", "")
        file_name = base_name + "_rebuilt.docx"
        
        end_time = datetime.datetime.utcnow()
        processing_time = (end_time - start_time).total_seconds()
        
        logger.info(f"Reconstruction completed in {processing_time:.2f}s")
        
        return {
            "success": True,
            "file_name": file_name,
            "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "file_base64": file_base64,
            "segment_count": len(segments),
            "processing_time_seconds": round(processing_time, 2)
        }
        
    except ValueError as e:
        logger.error(f"Validation error: {e}")
        raise HTTPException(status_code=400, detail=str(e))
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