import html
import io
import json
import math
import os
import re
import urllib.request
import zipfile
import pandas as pd
import streamlit as st

# Safe handling for optional PDF/Word document extractors
try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import docx
except ImportError:
    docx = None

st.set_page_config(
    page_title="Aryavarta Automation - Interactive CAD GA & SLD Engine",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
div[data-testid="stMetricValue"],
div[data-testid="stMetricValue"] *,
div[data-testid="stMetricValue"] > div,
div[data-testid="stMetricValue"] span,
div[data-testid="stMetricValue"] p {
    white-space: normal !important;
    word-break: break-word !important;
    overflow-wrap: break-word !important;
    text-overflow: unset !important;
    overflow: visible !important;
    font-size: 1.15rem !important;
    line-height: 1.3 !important;
    font-weight: 700 !important;
}

div[data-testid="stMetricLabel"],
div[data-testid="stMetricLabel"] *,
div[data-testid="stMetricLabel"] > div,
div[data-testid="stMetricLabel"] label,
div[data-testid="stMetricLabel"] p {
    white-space: normal !important;
    word-break: break-word !important;
    overflow-wrap: break-word !important;
    text-overflow: unset !important;
    overflow: visible !important;
    font-size: 0.88rem !important;
    font-weight: 600 !important;
}

div[data-testid="stMetric"] {
    background-color: #f8fafc;
    border: 1px solid #e2e8f0;
    padding: 12px 14px;
    border-radius: 8px;
    min-height: 95px;
    height: auto !important;
}

@media (prefers-color-scheme: dark) {
    div[data-testid="stMetric"] {
        background-color: #1e293b;
        border: 1px solid #334155;
    }
}
</style>
""", unsafe_allow_html=True)

def xml_escape(text):
    """Safely escapes XML string entities to prevent SVG render errors."""
    return html.escape(str(text))

COMPONENT_DIMENSIONS_LIBRARY = {
    "200A 3P MCCB": {"Height_mm": 185, "Width_mm": 105, "Depth_mm": 88, "Type": "MCCB"},
    "100A 3P MCCB": {"Height_mm": 165, "Width_mm": 105, "Depth_mm": 86, "Type": "MCCB"},
    "630A 3P MCCB": {"Height_mm": 275, "Width_mm": 140, "Depth_mm": 103, "Type": "MCCB"},
    "MPCB (5.5-8A)": {"Height_mm": 90, "Width_mm": 45, "Depth_mm": 75, "Type": "MPCB"},
    "Power Contactor (12A)": {"Height_mm": 85, "Width_mm": 45, "Depth_mm": 80, "Type": "Contactor"},
    "VFD Drive (3.7 kW)": {"Height_mm": 220, "Width_mm": 110, "Depth_mm": 150, "Type": "VFD"},
    "PLC ET200SP CPU": {"Height_mm": 117, "Width_mm": 100, "Depth_mm": 75, "Type": "PLC"},
    "PLC I/O Module": {"Height_mm": 117, "Width_mm": 15, "Depth_mm": 75, "Type": "PLC"},
    "PLC Slim Relay": {"Height_mm": 90, "Width_mm": 6, "Depth_mm": 75, "Type": "Relay"},
    "Janatics Valve Manifold": {"Height_mm": 120, "Width_mm": 240, "Depth_mm": 80, "Type": "Pneumatic"},
    "Pneumatic PVC JB": {"Height_mm": 400, "Width_mm": 300, "Depth_mm": 150, "Type": "Enclosure"},
    "Digital KWH Meter": {"Height_mm": 96, "Width_mm": 96, "Depth_mm": 60, "Type": "Meter"},
    "Salzer Switch": {"Height_mm": 90, "Width_mm": 90, "Depth_mm": 70, "Type": "Switch"},
    "Power TB Strip": {"Height_mm": 60, "Width_mm": 250, "Depth_mm": 45, "Type": "Terminal"},
    "Control TB Strip": {"Height_mm": 50, "Width_mm": 300, "Depth_mm": 45, "Type": "Terminal"}
}

def extract_text_from_file(uploaded_file):
    if uploaded_file is None:
        return ""
    fname = uploaded_file.name.lower()
    text = ""
    try:
        if fname.endswith(".txt"):
            text = uploaded_file.read().decode("utf-8", errors="ignore")
        elif fname.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            text = df.to_string()
        elif fname.endswith(".xlsx") or fname.endswith(".xls"):
            xls = pd.ExcelFile(uploaded_file)
            sheet_texts = []
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                sheet_texts.append(f"--- Sheet: {sheet_name} ---\n" + df.to_string())
            text = "\n\n".join(sheet_texts)
        elif fname.endswith(".pdf"):
            try:
                import pypdf
                reader = pypdf.PdfReader(uploaded_file)
                text = "\n".join([page.extract_text() or "" for page in reader.pages])
            except Exception:
                content = uploaded_file.read().decode("latin1", errors="ignore")
                text = " ".join(re.findall(r"\((.*?)\)", content))
        elif fname.endswith(".docx"):
            try:
                import xml.etree.ElementTree as ET
                with zipfile.ZipFile(uploaded_file) as z:
                    xml_content = z.read("word/document.xml")
                    tree = ET.fromstring(xml_content)
                    text = " ".join([elem.text for elem in tree.iter() if elem.text])
            except Exception:
                text = uploaded_file.read().decode("utf-8", errors="ignore")
    except Exception as e:
        text = f"Error reading file: {str(e)}"
    return text

def parse_bom_items_from_text(raw_text):
    text_lower = raw_text.lower()
    
    # Defaults based on 05_AAQ_3PS_DIG_ 26706_030 R4
    mccb_rating = "200A 3P MCCB"
    if "630a" in text_lower: mccb_rating = "630A 3P MCCB"
    elif "400a" in text_lower: mccb_rating = "400A 3P MCCB"
    elif "100a" in text_lower: mccb_rating = "100A 3P MCCB"
    
    dol_match = re.search(r'(\d+)\s*(?:nos|qty)?\s*dol', text_lower)
    dol_qty = int(dol_match.group(1)) if dol_match else 19

    vfd_match = re.search(r'(\d+)\s*(?:nos|qty)?\s*vfd', text_lower)
    vfd_qty = int(vfd_match.group(1)) if vfd_match else 2

    panel_h = 1800
    panel_w = 1600
    part1_w = 1000
    part2_w = 600
    panel_d = 450

    fab_match = re.search(r'(\d+)\s*\(h\)\s*\*\s*(\d+)\s*\(w\)\s*\*\s*(\d+)\s*\(d\)', text_lower)
    if fab_match:
        panel_h = int(fab_match.group(1))
        panel_w = int(fab_match.group(2))
        panel_d = int(fab_match.group(3))
        part1_w = int(panel_w * 0.625)
        part2_w = panel_w - part1_w

    return {
        "panel_h": panel_h,
        "panel_w": panel_w,
        "part1_w": part1_w,
        "part2_w": part2_w,
        "panel_d": panel_d,
        "mccb_rating": mccb_rating,
        "dol_qty": dol_qty,
        "vfd_qty": vfd_qty
    }

def auto_layout_bom_to_cad(parsed_bom):
    items = []
    
    p_h = parsed_bom["panel_h"]
    p1_w = parsed_bom["part1_w"]
    p2_w = parsed_bom["part2_w"]
    
    # -------------------------------------------------------------
    # BAY 1 (MCC SECTION - 1000mm Width)
    # -------------------------------------------------------------
    # Incomer Chamber (Top 350mm)
    items.append({
        "Tag": "INCOMER", "Item Name": parsed_bom["mccb_rating"], "Type": "MCCB",
        "Bay": 1, "Pos X (mm)": 100.0, "Pos Y (mm)": 80.0, "Width (mm)": 140.0, "Height (mm)": 200.0,
        "Mounting": "Interior"
    })
    items.append({
        "Tag": "CT_SET", "Item Name": "Incomer CT 250/5 Set", "Type": "CT",
        "Bay": 1, "Pos X (mm)": 270.0, "Pos Y (mm)": 80.0, "Width (mm)": 180.0, "Height (mm)": 70.0,
        "Mounting": "Interior"
    })
    items.append({
        "Tag": "SPP", "Item Name": "VSP D2 Single Phase Preventer", "Type": "Relay",
        "Bay": 1, "Pos X (mm)": 480.0, "Pos Y (mm)": 80.0, "Width (mm)": 75.0, "Height (mm)": 90.0,
        "Mounting": "Interior"
    })
    
    # DOL Feeders (19x Feeders in rows)
    dol_qty = parsed_bom["dol_qty"]
    start_y = 380.0
    cur_x = 80.0
    cur_y = start_y
    row_max_h = 160.0
    
    for i in range(1, dol_qty + 1):
        if cur_x + 90.0 > (p1_w - 80.0):
            cur_x = 80.0
            cur_y += row_max_h + 30.0
            
        items.append({
            "Tag": f"F{i}_MPCB", "Item Name": f"DOL F{i} MPCB 5.5-8A", "Type": "MPCB",
            "Bay": 1, "Pos X (mm)": cur_x, "Pos Y (mm)": cur_y, "Width (mm)": 45.0, "Height (mm)": 85.0,
            "Mounting": "Interior"
        })
        items.append({
            "Tag": f"F{i}_CONT", "Item Name": f"DOL F{i} Contactor 12A", "Type": "Contactor",
            "Bay": 1, "Pos X (mm)": cur_x, "Pos Y (mm)": cur_y + 90.0, "Width (mm)": 45.0, "Height (mm)": 75.0,
            "Mounting": "Interior"
        })
        cur_x += 70.0

    # VFD Feeders (2x Feeders)
    vfd_qty = parsed_bom["vfd_qty"]
    cur_y += 180.0
    cur_x = 80.0
    for v in range(1, vfd_qty + 1):
        f_num = dol_qty + v
        items.append({
            "Tag": f"F{f_num}_MCB", "Item Name": f"VFD F{f_num} MCB 10A 3P", "Type": "MCB",
            "Bay": 1, "Pos X (mm)": cur_x, "Pos Y (mm)": cur_y, "Width (mm)": 54.0, "Height (mm)": 85.0,
            "Mounting": "Interior"
        })
        items.append({
            "Tag": f"F{f_num}_VFD", "Item Name": f"VFD F{f_num} Drive 3.7kW", "Type": "VFD",
            "Bay": 1, "Pos X (mm)": cur_x + 65.0, "Pos Y (mm)": cur_y, "Width (mm)": 110.0, "Height (mm)": 200.0,
            "Mounting": "Interior"
        })
        cur_x += 210.0

    # Power Terminal Strips (Bottom Bay 1)
    items.append({
        "Tag": "POWER_TB", "Item Name": "Power TB 10mm² / 25mm² Array", "Type": "Terminal",
        "Bay": 1, "Pos X (mm)": 80.0, "Pos Y (mm)": p_h - 140.0, "Width (mm)": p1_w - 160.0, "Height (mm)": 60.0,
        "Mounting": "Interior"
    })

    # -------------------------------------------------------------
    # BAY 2 (PLC & PNEUMATIC SECTION - 600mm Width)
    # -------------------------------------------------------------
    # PLC ET200SP Rack Top
    items.append({
        "Tag": "PLC_CPU", "Item Name": "Siemens ET200SP CPU + I/O Modules", "Type": "PLC",
        "Bay": 2, "Pos X (mm)": 50.0, "Pos Y (mm)": 80.0, "Width (mm)": 320.0, "Height (mm)": 117.0,
        "Mounting": "Interior"
    })
    items.append({
        "Tag": "PLC_RELAYS", "Item Name": "48x PLC Slim Interlocking Relays", "Type": "Relay",
        "Bay": 2, "Pos X (mm)": 50.0, "Pos Y (mm)": 230.0, "Width (mm)": 380.0, "Height (mm)": 90.0,
        "Mounting": "Interior"
    })
    
    # Pneumatic Section
    items.append({
        "Tag": "PNEUMATIC_JB1", "Item Name": "Janatics 12-Way Valve Manifold JB", "Type": "Pneumatic",
        "Bay": 2, "Pos X (mm)": 50.0, "Pos Y (mm)": 400.0, "Width (mm)": 260.0, "Height (mm)": 140.0,
        "Mounting": "Interior"
    })
    items.append({
        "Tag": "PNEUMATIC_JB2", "Item Name": "Janatics 8-Way Valve Manifold JB", "Type": "Pneumatic",
        "Bay": 2, "Pos X (mm)": 50.0, "Pos Y (mm)": 570.0, "Width (mm)": 220.0, "Height (mm)": 140.0,
        "Mounting": "Interior"
    })

    # Control Terminal Strips (Bottom Bay 2)
    items.append({
        "Tag": "CTRL_TB_FUSE", "Item Name": "110x Control Fuse TB 4mm² Strip", "Type": "Terminal",
        "Bay": 2, "Pos X (mm)": 50.0, "Pos Y (mm)": p_h - 220.0, "Width (mm)": p2_w - 100.0, "Height (mm)": 50.0,
        "Mounting": "Interior"
    })
    items.append({
        "Tag": "CTRL_TB_NONFUSE", "Item Name": "220x Control Non-Fuse TB 4mm² Strip", "Type": "Terminal",
        "Bay": 2, "Pos X (mm)": 50.0, "Pos Y (mm)": p_h - 140.0, "Width (mm)": p2_w - 100.0, "Height (mm)": 50.0,
        "Mounting": "Interior"
    })

    # -------------------------------------------------------------
    # DOOR MOUNTED ITEMS (GA OUTER)
    # -------------------------------------------------------------
    items.append({
        "Tag": "SALZER_SWITCH", "Item Name": "Salzer Main ON/OFF Rotary Switch", "Type": "DoorItem",
        "Bay": 1, "Pos X (mm)": 80.0, "Pos Y (mm)": 100.0, "Width (mm)": 80.0, "Height (mm)": 80.0,
        "Mounting": "Door"
    })
    items.append({
        "Tag": "KWH_METER", "Item Name": "Digital KWH Meter (96x96mm)", "Type": "DoorItem",
        "Bay": 1, "Pos X (mm)": 200.0, "Pos Y (mm)": 90.0, "Width (mm)": 96.0, "Height (mm)": 96.0,
        "Mounting": "Door"
    })
    items.append({
        "Tag": "LAMPS_RYB", "Item Name": "Phase Indicating Lamps (R, Y, B)", "Type": "DoorItem",
        "Bay": 1, "Pos X (mm)": 340.0, "Pos Y (mm)": 110.0, "Width (mm)": 120.0, "Height (mm)": 30.0,
        "Mounting": "Door"
    })
    items.append({
        "Tag": "LOUVER_FAN1", "Item Name": "4\" Louver & Filter Fan Unit (Top)", "Type": "DoorItem",
        "Bay": 1, "Pos X (mm)": p1_w - 180.0, "Pos Y (mm)": 80.0, "Width (mm)": 120.0, "Height (mm)": 120.0,
        "Mounting": "Door"
    })
    items.append({
        "Tag": "LOUVER_FAN2", "Item Name": "4\" Louver & Filter Fan Unit (Bottom)", "Type": "DoorItem",
        "Bay": 1, "Pos X (mm)": p1_w - 180.0, "Pos Y (mm)": p_h - 220.0, "Width (mm)": 120.0, "Height (mm)": 120.0,
        "Mounting": "Door"
    })
    
    # Door 2 Items
    items.append({
        "Tag": "PLC_24V_LAMP", "Item Name": "White 24V DC Supply ON Lamp", "Type": "DoorItem",
        "Bay": 2, "Pos X (mm)": 60.0, "Pos Y (mm)": 100.0, "Width (mm)": 30.0, "Height (mm)": 30.0,
        "Mounting": "Door"
    })
    items.append({
        "Tag": "KEY_SELECTOR", "Item Name": "Teknic Key Selector Switch", "Type": "DoorItem",
        "Bay": 2, "Pos X (mm)": 120.0, "Pos Y (mm)": 100.0, "Width (mm)": 40.0, "Height (mm)": 40.0,
        "Mounting": "Door"
    })
    items.append({
        "Tag": "ALARM_BUZZER", "Item Name": "24V AC/DC Alarm Buzzer", "Type": "DoorItem",
        "Bay": 2, "Pos X (mm)": 180.0, "Pos Y (mm)": 100.0, "Width (mm)": 40.0, "Height (mm)": 40.0,
        "Mounting": "Door"
    })

    return pd.DataFrame(items)

def render_svg_title_block(start_x, start_y, width, height, project_no="AAP_260706_005", drg_no="01", sheet_title="MCC + PLC CONTROL PANEL"):
    svg = []
    svg.append(f'<g id="TitleBlock">')
    svg.append(f'<rect x="{start_x}" y="{start_y}" width="{width}" height="{height}" fill="#1e293b" stroke="#cbd5e1" stroke-width="2"/>')
    svg.append(f'<line x1="{start_x}" y1="{start_y + 35}" x2="{start_x + width}" y2="{start_y + 35}" stroke="#64748b" stroke-width="1.5"/>')
    svg.append(f'<line x1="{start_x}" y1="{start_y + 70}" x2="{start_x + width}" y2="{start_y + 70}" stroke="#64748b" stroke-width="1.5"/>')
    svg.append(f'<line x1="{start_x + width/2}" y1="{start_y + 35}" x2="{start_x + width/2}" y2="{start_y + height}" stroke="#64748b" stroke-width="1.5"/>')
    
    # Title Text
    svg.append(f'<text x="{start_x + width/2}" y="{start_y + 24}" fill="#38bdf8" font-size="14" font-weight="bold" text-anchor="middle">{xml_escape(sheet_title)}</text>')
    svg.append(f'<text x="{start_x + 10}" y="{start_y + 55}" fill="#f8fafc" font-size="11" font-weight="bold">CONSULTANT: ARYAVARTA AUTOMATION</text>')
    svg.append(f'<text x="{start_x + width/2 + 10}" y="{start_y + 55}" fill="#f8fafc" font-size="11" font-weight="bold">CLIENT: 3PS ENGINEERS PVT. LTD.</text>')
    svg.append(f'<text x="{start_x + 10}" y="{start_y + 90}" fill="#94a3b8" font-size="10">PROJ NO: {xml_escape(project_no)}</text>')
    svg.append(f'<text x="{start_x + width/2 + 10}" y="{start_y + 90}" fill="#94a3b8" font-size="10">DRG NO: {xml_escape(drg_no)} | SCALE: NTS</text>')
    svg.append(f'</g>')
    return "".join(svg)

def generate_internal_ga_svg(cad_df, panel_h=1800, part1_w=1000, part2_w=600):
    total_w = part1_w + part2_w
    svg_w = total_w + 240
    svg_h = panel_h + 260
    
    svg = [f'<svg width="100%" height="auto" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#090d16; border-radius:8px; font-family:sans-serif;">']
    
    # Title Header
    svg.append(f'<text x="{svg_w/2}" y="35" fill="#38bdf8" font-size="22" font-weight="bold" text-anchor="middle">INTERNAL GENERAL ARRANGEMENT (GA) - GA INTERNAL r2 STANDARD</text>')
    
    ox = 80
    oy = 70
    
    # Outer Frame
    svg.append(f'<rect x="{ox}" y="{oy}" width="{total_w}" height="{panel_h}" fill="#020617" stroke="#38bdf8" stroke-width="4"/>')
    # Partition Line
    svg.append(f'<line x1="{ox + part1_w}" y1="{oy}" x2="{ox + part1_w}" y2="{oy + panel_h}" stroke="#38bdf8" stroke-width="3" stroke-dasharray="8,4"/>')
    
    # Top Chamber Division (350mm)
    svg.append(f'<line x1="{ox}" y1="{oy + 350}" x2="{ox + part1_w}" y2="{oy + 350}" stroke="#64748b" stroke-width="2"/>')
    svg.append(f'<text x="{ox + part1_w/2}" y="{oy + 30}" fill="#94a3b8" font-size="12" font-weight="bold" text-anchor="middle">MAIN INCOMER / BUSBAR CHAMBER (350mm)</text>')
    
    # Wire Ducts (60x80mm, 60x60mm)
    for duct_x in [ox + 20, ox + part1_w - 50, ox + part1_w + 20, ox + total_w - 50]:
        svg.append(f'<rect x="{duct_x}" y="{oy + 360}" width="30" height="{panel_h - 520}" fill="#1e293b" stroke="#475569" stroke-width="1.5" stroke-dasharray="4,2"/>')
        svg.append(f'<text x="{duct_x + 15}" y="{oy + panel_h/2}" fill="#64748b" font-size="9" text-anchor="middle" transform="rotate(-90,{duct_x + 15},{oy + panel_h/2})">WIRE DUCT 60x80mm</text>')

    # Render Internal Components
    int_comps = cad_df[cad_df["Mounting"] == "Interior"] if not cad_df.empty else pd.DataFrame()
    for _, c in int_comps.iterrows():
        bay = int(c["Bay"])
        bx = ox if bay == 1 else (ox + part1_w)
        cx = bx + float(c["Pos X (mm)"])
        cy = oy + float(c["Pos Y (mm)"])
        cw = float(c["Width (mm)"])
        ch = float(c["Height (mm)"])
        c_name = xml_escape(c["Item Name"])
        c_tag = xml_escape(c["Tag"])
        c_type = str(c["Type"])
        
        color = "#2563eb" if c_type == "VFD" else ("#059669" if c_type in ["MCCB","MCB","MPCB"] else ("#d97706" if c_type == "Contactor" else ("#7c3aed" if c_type == "PLC" else "#334155")))
        
        svg.append(f'<g><title>{c_name} ({c_tag})</title>')
        svg.append(f'<rect x="{cx}" y="{cy}" width="{cw}" height="{ch}" rx="4" fill="{color}" stroke="#ffffff" stroke-width="1.5"/>')
        svg.append(f'<text x="{cx + cw/2}" y="{cy + ch/2 + 4}" fill="#ffffff" font-size="10" font-weight="bold" text-anchor="middle">{c_tag}</text>')
        svg.append('</g>')

    # Dimension Annotations
    svg.append(f'<line x1="{ox - 30}" y1="{oy}" x2="{ox - 30}" y2="{oy + panel_h}" stroke="#ef4444" stroke-width="2"/>')
    svg.append(f'<text x="{ox - 45}" y="{oy + panel_h/2}" fill="#ef4444" font-size="14" font-weight="bold" text-anchor="middle" transform="rotate(-90,{ox - 45},{oy + panel_h/2})">{panel_h} mm (HEIGHT)</text>')

    svg.append(f'<line x1="{ox}" y1="{oy + panel_h + 30}" x2="{ox + total_w}" y2="{oy + panel_h + 30}" stroke="#ef4444" stroke-width="2"/>')
    svg.append(f'<text x="{ox + total_w/2}" y="{oy + panel_h + 50}" fill="#ef4444" font-size="14" font-weight="bold" text-anchor="middle">{total_w} mm ({part1_w}mm + {part2_w}mm PARTITIONS)</text>')

    # Title Block
    svg.append(render_svg_title_block(ox + total_w - 380, oy + panel_h - 120, 380, 120, "AAP_260706_005", "01", "MCC + PLC INTERNAL GA DIAGRAM"))

    svg.append('</svg>')
    return "".join(svg)

def generate_outer_ga_svg(cad_df, panel_h=1800, part1_w=1000, part2_w=600):
    total_w = part1_w + part2_w
    svg_w = total_w + 240
    svg_h = panel_h + 260
    
    svg = [f'<svg width="100%" height="auto" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#0b0f19; border-radius:8px; font-family:sans-serif;">']
    svg.append(f'<text x="{svg_w/2}" y="35" fill="#38bdf8" font-size="22" font-weight="bold" text-anchor="middle">OUTER ELEVATION GA VIEW - GA OUTER R4 STANDARD</text>')
    
    ox = 80
    oy = 70
    
    # Outer Door Frames
    svg.append(f'<rect x="{ox}" y="{oy}" width="{part1_w}" height="{panel_h}" fill="#0f172a" stroke="#38bdf8" stroke-width="4"/>')
    svg.append(f'<rect x="{ox + part1_w}" y="{oy}" width="{part2_w}" height="{panel_h}" fill="#0f172a" stroke="#38bdf8" stroke-width="4"/>')
    
    # Door Hinges & Handles
    svg.append(f'<rect x="{ox + 10}" y="{oy + 100}" width="12" height="40" fill="#64748b"/>')
    svg.append(f'<rect x="{ox + 10}" y="{oy + panel_h - 140}" width="12" height="40" fill="#64748b"/>')
    svg.append(f'<rect x="{ox + part1_w - 25}" y="{oy + panel_h/2 - 30}" width="15" height="60" rx="4" fill="#334155" stroke="#cbd5e1" stroke-width="2"/>')

    svg.append(f'<rect x="{ox + part1_w + 10}" y="{oy + 100}" width="12" height="40" fill="#64748b"/>')
    svg.append(f'<rect x="{ox + part1_w + 10}" y="{oy + panel_h - 140}" width="12" height="40" fill="#64748b"/>')
    svg.append(f'<rect x="{ox + total_w - 25}" y="{oy + panel_h/2 - 30}" width="15" height="60" rx="4" fill="#334155" stroke="#cbd5e1" stroke-width="2"/>')

    # Render Door Mounted Components
    door_comps = cad_df[cad_df["Mounting"] == "Door"] if not cad_df.empty else pd.DataFrame()
    for _, c in door_comps.iterrows():
        bay = int(c["Bay"])
        bx = ox if bay == 1 else (ox + part1_w)
        cx = bx + float(c["Pos X (mm)"])
        cy = oy + float(c["Pos Y (mm)"])
        cw = float(c["Width (mm)"])
        ch = float(c["Height (mm)"])
        c_name = xml_escape(c["Item Name"])
        c_tag = xml_escape(c["Tag"])
        
        svg.append(f'<g><title>{c_name} ({c_tag})</title>')
        if "LAMP" in c_tag:
            for i, col in enumerate(["#ef4444", "#eab308", "#3b82f6"]):
                svg.append(f'<circle cx="{cx + 15 + i*35}" cy="{cy + 15}" r="10" fill="{col}"/>')
        elif "METER" in c_tag or "KWH" in c_tag:
            svg.append(f'<rect x="{cx}" y="{cy}" width="{cw}" height="{ch}" rx="4" fill="#020617" stroke="#38bdf8" stroke-width="2"/>')
            svg.append(f'<text x="{cx + cw/2}" y="{cy + ch/2 + 5}" fill="#38bdf8" font-family="monospace" font-size="12" font-weight="bold" text-anchor="middle">415.2 V</text>')
        elif "LOUVER" in c_tag or "FAN" in c_tag:
            svg.append(f'<rect x="{cx}" y="{cy}" width="{cw}" height="{ch}" rx="4" fill="#1e293b" stroke="#64748b" stroke-width="2"/>')
            for l_idx in range(5):
                svg.append(f'<line x1="{cx + 10}" y1="{cy + 20 + l_idx*20}" x2="{cx + cw - 10}" y2="{cy + 20 + l_idx*20}" stroke="#94a3b8" stroke-width="2"/>')
        else:
            svg.append(f'<rect x="{cx}" y="{cy}" width="{cw}" height="{ch}" rx="4" fill="#334155" stroke="#f8fafc" stroke-width="1.5"/>')
            svg.append(f'<text x="{cx + cw/2}" y="{cy + ch/2 + 4}" fill="#ffffff" font-size="10" font-weight="bold" text-anchor="middle">{c_tag}</text>')
        svg.append('</g>')

    # Dimension Annotations
    svg.append(f'<line x1="{ox}" y1="{oy + panel_h + 30}" x2="{ox + total_w}" y2="{oy + panel_h + 30}" stroke="#ef4444" stroke-width="2"/>')
    svg.append(f'<text x="{ox + total_w/2}" y="{oy + panel_h + 50}" fill="#ef4444" font-size="14" font-weight="bold" text-anchor="middle">{total_w} mm ({part1_w}mm PARTITION 1 + {part2_w}mm PARTITION 2)</text>')

    # Title Block
    svg.append(render_svg_title_block(ox + total_w - 380, oy + panel_h - 120, 380, 120, "AAP_260706_005", "01", "MCC + PLC OUTER GA ELEVATION"))

    svg.append('</svg>')
    return "".join(svg)

def generate_sld_sheet1_svg():
    svg_w = 850
    svg_h = 520
    svg = [f'<svg width="100%" height="auto" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#0f172a; border-radius:8px; font-family:sans-serif;">']
    
    svg.append(f'<text x="{svg_w/2}" y="30" fill="#38bdf8" font-size="18" font-weight="bold" text-anchor="middle">POWER SLD - SHEET 1: INCOMER & FEEDER SCHEDULE (I STANDARD)</text>')
    
    # Mains Busbar Top
    svg.append('<line x1="50" y1="70" x2="800" y2="70" stroke="#ef4444" stroke-width="3"/>')
    svg.append('<line x1="50" y1="76" x2="800" y2="76" stroke="#eab308" stroke-width="3"/>')
    svg.append('<line x1="50" y1="82" x2="800" y2="82" stroke="#3b82f6" stroke-width="3"/>')
    svg.append('<text x="425" y="60" fill="#94a3b8" font-family="monospace" font-size="11" text-anchor="middle">MAINS INCOMING: 3PH, NEUTRAL, 440V, 50Hz (35 mm sq. CU CABLE)</text>')

    # Incomer Switchgear Line
    svg.append('<line x1="425" y1="82" x2="425" y2="120" stroke="#cbd5e1" stroke-width="2.5"/>')
    svg.append('<rect x="355" y="120" width="140" height="40" rx="4" fill="#1e293b" stroke="#38bdf8" stroke-width="2"/>')
    svg.append('<text x="425" y="145" fill="#38bdf8" font-size="12" font-weight="bold" text-anchor="middle">3P MCCB 200A</text>')
    
    svg.append('<line x1="425" y1="160" x2="425" y2="200" stroke="#cbd5e1" stroke-width="2.5"/>')
    svg.append('<circle cx="425" cy="180" r="12" fill="none" stroke="#eab308" stroke-width="2"/>')
    svg.append('<text x="445" y="184" fill="#eab308" font-size="10">CT 250/5A</text>')

    # Feeder Schedule Table
    start_y = 240
    svg.append(f'<rect x="50" y="{start_y}" width="750" height="150" fill="#1e293b" stroke="#64748b" stroke-width="2"/>')
    svg.append(f'<text x="425" y="{start_y + 22}" fill="#38bdf8" font-size="12" font-weight="bold" text-anchor="middle">FEEDER SCHEDULE TABLE (FEEDERS F1 TO F9 - 3.7 kW DOL)</text>')
    
    headers = ["FEEDER NO", "CABLE SIZE", "MPCB RATING", "CONTACTOR", "TERMINAL SIZE", "MOTOR kW"]
    for idx, h in enumerate(headers):
        svg.append(f'<text x="70" y="{start_y + 50 + idx*20}" fill="#94a3b8" font-size="10" font-weight="bold">{h}</text>')
        svg.append(f'<text x="220" y="{start_y + 50 + idx*20}" fill="#f8fafc" font-size="10">F1 to F9: 2.5mm² Cu | MPCB 5.5-8A | CONT 12A | TB 6mm² | 3.7 kW Motor</text>')

    # Title Block
    svg.append(render_svg_title_block(420, 400, 380, 100, "AAP_260706_005", "01", "POWER SLD - INCOMER SHEET 1"))

    svg.append('</svg>')
    return "".join(svg)

def generate_sld_sheet2_svg():
    svg_w = 850
    svg_h = 520
    svg = [f'<svg width="100%" height="auto" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#0f172a; border-radius:8px; font-family:sans-serif;">']
    
    svg.append(f'<text x="{svg_w/2}" y="30" fill="#38bdf8" font-size="18" font-weight="bold" text-anchor="middle">POWER SLD - SHEET 2: DOL FEEDERS F10 TO F19 (J STANDARD)</text>')
    
    # Busbar Top
    svg.append('<line x1="50" y1="60" x2="800" y2="60" stroke="#ef4444" stroke-width="3"/>')
    
    # 5 Feeder Schematic Branches
    feeder_x = [120, 260, 400, 540, 680]
    for idx, fx in enumerate(feeder_x):
        f_num = 10 + idx
        svg.append(f'<line x1="{fx}" y1="60" x2="{fx}" y2="100" stroke="#cbd5e1" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="80" fill="#94a3b8" font-size="9" text-anchor="middle">F{f_num}</text>')
        
        # MPCB Symbol
        svg.append(f'<rect x="{fx-30}" y="100" width="60" height="35" rx="3" fill="#1e293b" stroke="#059669" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="122" fill="#059669" font-size="9" font-weight="bold" text-anchor="middle">MPCB 5.5-8A</text>')
        
        svg.append(f'<line x1="{fx}" y1="135" x2="{fx}" y2="175" stroke="#cbd5e1" stroke-width="2"/>')
        
        # Contactor Symbol
        svg.append(f'<rect x="{fx-30}" y="175" width="60" height="35" rx="3" fill="#1e293b" stroke="#d97706" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="197" fill="#d97706" font-size="9" font-weight="bold" text-anchor="middle">CONT 12A</text>')

        svg.append(f'<line x1="{fx}" y1="210" x2="{fx}" y2="250" stroke="#cbd5e1" stroke-width="2"/>')
        
        # Motor
        svg.append(f'<circle cx="{fx}" cy="268" r="18" fill="#020617" stroke="#38bdf8" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="273" fill="#ffffff" font-size="11" font-weight="bold" text-anchor="middle">M 3~</text>')
        svg.append(f'<text x="{fx}" y="305" fill="#38bdf8" font-size="10" font-weight="bold" text-anchor="middle">3.7 kW</text>')

    # Title Block
    svg.append(render_svg_title_block(420, 390, 380, 110, "AAP_260706_005", "01", "POWER SLD - DOL FEEDERS SHEET 2"))

    svg.append('</svg>')
    return "".join(svg)

def generate_sld_sheet3_svg():
    svg_w = 850
    svg_h = 520
    svg = [f'<svg width="100%" height="auto" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#0f172a; border-radius:8px; font-family:sans-serif;">']
    
    svg.append(f'<text x="{svg_w/2}" y="30" fill="#38bdf8" font-size="18" font-weight="bold" text-anchor="middle">POWER SLD - SHEET 3: VFD FEEDERS F2 & F3 (K STANDARD)</text>')
    
    # Busbar Top
    svg.append('<line x1="50" y1="60" x2="800" y2="60" stroke="#ef4444" stroke-width="3"/>')
    
    vfd_x = [280, 520]
    for idx, fx in enumerate(vfd_x):
        f_num = 20 + idx
        svg.append(f'<line x1="{fx}" y1="60" x2="{fx}" y2="100" stroke="#cbd5e1" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="80" fill="#94a3b8" font-size="10" text-anchor="middle">CU WIRE 2.5mm²</text>')
        
        # MCB Symbol
        svg.append(f'<rect x="{fx-35}" y="100" width="70" height="35" rx="3" fill="#1e293b" stroke="#059669" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="122" fill="#059669" font-size="10" font-weight="bold" text-anchor="middle">MCB 10A 3P</text>')
        
        svg.append(f'<line x1="{fx}" y1="135" x2="{fx}" y2="175" stroke="#cbd5e1" stroke-width="2"/>')
        
        # VFD Drive Symbol Box
        svg.append(f'<rect x="{fx-45}" y="175" width="90" height="60" rx="4" fill="#020617" stroke="#2563eb" stroke-width="2"/>')
        svg.append(f'<line x1="{fx-45}" y1="235" x2="{fx+45}" y2="175" stroke="#2563eb" stroke-width="1.5"/>')
        svg.append(f'<text x="{fx-20}" y="200" fill="#2563eb" font-size="10" font-weight="bold">VFD</text>')
        svg.append(f'<text x="{fx+10}" y="222" fill="#ffffff" font-size="12">~</text>')
        svg.append(f'<text x="{fx}" y="250" fill="#2563eb" font-size="10" font-weight="bold" text-anchor="middle">3.7 kW VFD</text>')

        svg.append(f'<line x1="{fx}" y1="255" x2="{fx}" y2="290" stroke="#cbd5e1" stroke-width="2"/>')
        
        # Motor
        svg.append(f'<circle cx="{fx}" cy="310" r="18" fill="#020617" stroke="#38bdf8" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="315" fill="#ffffff" font-size="11" font-weight="bold" text-anchor="middle">M 3~</text>')
        svg.append(f'<text x="{fx}" y="348" fill="#38bdf8" font-size="10" font-weight="bold" text-anchor="middle">3.7 kW</text>')

    # Title Block
    svg.append(render_svg_title_block(420, 390, 380, 110, "AAP_260706_005", "01", "POWER SLD - VFD FEEDERS SHEET 3"))

    svg.append('</svg>')
    return "".join(svg)

st.sidebar.title("📐 Aryavarta Automation")
st.sidebar.markdown("**Interactive CAD GA & SLD Engine**")
st.sidebar.success("🟢 100% Offline | Chikhali Engineering Plant")

st.header("📐 Auto-Generated Interactive CAD GA Drawing & Power SLD Diagram Engine")
st.markdown("Upload company Bill of Materials (BOM) files to generate exact **Internal GA**, **Outer Elevation GA**, and **Multi-Sheet Power SLD Diagrams** conforming to **Aryavarta & 3PS Engineering Standards**.")

# File Uploader
with st.expander("📁 Upload BOM File(s) - Multi-File Supported (.xlsx, .pdf, .csv, .docx)", expanded=True):
    uploaded_boms = st.file_uploader(
        "Select BOM Files (e.g. 05_AAQ_3PS_DIG_ 26706_030 R4)",
        type=["xlsx", "xls", "pdf", "csv", "docx"],
        accept_multiple_files=True
    )

combined_bom_text = ""
if uploaded_boms:
    texts = []
    for ub in uploaded_boms:
        t = extract_text_from_file(ub)
        texts.append(f"--- BOM: {ub.name} ---\n{t}")
    combined_bom_text = "\n\n".join(texts)
    st.success(f"Loaded {len(uploaded_boms)} BOM file(s) successfully!")

parsed_bom_specs = parse_bom_items_from_text(combined_bom_text)

# Auto-Initialize CAD Layout State
if "cad_layout_df" not in st.session_state or st.sidebar.button("🔄 Reset CAD Layout to Auto-Arrangement"):
    st.session_state["cad_layout_df"] = auto_layout_bom_to_cad(parsed_bom_specs)

# Display CAD Control Dashboard
st.subheader("🎛️ Interactive Panel Sizing & Partition Dimensions")
c1, c2, c3, c4 = st.columns(4)
p_h_val = c1.number_input("Panel Height (mm)", value=parsed_bom_specs["panel_h"], step=100)
p1_w_val = c2.number_input("Partition 1 Width - MCC (mm)", value=parsed_bom_specs["part1_w"], step=50)
p2_w_val = c3.number_input("Partition 2 Width - PLC (mm)", value=parsed_bom_specs["part2_w"], step=50)
p_d_val = c4.number_input("Panel Depth (mm)", value=parsed_bom_specs["panel_d"], step=50)

st.markdown("#### 📝 Edit Component Coordinates & Placement Parameters")
st.info("💡 **Live Interactive CAD:** Edit **Bay (1=MCC, 2=PLC)**, **Pos X (mm)**, **Pos Y (mm)**, **Width (mm)**, **Height (mm)**, or **Item Name** in the table below to reposition components inside the panel in real time!")

edited_cad_df = st.data_editor(
    st.session_state["cad_layout_df"],
    num_rows="dynamic",
    key="cad_layout_table_editor",
    column_config={
        "Bay": st.column_config.NumberColumn("Bay / Partition", min_value=1, max_value=2, step=1),
        "Pos X (mm)": st.column_config.NumberColumn("Pos X (mm)", step=10.0),
        "Pos Y (mm)": st.column_config.NumberColumn("Pos Y (mm)", step=10.0),
        "Width (mm)": st.column_config.NumberColumn("Width (mm)", step=5.0),
        "Height (mm)": st.column_config.NumberColumn("Height (mm)", step=5.0),
        "Mounting": st.column_config.SelectboxColumn("Mounting Area", options=["Interior", "Door"])
    }
)

st.session_state["cad_layout_df"] = edited_cad_df

st.divider()
tab_int_ga, tab_outer_ga, tab_sld_1, tab_sld_2, tab_sld_3 = st.tabs([
    "🖼️ Internal GA View (Interior Mounting)",
    "🚪 Outer Elevation GA View (Doors & Meters)",
    "⚡ Power SLD - Sheet 1 (Incomer)",
    "⚡ Power SLD - Sheet 2 (DOL Feeders)",
    "⚡ Power SLD - Sheet 3 (VFD Feeders)"
])

with tab_int_ga:
    int_svg_str = generate_internal_ga_svg(st.session_state["cad_layout_df"], p_h_val, p1_w_val, p2_w_val)
    st.markdown(int_svg_str, unsafe_allow_html=True)
    st.download_button(
        "📥 Download Internal GA (.SVG)",
        int_svg_str,
        "GA_INTERNAL_r2.svg",
        mime="image/svg+xml"
    )

with tab_outer_ga:
    outer_svg_str = generate_outer_ga_svg(st.session_state["cad_layout_df"], p_h_val, p1_w_val, p2_w_val)
    st.markdown(outer_svg_str, unsafe_allow_html=True)
    st.download_button(
        "📥 Download Outer GA (.SVG)",
        outer_svg_str,
        "GA_OUTER_R4.svg",
        mime="image/svg+xml"
    )

with tab_sld_1:
    sld1_svg_str = generate_sld_sheet1_svg()
    st.markdown(sld1_svg_str, unsafe_allow_html=True)
    st.download_button(
        "📥 Download SLD Sheet 1 (.SVG)",
        sld1_svg_str,
        "SLD_Sheet1_Incomer.svg",
        mime="image/svg+xml"
    )

with tab_sld_2:
    sld2_svg_str = generate_sld_sheet2_svg()
    st.markdown(sld2_svg_str, unsafe_allow_html=True)
    st.download_button(
        "📥 Download SLD Sheet 2 (.SVG)",
        sld2_svg_str,
        "SLD_Sheet2_DOL_Feeders.svg",
        mime="image/svg+xml"
    )

with tab_sld_3:
    sld3_svg_str = generate_sld_sheet3_svg()
    st.markdown(sld3_svg_str, unsafe_allow_html=True)
    st.download_button(
        "📥 Download SLD Sheet 3 (.SVG)",
        sld3_svg_str,
        "SLD_Sheet3_VFD_Feeders.svg",
        mime="image/svg+xml"
    )

st.success("✅ CAD GA Drawings & Power SLD Diagram Engine Loaded Successfully!")
