import io
import json
import math
import os
import re
import urllib.request
import zipfile
import pandas as pd
import streamlit as st

# Optional dependencies with safe import handling
try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import docx
except ImportError:
    docx = None

# ReportLab Imports for PDF Generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(
    page_title="Aryavarta Automation - Sales & Engineering Suite",
    page_icon="⚡",
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

DB_FILE = "switchgear_prices.xlsx"
QUOTES_FILE = "quotes_register.xlsx"

def save_quote_to_history(client, panel_type, kw, brand, price, margin):
    """Saves generated quote record to persistent quotes register Excel file."""
    new_data = pd.DataFrame([{
        "Date": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "Client": client, "Scope": f"{panel_type} ({kw})", "Brand": brand,
        "Quote_Price_INR": price, "Margin_Pct": margin, "Status": "Pending"
    }])
    df_updated = pd.concat([pd.read_excel(QUOTES_FILE), new_data], ignore_index=True) if os.path.exists(QUOTES_FILE) else new_data
    df_updated.to_excel(QUOTES_FILE, index=False)

def init_database():
    """Initializes local Excel price database if not present."""
    if not os.path.exists(DB_FILE):
        data = [
            {"Category": "VFD", "Brand": "Danfoss", "Specification": "15 kW FC-51", "Unit_Price_INR": 35000},
            {"Category": "VFD", "Brand": "Schneider", "Specification": "15 kW ATV320", "Unit_Price_INR": 38000},
            {"Category": "VFD", "Brand": "ABB", "Specification": "15 kW ACS380", "Unit_Price_INR": 41000},
            {"Category": "VFD", "Brand": "Delta", "Specification": "15 kW MS300", "Unit_Price_INR": 29000},
            {"Category": "VFD", "Brand": "Siemens", "Specification": "15 kW G120C", "Unit_Price_INR": 42000},
            {"Category": "MCCB", "Brand": "L&T", "Specification": "100A 3P 25kA DN1", "Unit_Price_INR": 4200},
            {"Category": "MCCB", "Brand": "Siemens", "Specification": "100A 3P 25kA 3VM", "Unit_Price_INR": 4800},
            {"Category": "MCCB", "Brand": "Schneider", "Specification": "100A 3P 25kA EasyPact", "Unit_Price_INR": 4600},
            {"Category": "MCCB", "Brand": "ABB", "Specification": "100A 3P 25kA Formula", "Unit_Price_INR": 4500},
            {"Category": "Contactor", "Brand": "L&T", "Specification": "32A 3P AC3 MNX", "Unit_Price_INR": 1800},
            {"Category": "Contactor", "Brand": "Siemens", "Specification": "32A 3P AC3 3TF", "Unit_Price_INR": 2100},
            {"Category": "Contactor", "Brand": "Schneider", "Specification": "32A 3P AC3 TeSys", "Unit_Price_INR": 2000},
            {"Category": "Enclosure", "Brand": "Standard Sheet Metal", "Specification": "IP54 Floor Mount (1200x800x400)", "Unit_Price_INR": 14000},
            {"Category": "Enclosure", "Brand": "Rittal", "Specification": "IP55 Floor Mount (1200x800x400)", "Unit_Price_INR": 26000},
            {"Category": "Busbar & Wire", "Brand": "Polycab", "Specification": "Copper Busbar + Harness Set", "Unit_Price_INR": 5500},
            {"Category": "Accessories", "Brand": "Generic", "Specification": "Push Buttons, Meters, Relays", "Unit_Price_INR": 3500},
        ]
        pd.DataFrame(data).to_excel(DB_FILE, index=False)

init_database()

@st.cache_data(ttl=2)
def load_prices():
    return pd.read_excel(DB_FILE)

df_prices = load_prices()

COMPONENT_DIMENSIONS_LIBRARY = {
    "100A 3P MCCB": {"Height_mm": 165, "Width_mm": 105, "Depth_mm": 86, "Type": "MCCB"},
    "250A 3P MCCB": {"Height_mm": 185, "Width_mm": 105, "Depth_mm": 88, "Type": "MCCB"},
    "400A 3P MCCB": {"Height_mm": 275, "Width_mm": 140, "Depth_mm": 103, "Type": "MCCB"},
    "630A 3P MCCB": {"Height_mm": 275, "Width_mm": 140, "Depth_mm": 103, "Type": "MCCB"},
    "800A 4P Drawout ACB": {"Height_mm": 430, "Width_mm": 410, "Depth_mm": 355, "Type": "ACB"},
    "1250A 4P Drawout ACB": {"Height_mm": 430, "Width_mm": 410, "Depth_mm": 355, "Type": "ACB"},
    "1600A 4P Drawout ACB": {"Height_mm": 430, "Width_mm": 510, "Depth_mm": 355, "Type": "ACB"},
    "VFD Drive (7.5 kW)": {"Height_mm": 290, "Width_mm": 150, "Depth_mm": 170, "Type": "VFD"},
    "VFD Drive (15 kW)": {"Height_mm": 350, "Width_mm": 180, "Depth_mm": 185, "Type": "VFD"},
    "VFD Drive (22 kW)": {"Height_mm": 410, "Width_mm": 220, "Depth_mm": 210, "Type": "VFD"},
    "VFD Drive (37 kW)": {"Height_mm": 520, "Width_mm": 270, "Depth_mm": 240, "Type": "VFD"},
    "VFD Drive (55 kW)": {"Height_mm": 650, "Width_mm": 320, "Depth_mm": 290, "Type": "VFD"},
    "VFD Drive (75 kW)": {"Height_mm": 750, "Width_mm": 380, "Depth_mm": 320, "Type": "VFD"},
    "Soft Starter (45 kW)": {"Height_mm": 320, "Width_mm": 180, "Depth_mm": 210, "Type": "SoftStarter"},
    "Power Contactor (32A AC3)": {"Height_mm": 85, "Width_mm": 45, "Depth_mm": 95, "Type": "Contactor"},
    "Power Contactor (65A AC3)": {"Height_mm": 125, "Width_mm": 70, "Depth_mm": 125, "Type": "Contactor"},
    "Power Contactor (110A AC3)": {"Height_mm": 160, "Width_mm": 120, "Depth_mm": 150, "Type": "Contactor"},
    "Thermal Overload Relay (TOR)": {"Height_mm": 90, "Width_mm": 45, "Depth_mm": 85, "Type": "Relay"},
    "Star-Delta Timer Relay": {"Height_mm": 90, "Width_mm": 22.5, "Depth_mm": 75, "Type": "Relay"},
    "PLC Base CPU Unit": {"Height_mm": 100, "Width_mm": 130, "Depth_mm": 75, "Type": "PLC"},
    "PLC I/O Expansion Module": {"Height_mm": 100, "Width_mm": 45, "Depth_mm": 75, "Type": "PLC"},
    "Touchscreen HMI (7 Inch)": {"Height_mm": 140, "Width_mm": 200, "Depth_mm": 45, "Type": "HMI"},
    "Touchscreen HMI (10 Inch)": {"Height_mm": 200, "Width_mm": 270, "Depth_mm": 50, "Type": "HMI"},
    "24V DC SMPS (10A)": {"Height_mm": 125, "Width_mm": 60, "Depth_mm": 125, "Type": "PowerSupply"},
    "Control Tx (500VA)": {"Height_mm": 140, "Width_mm": 120, "Depth_mm": 110, "Type": "Transformer"},
    "3% AC Line Reactor Choke": {"Height_mm": 210, "Width_mm": 180, "Depth_mm": 150, "Type": "Reactor"},
    "Digital MFM Meter": {"Height_mm": 96, "Width_mm": 96, "Depth_mm": 60, "Type": "Meter"},
    "Metering CT Set (Set of 3)": {"Height_mm": 110, "Width_mm": 240, "Depth_mm": 60, "Type": "CT"},
    "Space Heater + Thermostat": {"Height_mm": 120, "Width_mm": 80, "Depth_mm": 50, "Type": "Heater"},
    "Panel Cooling Filter Fan Unit": {"Height_mm": 250, "Width_mm": 250, "Depth_mm": 120, "Type": "Fan"},
    "Power DIN Terminal Block Strip": {"Height_mm": 60, "Width_mm": 250, "Depth_mm": 45, "Type": "Terminal"}
}

def get_incomer_default_components(incomer_type, brand="Siemens", ex_rate=1.0):
    inc_lower = str(incomer_type).lower()
    if "630" in inc_lower:
        return [
            {"Item": "Main Incomer 630A 3P MCCB", "Specification": f"630A 3P 50kA Microprocessor ({brand} 3VM)", "Brand": brand, "Qty": 1, "Unit Price": 28500.0 / ex_rate},
            {"Item": "Extended Door Rotary Handle Kit (ROM)", "Specification": "630A Door Operating Mechanism + Shaft", "Brand": brand, "Qty": 1, "Unit Price": 3200.0 / ex_rate},
            {"Item": "Terminal Spreader Links Kit", "Specification": "630A Cable Spreader Set (3 Nos)", "Brand": brand, "Qty": 1, "Unit Price": 2800.0 / ex_rate},
            {"Item": "Phase Barriers / Insulating Shrouds", "Specification": "630A Inter-Phase Barrier Set", "Brand": brand, "Qty": 1, "Unit Price": 950.0 / ex_rate},
            {"Item": "Auxiliary & Alarm Contact Block", "Specification": "1NO+1NC Aux + 1 Trip Alarm Switch", "Brand": brand, "Qty": 1, "Unit Price": 1850.0 / ex_rate},
            {"Item": "Shunt Trip Release Coil", "Specification": "230V AC Remote Emergency Trip Coil", "Brand": brand, "Qty": 1, "Unit Price": 2400.0 / ex_rate},
            {"Item": "Incomer Metering CT Set", "Specification": "600/5A Class 0.5 Measuring CTs", "Brand": "Rishabh", "Qty": 1, "Unit Price": 3600.0 / ex_rate},
            {"Item": "Incomer Protection MCB", "Specification": "6A 3P 10kA Control MCB", "Brand": brand, "Qty": 1, "Unit Price": 850.0 / ex_rate}
        ]
    elif "400" in inc_lower:
        return [
            {"Item": "Main Incomer 400A 3P MCCB", "Specification": f"400A 3P 36kA ({brand})", "Brand": brand, "Qty": 1, "Unit Price": 18500.0 / ex_rate},
            {"Item": "Extended Door Rotary Handle Kit (ROM)", "Specification": "400A Door Mechanism + Shaft", "Brand": brand, "Qty": 1, "Unit Price": 2600.0 / ex_rate},
            {"Item": "Terminal Spreader Links Kit", "Specification": "400A Terminal Spreader Set", "Brand": brand, "Qty": 1, "Unit Price": 2200.0 / ex_rate},
            {"Item": "Auxiliary Contact Block", "Specification": "1NO+1NC Aux Contact", "Brand": brand, "Qty": 1, "Unit Price": 1450.0 / ex_rate},
            {"Item": "Incomer Metering CT Set", "Specification": "400/5A Class 0.5 Measuring CTs", "Brand": "Rishabh", "Qty": 1, "Unit Price": 2900.0 / ex_rate}
        ]
    elif "250" in inc_lower:
        return [
            {"Item": "Main Incomer 250A 3P MCCB", "Specification": f"250A 3P 36kA ({brand})", "Brand": brand, "Qty": 1, "Unit Price": 11500.0 / ex_rate},
            {"Item": "Extended Door Rotary Handle Kit (ROM)", "Specification": "250A ROM Kit", "Brand": brand, "Qty": 1, "Unit Price": 2100.0 / ex_rate},
            {"Item": "Incomer Metering CT Set", "Specification": "250/5A Class 0.5 CTs", "Brand": "Rishabh", "Qty": 1, "Unit Price": 2400.0 / ex_rate}
        ]
    elif "100" in inc_lower:
        return [
            {"Item": "Main Incomer 100A 3P MCCB", "Specification": f"100A 3P 25kA ({brand})", "Brand": brand, "Qty": 1, "Unit Price": 4800.0 / ex_rate},
            {"Item": "Extended Door Rotary Handle Kit (ROM)", "Specification": "100A ROM Kit", "Brand": brand, "Qty": 1, "Unit Price": 1600.0 / ex_rate}
        ]
    else: # ACB
        rating_str = "800A" if "800" in inc_lower else ("1250A" if "1250" in inc_lower else "1600A")
        acb_base_price = 95000.0 if "800" in inc_lower else (135000.0 if "1250" in inc_lower else 175000.0)
        return [
            {"Item": f"Main Incomer {rating_str} 4P Drawout ACB", "Specification": f"{rating_str} 4P 50kA ETU ({brand} 3WA)", "Brand": brand, "Qty": 1, "Unit Price": acb_base_price / ex_rate},
            {"Item": "ACB Motorized Racking & Shunt Trip Release", "Specification": "230V AC Motor Mechanism", "Brand": brand, "Qty": 1, "Unit Price": 18500.0 / ex_rate},
            {"Item": "Incomer Metering CT Set & Protection MCB", "Specification": f"{rating_str}/5A CT Set + MCB", "Brand": "Rishabh", "Qty": 1, "Unit Price": 6800.0 / ex_rate}
        ]

def get_feeder_subcomponents(feeder_type, rating_kw_str, brand="Siemens", ex_rate=1.0):
    kw_val = float(rating_kw_str.split()[0]) if rating_kw_str and rating_kw_str.split()[0].replace('.','',1).isdigit() else 15.0
    flc = (kw_val * 1000) / (1.732 * 415 * 0.85 * 0.88)
    
    if "Star-Delta" in feeder_type:
        i_phase = flc / 1.732
        main_a = math.ceil(i_phase * 1.15)
        star_a = math.ceil((flc / 3.0) * 1.15)
        mccb_a = max(32, math.ceil(flc * 1.5 / 10.0) * 10)
        return [
            {"Item": f"Star-Delta Incomer MCCB ({mccb_a}A)", "Specification": f"{mccb_a}A 3P Breaker ({brand})", "Brand": brand, "Qty": 1, "Unit Price": (3500 + mccb_a * 18) / ex_rate},
            {"Item": f"Main Power Contactor ({main_a}A AC-3)", "Specification": f"{main_a}A 3P Contactor ({brand})", "Brand": brand, "Qty": 1, "Unit Price": (1800 + main_a * 25) / ex_rate},
            {"Item": f"Delta Power Contactor ({main_a}A AC-3)", "Specification": f"{main_a}A 3P Contactor ({brand})", "Brand": brand, "Qty": 1, "Unit Price": (1800 + main_a * 25) / ex_rate},
            {"Item": f"Star Power Contactor ({star_a}A AC-3)", "Specification": f"{star_a}A 3P Contactor ({brand})", "Brand": brand, "Qty": 1, "Unit Price": (1200 + star_a * 22) / ex_rate},
            {"Item": "Star-Delta Interlock & Electronic Timer", "Specification": "Contactor Interlock + 30s Timer", "Brand": brand, "Qty": 1, "Unit Price": 3450.0 / ex_rate},
            {"Item": "Thermal Overload Relay (TOR)", "Specification": "Adjustable Bimetallic TOR", "Brand": brand, "Qty": 1, "Unit Price": 2400.0 / ex_rate}
        ]
    elif "DOL" in feeder_type:
        dol_cont_a = math.ceil(flc * 1.15)
        return [
            {"Item": f"DOL Incomer MPCB ({dol_cont_a}A)", "Specification": f"Adjustable MPCB ({brand})", "Brand": brand, "Qty": 1, "Unit Price": 3200.0 / ex_rate},
            {"Item": f"DOL Power Contactor ({dol_cont_a}A AC-3)", "Specification": f"{dol_cont_a}A 3P Contactor ({brand})", "Brand": brand, "Qty": 1, "Unit Price": (1400 + dol_cont_a * 20) / ex_rate}
        ]
    elif "VFD" in feeder_type:
        mccb_a = max(32, math.ceil(flc * 1.25 / 10.0) * 10)
        return [
            {"Item": f"VFD Power Unit ({rating_kw_str})", "Specification": f"{rating_kw_str} 415V Drive ({brand})", "Brand": brand, "Qty": 1, "Unit Price": (22000 + kw_val * 950) / ex_rate},
            {"Item": f"VFD Incomer MCCB ({mccb_a}A)", "Specification": f"{mccb_a}A 3P Breaker ({brand})", "Brand": brand, "Qty": 1, "Unit Price": (3200 + mccb_a * 15) / ex_rate},
            {"Item": "Semiconductor Fast Fuses (aR)", "Specification": "Fast Fuse Set (Set of 3)", "Brand": brand, "Qty": 1, "Unit Price": 5400.0 / ex_rate},
            {"Item": "3% Input AC Line Reactor Choke", "Specification": "Harmonic Choke", "Brand": "Elcon", "Qty": 1, "Unit Price": (2500 + kw_val * 120) / ex_rate},
            {"Item": "Door Keypad BOP & Potentiometer", "Specification": "BOP Keypad + Speed Pot", "Brand": brand, "Qty": 1, "Unit Price": 3200.0 / ex_rate}
        ]
    return []

def auto_layout_panel_components(bom_df, panel_h_mm, bay_w_mm, busbar_pos="Top", cable_alley_w_mm=200):
    busbar_h_mm = 250
    gland_h_mm = 150
    usable_h_mm = panel_h_mm - busbar_h_mm - gland_h_mm
    
    bays = []
    current_bay_index = 0
    current_bay_y = busbar_h_mm if busbar_pos == "Top" else 50
    current_row_max_h = 0
    current_row_x = 30
    
    bays.append({"bay_num": 1, "components": [], "width_mm": bay_w_mm, "height_mm": panel_h_mm})
    
    for _, row in bom_df.iterrows():
        item_name = str(row["Item"])
        qty = int(row["Qty"])
        
        # Match dimensions from standard library or fallback
        dim = None
        for k, v in COMPONENT_DIMENSIONS_LIBRARY.items():
            if k.lower() in item_name.lower():
                dim = v
                break
        if not dim:
            dim = {"Height_mm": 120, "Width_mm": 100, "Depth_mm": 90, "Type": "General"}
            
        comp_h = dim["Height_mm"]
        comp_w = dim["Width_mm"]
        
        for q in range(qty):
            # Check horizontal overflow in current row
            if current_row_x + comp_w + 20 > bay_w_mm - 30:
                # Wrap to next row down
                current_row_x = 30
                current_bay_y += current_row_max_h + 25
                current_row_max_h = 0
                
            # Check vertical overflow in current bay
            max_allowed_y = (panel_h_mm - gland_h_mm) if busbar_pos == "Top" else (panel_h_mm - busbar_h_mm - gland_h_mm)
            if current_bay_y + comp_h > max_allowed_y:
                # Overflow to next panel bay!
                current_bay_index += 1
                current_bay_y = busbar_h_mm if busbar_pos == "Top" else 50
                current_row_x = 30
                current_row_max_h = 0
                bays.append({"bay_num": current_bay_index + 1, "components": [], "width_mm": bay_w_mm, "height_mm": panel_h_mm})
                
            # Place component in current bay
            bays[current_bay_index]["components"].append({
                "name": item_name,
                "x": current_row_x,
                "y": current_bay_y,
                "w": comp_w,
                "h": comp_h,
                "type": dim["Type"]
            })
            
            current_row_x += comp_w + 20
            if comp_h > current_row_max_h:
                current_row_max_h = comp_h
                
    return bays

def generate_internal_ga_svg(bays, panel_h_mm, bay_w_mm, busbar_pos="Top", cable_alley_w=200):
    num_bays = len(bays)
    total_w_mm = (num_bays * bay_w_mm) + cable_alley_w
    scale = 0.28
    svg_w = int(total_w_mm * scale) + 80
    svg_h = int(panel_h_mm * scale) + 80
    
    svg = [f'<svg width="{svg_w}" height="{svg_h}" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#0f172a; border-radius:8px;">']
    svg.append(f'<text x="{svg_w/2}" y="25" fill="#38bdf8" font-family="sans-serif" font-size="14" font-weight="bold" text-anchor="middle">INTERNAL GENERAL ARRANGEMENT (GA) - {num_bays} BAY PANEL ({panel_h_mm}x{total_w_mm}mm)</text>')
    
    start_x = 40
    start_y = 40
    
    # Render Cable Alley
    alley_w_px = cable_alley_w * scale
    panel_h_px = panel_h_mm * scale
    svg.append(f'<rect x="{start_x}" y="{start_y}" width="{alley_w_px}" height="{panel_h_px}" fill="#1e293b" stroke="#64748b" stroke-width="2"/>')
    svg.append(f'<text x="{start_x + alley_w_px/2}" y="{start_y + panel_h_px/2}" fill="#94a3b8" font-family="monospace" font-size="11" text-anchor="middle" transform="rotate(-90,{start_x + alley_w_px/2},{start_y + panel_h_px/2})">CABLE ALLEY ({cable_alley_w}mm)</text>')
    
    cur_x = start_x + alley_w_px
    for b_idx, bay in enumerate(bays):
        bay_w_px = bay_w_mm * scale
        # Main Column Frame
        svg.append(f'<rect x="{cur_x}" y="{start_y}" width="{bay_w_px}" height="{panel_h_px}" fill="#020617" stroke="#38bdf8" stroke-width="2"/>')
        
        # Busbar Chamber
        bb_h_px = 250 * scale
        bb_y = start_y if busbar_pos == "Top" else (start_y + panel_h_px - bb_h_px)
        svg.append(f'<rect x="{cur_x+5}" y="{bb_y}" width="{bay_w_px-10}" height="{bb_h_px}" fill="#1e293b" stroke="#e2e8f0" stroke-width="1.5"/>')
        # Busbars R, Y, B, N
        for i, color in enumerate(["#ef4444", "#eab308", "#3b82f6", "#000000"]):
            bar_y = bb_y + 12 + (i * 12)
            svg.append(f'<rect x="{cur_x+15}" y="{bar_y}" width="{bay_w_px-30}" height="8" fill="{color}" stroke="#cbd5e1" stroke-width="0.5"/>')
        svg.append(f'<text x="{cur_x + bay_w_px/2}" y="{bb_y + bb_h_px/2 + 4}" fill="#f8fafc" font-family="sans-serif" font-size="9" font-weight="bold" text-anchor="middle">MAIN BUSBAR CHAMBER (R-Y-B-N)</text>')
        
        # Components Layout
        for comp in bay["components"]:
            cx = cur_x + (comp["x"] * scale)
            cy = start_y + (comp["y"] * scale)
            cw = comp["w"] * scale
            ch = comp["h"] * scale
            
            c_color = "#2563eb" if comp["type"] == "VFD" else ("#059669" if comp["type"] == "MCCB" else ("#d97706" if comp["type"] == "Contactor" else "#475569"))
            svg.append(f'<rect x="{cx}" y="{cy}" width="{cw}" height="{ch}" rx="3" fill="{c_color}" stroke="#e2e8f0" stroke-width="1.5"/>')
            svg.append(f'<text x="{cx + cw/2}" y="{cy + ch/2 + 3}" fill="#ffffff" font-family="sans-serif" font-size="8" font-weight="bold" text-anchor="middle">{comp["name"][:16]}</text>')
            
        # Column Title
        svg.append(f'<text x="{cur_x + bay_w_px/2}" y="{start_y + panel_h_px - 15}" fill="#38bdf8" font-family="sans-serif" font-size="10" font-weight="bold" text-anchor="middle">BAY #{b_idx+1} ({bay_w_mm}mm)</text>')
        cur_x += bay_w_px

    svg.append('</svg>')
    return "".join(svg)

def generate_outer_ga_svg(bays, panel_h_mm, bay_w_mm, brand="Siemens", panel_type="MCC Panel"):
    num_bays = len(bays)
    total_w_mm = (num_bays * bay_w_mm) + 200
    scale = 0.28
    svg_w = int(total_w_mm * scale) + 80
    svg_h = int(panel_h_mm * scale) + 80
    
    svg = [f'<svg width="{svg_w}" height="{svg_h}" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#0d1117; border-radius:8px;">']
    svg.append(f'<text x="{svg_w/2}" y="25" fill="#38bdf8" font-family="sans-serif" font-size="14" font-weight="bold" text-anchor="middle">OUTER ELEVATION GA VIEW - {panel_type} ({num_bays} BAY)</text>')
    
    start_x = 40
    start_y = 40
    alley_w_px = 200 * scale
    panel_h_px = panel_h_mm * scale
    
    # Cable Alley Door
    svg.append(f'<rect x="{start_x}" y="{start_y}" width="{alley_w_px}" height="{panel_h_px}" fill="#1e293b" stroke="#64748b" stroke-width="2"/>')
    svg.append(f'<circle cx="{start_x + alley_w_px - 15}" cy="{start_y + panel_h_px/2}" r="5" fill="#475569" stroke="#cbd5e1" stroke-width="1.5"/>')
    
    cur_x = start_x + alley_w_px
    for b_idx in range(num_bays):
        bay_w_px = bay_w_mm * scale
        # Main Front Door Outer Shell
        svg.append(f'<rect x="{cur_x}" y="{start_y}" width="{bay_w_px}" height="{panel_h_px}" fill="#0f172a" stroke="#38bdf8" stroke-width="2.5"/>')
        # Door Hinges
        svg.append(f'<rect x="{cur_x + 5}" y="{start_y + 60}" width="6" height="20" fill="#64748b"/>')
        svg.append(f'<rect x="{cur_x + 5}" y="{start_y + panel_h_px - 80}" width="6" height="20" fill="#64748b"/>')
        # Concealed Door Lock Handle
        svg.append(f'<rect x="{cur_x + bay_w_px - 20}" y="{start_y + panel_h_px/2 - 15}" width="12" height="30" rx="3" fill="#334155" stroke="#94a3b8" stroke-width="1.5"/>')
        
        # Metering & Indicator Panel Header
        svg.append(f'<rect x="{cur_x + 20}" y="{start_y + 20}" width="{bay_w_px - 40}" height="70" rx="4" fill="#020617" stroke="#0ea5e9" stroke-width="1.5"/>')
        # Digital MFM Display
        svg.append(f'<rect x="{cur_x + 30}" y="{start_y + 30}" width="60" height="50" rx="2" fill="#0f172a" stroke="#38bdf8" stroke-width="1"/>')
        svg.append(f'<text x="{cur_x + 60}" y="{start_y + 58}" fill="#38bdf8" font-family="monospace" font-size="11" font-weight="bold" text-anchor="middle">415.2 V</text>')
        # Phase Indicator Lamps R, Y, B
        for idx, color in enumerate(["#ef4444", "#eab308", "#3b82f6"]):
            svg.append(f'<circle cx="{cur_x + bay_w_px - 40}" cy="{start_y + 35 + (idx*18)}" r="6" fill="{color}"/>')
            
        # VFD Keypad / Pushbuttons Control Cluster
        svg.append(f'<rect x="{cur_x + 20}" y="{start_y + 110}" width="{bay_w_px - 40}" height="100" rx="4" fill="#1e293b" stroke="#64748b" stroke-width="1"/>')
        svg.append(f'<circle cx="{cur_x + 50}" cy="{start_y + 140}" r="10" fill="#22c55e"/>') # Green Start
        svg.append(f'<circle cx="{cur_x + 90}" cy="{start_y + 140}" r="10" fill="#ef4444"/>') # Red Stop
        svg.append(f'<circle cx="{cur_x + 130}" cy="{start_y + 140}" r="10" fill="#eab308"/>') # Yellow Trip
        
        # Bottom Ventilation Louver Filter Fan
        svg.append(f'<rect x="{cur_x + 25}" y="{start_y + panel_h_px - 100}" width="80" height="60" rx="3" fill="#1e293b" stroke="#94a3b8" stroke-width="1.5"/>')
        for l_line in range(4):
            svg.append(f'<line x1="{cur_x + 35}" y1="{start_y + panel_h_px - 90 + (l_line*12)}" x2="{cur_x + 95}" y2="{start_y + panel_h_px - 90 + (l_line*12)}" stroke="#64748b" stroke-width="2"/>')
            
        # Rating Nameplate
        svg.append(f'<rect x="{cur_x + bay_w_px/2 - 70}" y="{start_y + panel_h_px - 32}" width="140" height="22" rx="2" fill="#f8fafc" stroke="#475569"/>')
        svg.append(f'<text x="{cur_x + bay_w_px/2}" y="{start_y + panel_h_px - 18}" fill="#0f172a" font-family="sans-serif" font-size="8" font-weight="bold" text-anchor="middle">ARYAVARTA ({brand})</text>')
        cur_x += bay_w_px

    svg.append('</svg>')
    return "".join(svg)

def generate_detailed_sld_svg(incomer_info="630A 3P MCCB", feeder_items=None, brand="Siemens"):
    svg_w = 540
    svg_h = 320
    svg = [f'<svg width="{svg_w}" height="{svg_h}" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#0f172a; border-radius:8px;">']
    svg.append(f'<text x="{svg_w/2}" y="25" fill="#38bdf8" font-family="sans-serif" font-size="13" font-weight="bold" text-anchor="middle">POWER SINGLE LINE DIAGRAM (SLD)</text>')
    
    # Main Busbar
    svg.append('<line x1="40" y1="50" x2="500" y2="50" stroke="#ef4444" stroke-width="3"/>')
    svg.append('<line x1="40" y1="56" x2="500" y2="56" stroke="#eab308" stroke-width="3"/>')
    svg.append('<line x1="40" y1="62" x2="500" y2="62" stroke="#3b82f6" stroke-width="3"/>')
    svg.append('<text x="270" y="42" fill="#94a3b8" font-family="monospace" font-size="10" text-anchor="middle">415V 3-PHASE 50Hz MAIN BUSBAR</text>')
    
    # Incomer Breaker
    svg.append('<line x1="270" y1="62" x2="270" y2="95" stroke="#cbd5e1" stroke-width="2"/>')
    svg.append('<rect x="220" y="95" width="100" height="40" rx="4" fill="#1e293b" stroke="#38bdf8" stroke-width="2"/>')
    svg.append(f'<text x="270" y="118" fill="#38bdf8" font-family="sans-serif" font-size="10" font-weight="bold" text-anchor="middle">{incomer_info[:15]}</text>')
    
    # CT Ring
    svg.append('<line x1="270" y1="135" x2="270" y2="175" stroke="#cbd5e1" stroke-width="2"/>')
    svg.append('<circle cx="270" cy="155" r="12" fill="none" stroke="#eab308" stroke-width="2"/>')
    svg.append('<text x="295" y="158" fill="#eab308" font-family="sans-serif" font-size="9">CT 600/5A</text>')
    
    # Distribution Feeders
    feeder_x_positions = [120, 270, 420]
    f_titles = ["VFD Drive Feeder", "Star-Delta Starter", "Auxiliary DOL Feeder"]
    
    svg.append('<line x1="120" y1="175" x2="420" y2="175" stroke="#cbd5e1" stroke-width="2"/>')
    
    for idx, fx in enumerate(feeder_x_positions):
        svg.append(f'<line x1="{fx}" y1="175" x2="{fx}" y2="205" stroke="#cbd5e1" stroke-width="2"/>')
        svg.append(f'<rect x="{fx-45}" y="205" width="90" height="38" rx="3" fill="#020617" stroke="#22c55e" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="228" fill="#22c55e" font-family="sans-serif" font-size="9" font-weight="bold" text-anchor="middle">{f_titles[idx]}</text>')
        svg.append(f'<line x1="{fx}" y1="243" x2="{fx}" y2="265" stroke="#cbd5e1" stroke-width="2"/>')
        svg.append(f'<circle cx="{fx}" cy="280" r="14" fill="#1e293b" stroke="#e2e8f0" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="284" fill="#ffffff" font-family="sans-serif" font-size="10" font-weight="bold" text-anchor="middle">M 3~</text>')
        
    svg.append('</svg>')
    return "".join(svg)

def generate_excel_proposal(client_name, panel_type, motor_kw, preferred_brand, curr_opt, sell_price, curr_sym, bom_df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        summary_df = pd.DataFrame([
            {"Field": "Client Name", "Value": str(client_name)},
            {"Field": "Quote Ref", "Value": f"AA/QT/{os.urandom(2).hex().upper()}"},
            {"Field": "Scope of Supply", "Value": f"{panel_type} ({motor_kw})"},
            {"Field": "Switchgear Brand", "Value": str(preferred_brand)},
            {"Field": "Currency", "Value": str(curr_opt)},
            {"Field": f"Total Basic Price ({curr_sym})", "Value": f"{curr_sym} {sell_price:,.2f}"},
            {"Field": "GST / Import Taxes", "Value": "18% Extra as applicable"},
            {"Field": "Delivery Schedule", "Value": "2-3 Weeks Ex-Factory Chikhali, Pune"},
            {"Field": "Payment Terms", "Value": "50% Advance, 50% Before Dispatch"}
        ])
        summary_df.to_excel(w, index=False, sheet_name="Commercial Offer")
        
        export_bom = bom_df.copy()
        export_bom.to_excel(w, index=False, sheet_name="Bill of Materials")
        
        tds_df = pd.DataFrame([
            {"Parameter": "Operating Voltage", "Specification": "415V AC ± 10%, 3Ø 50Hz"},
            {"Parameter": "Design Standard", "Specification": "IEC 61439-1 / IS 8623"},
            {"Parameter": "Enclosure Rating", "Specification": "IP54 / IP55 Sheet Metal"},
            {"Parameter": "Switchgear Brand", "Specification": str(preferred_brand)},
            {"Parameter": "Short Circuit Fault Level", "Specification": "25kA for 1s"}
        ])
        tds_df.to_excel(w, index=False, sheet_name="Technical Specs")
    buf.seek(0)
    return buf.getvalue()

def generate_pdf_quotation(client_name, panel_type, motor_kw, brand, bom_df, labor, margin, total_price, curr_sym="₹"):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    style_th_l = ParagraphStyle('THL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    style_th_r = ParagraphStyle('THR', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=2)
    style_th_c = ParagraphStyle('THC', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    
    style_td_l = ParagraphStyle('TDL', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=0)
    style_td_r = ParagraphStyle('TDR', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=2)
    style_td_c = ParagraphStyle('TDC', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=1)
    style_bold_r = ParagraphStyle('TDBR', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, alignment=2)
    style_bold_l = ParagraphStyle('TDBL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, alignment=0)

    story = [
        Paragraph("ARYAVARTA AUTOMATION", ParagraphStyle('H1', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"))),
        Paragraph("<b>Manufacturer & Exporter of Switchgear Panels, VFD Panels & Automation Systems</b>", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=9, leading=12)),
        Paragraph("Gat No. 1610, Dehu Alandi Road, Chikhali, Pune - 411062 | GSTIN: 27ABOFA4930E1ZH", ParagraphStyle('Sub2', parent=styles['Normal'], fontSize=8, leading=10)),
        Spacer(1, 8)
    ]
    meta = [
        [Paragraph(f"<b>Client:</b> {client_name}", style_td_l), Paragraph(f"<b>Quote Ref:</b> AA/QT/{os.urandom(2).hex().upper()}", style_td_l)],
        [Paragraph(f"<b>Scope:</b> {panel_type} ({motor_kw})", style_td_l), Paragraph(f"<b>Date:</b> {pd.Timestamp.now().strftime('%d-%b-%Y')}", style_td_l)]
    ]
    t_meta = Table(meta, colWidths=[270, 270])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.extend([t_meta, Spacer(1, 10), Paragraph("<b>Bill of Materials & Technical Specifications</b>", ParagraphStyle('H3', parent=styles['Heading3'], fontSize=11, leading=14))])
    
    table_data = [[
        Paragraph("Item Description & Specification", style_th_l),
        Paragraph("Brand", style_th_c),
        Paragraph("Qty", style_th_c),
        Paragraph(f"Unit ({curr_sym})", style_th_r),
        Paragraph(f"Total ({curr_sym})", style_th_r)
    ]]
    
    for idx, r in bom_df.iterrows():
        table_data.append([
            Paragraph(f"<b>{r['Item']}</b><br/><font color='#4A5568'>{r['Specification']}</font>", style_td_l),
            Paragraph(str(r['Brand']), style_td_c),
            Paragraph(str(r['Qty']), style_td_c),
            Paragraph(f"{curr_sym} {r['Unit Price']:,.2f}", style_td_r),
            Paragraph(f"{curr_sym} {r['Total Material Cost']:,.2f}", style_td_r)
        ])
    
    table_data.append([
        Paragraph("<b>Assembly & Wiring Labor</b>", style_td_l),
        Paragraph("Aryavarta", style_td_c),
        Paragraph("1", style_td_c),
        Paragraph(f"{curr_sym} {labor:,.2f}", style_td_r),
        Paragraph(f"{curr_sym} {labor:,.2f}", style_td_r)
    ])
    table_data.append([
        Paragraph("<b>TOTAL PRICE (Excl. Tax)</b>", style_bold_l),
        Paragraph("", style_td_c),
        Paragraph("", style_td_c),
        Paragraph("", style_td_r),
        Paragraph(f"<b>{curr_sym} {total_price:,.2f}</b>", style_bold_r)
    ])
    
    t_bom = Table(table_data, colWidths=[230, 70, 35, 100, 105])
    t_bom.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('SPAN', (0, -1), (2, -1)),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#EDF2F7"))
    ]))
    story.extend([t_bom, Spacer(1, 12), Paragraph("<b>Terms:</b> 1. Taxes/Duties extra as applicable. 2. Delivery 2-3 weeks. 3. Payment 50% advance, 50% dispatch.", ParagraphStyle('Terms', parent=styles['Normal'], fontSize=8, leading=11)), Spacer(1, 15), Paragraph("<b>For ARYAVARTA AUTOMATION</b><br/><br/>Authorized Signatory", ParagraphStyle('Sig', parent=styles['Normal'], fontSize=9, leading=12))])
    doc.build(story)
    buf.seek(0)
    return buf

def generate_fat_certificate_pdf(serial, project, client, inspector, r, y, b, hv_pass):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    style_th_l = ParagraphStyle('THL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    style_th_c = ParagraphStyle('THC', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    style_td_l = ParagraphStyle('TDL', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=0)
    style_td_c = ParagraphStyle('TDC', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=1)

    story = [
        Paragraph("ARYAVARTA AUTOMATION", ParagraphStyle('H1', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"))),
        Paragraph("<b>FACTORY ACCEPTANCE TEST (FAT) & QUALITY CERTIFICATE</b>", ParagraphStyle('H2', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor("#2B6CB0"))),
        Paragraph(f"Serial: {serial} | Client: {client} | Date: {pd.Timestamp.now().strftime('%d-%b-%Y')}", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=8, leading=10)),
        Spacer(1, 10), Paragraph("<b>1. Insulation Resistance Test (500V Megger)</b>", ParagraphStyle('H3', parent=styles['Heading3'], fontSize=10, leading=12))
    ]
    
    t_m_data = [
        [Paragraph("Phase", style_th_c), Paragraph("Measured Value", style_th_c), Paragraph("Min Required", style_th_c), Paragraph("Result", style_th_c)],
        [Paragraph("R-E", style_td_c), Paragraph(f"{r:.1f} MΩ", style_td_c), Paragraph("> 2.0 MΩ", style_td_c), Paragraph("PASS" if r>=2 else "FAIL", style_td_c)],
        [Paragraph("Y-E", style_td_c), Paragraph(f"{y:.1f} MΩ", style_td_c), Paragraph("> 2.0 MΩ", style_td_c), Paragraph("PASS" if y>=2 else "FAIL", style_td_c)],
        [Paragraph("B-E", style_td_c), Paragraph(f"{b:.1f} MΩ", style_td_c), Paragraph("> 2.0 MΩ", style_td_c), Paragraph("PASS" if b>=2 else "FAIL", style_td_c)]
    ]
    t_m = Table(t_m_data, colWidths=[135, 135, 135, 135])
    t_m.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    
    story.extend([t_m, Spacer(1, 10), Paragraph("<b>2. HV & Safety Checks</b>", ParagraphStyle('H3', parent=styles['Heading3'], fontSize=10, leading=12))])
    
    t_c_data = [
        [Paragraph("Check Item", style_th_l), Paragraph("Standard", style_th_c), Paragraph("Status", style_th_c)],
        [Paragraph("2.5kV HV Withstand (1 Min)", style_td_l), Paragraph("IEC 61439-1", style_td_c), Paragraph("PASSED" if hv_pass else "FAILED", style_td_c)],
        [Paragraph("Busbar Phase Clearance", style_td_l), Paragraph("IS 8623", style_td_c), Paragraph("PASSED", style_td_c)],
        [Paragraph("Door Earth Continuity", style_td_l), Paragraph("Safety Standard", style_td_c), Paragraph("PASSED", style_td_c)]
    ]
    t_c = Table(t_c_data, colWidths=[240, 180, 120])
    t_c.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#4A5568")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.extend([t_c, Spacer(1, 15), Paragraph(f"QA Inspector: {inspector} &nbsp;&nbsp;&nbsp;&nbsp; Quality Manager Stamp: ____________", ParagraphStyle('Sig', parent=styles['Normal'], fontSize=9, leading=12))])
    doc.build(story)
    buf.seek(0)
    return buf

def generate_proforma_invoice_pdf(client_name, panel_type, motor_kw, total_price, curr_sym="₹"):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    tax_amt = total_price * 0.18
    grand_tot = total_price + tax_amt
    
    style_th_l = ParagraphStyle('THL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    style_th_r = ParagraphStyle('THR', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=2)
    style_th_c = ParagraphStyle('THC', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    
    style_td_l = ParagraphStyle('TDL', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=0)
    style_td_r = ParagraphStyle('TDR', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=2)
    style_td_c = ParagraphStyle('TDC', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=1)

    story = [
        Paragraph("ARYAVARTA AUTOMATION", ParagraphStyle('H1', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"))),
        Paragraph("<b>PROFORMA INVOICE</b> (HSN Code: 85371010 - Industrial Control Panels)", ParagraphStyle('H2', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor("#2B6CB0"))),
        Paragraph("Gat No. 1610, Dehu Alandi Road, Chikhali, Pune - 411062 | GSTIN: 27ABOFA4930E1ZH", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=8, leading=10)),
        Spacer(1, 8)
    ]
    meta = [
        [Paragraph(f"<b>Billed To:</b> {client_name}", style_td_l), Paragraph(f"<b>PI No:</b> AA/PI/{os.urandom(2).hex().upper()}", style_td_l)],
        [Paragraph(f"<b>Item Description:</b> {panel_type} ({motor_kw})", style_td_l), Paragraph(f"<b>Date:</b> {pd.Timestamp.now().strftime('%d-%b-%Y')}", style_td_l)]
    ]
    t_meta = Table(meta, colWidths=[270, 270])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    
    t_inv_data = [
        [
            Paragraph("Particulars", style_th_l),
            Paragraph("HSN Code", style_th_c),
            Paragraph(f"Basic ({curr_sym})", style_th_r),
            Paragraph("Tax Rate", style_th_c),
            Paragraph(f"Tax ({curr_sym})", style_th_r),
            Paragraph(f"Total ({curr_sym})", style_th_r)
        ],
        [
            Paragraph(f"<b>{panel_type} ({motor_kw}) Control Panel</b>", style_td_l),
            Paragraph("85371010", style_td_c),
            Paragraph(f"{curr_sym} {total_price:,.2f}", style_td_r),
            Paragraph("18%", style_td_c),
            Paragraph(f"{curr_sym} {tax_amt:,.2f}", style_td_r),
            Paragraph(f"{curr_sym} {grand_tot:,.2f}", style_td_r)
        ]
    ]
    t_inv = Table(t_inv_data, colWidths=[175, 65, 85, 45, 85, 85])
    t_inv.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'TOP')
    ]))
    
    bank = [
        [Paragraph("<b>Bank Payment Details for Wire Transfer / RTGS / Swift:</b>", style_td_l)],
        [Paragraph("Bank Name: HDFC Bank Ltd | Branch: Chinchwad, Pune", style_td_l)],
        [Paragraph("Account Name: ARYAVARTA AUTOMATION | Account No: 50200084920193 | IFSC: HDFC0000148 | SWIFT: HDFCINBB", style_td_l)]
    ]
    t_bank = Table(bank, colWidths=[540])
    t_bank.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EDF2F7")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 4)
    ]))
    story.extend([t_meta, Spacer(1, 10), t_inv, Spacer(1, 12), t_bank, Spacer(1, 15), Paragraph("<b>For ARYAVARTA AUTOMATION</b><br/><br/>Authorized Signatory", ParagraphStyle('Sig', parent=styles['Normal'], fontSize=9, leading=12))])
    doc.build(story)
    buf.seek(0)
    return buf

def generate_delivery_challan_pdf(dc_no, client_name, address, vehicle_no, panel_type, gross_weight):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    style_th_l = ParagraphStyle('THL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    style_th_c = ParagraphStyle('THC', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    style_td_l = ParagraphStyle('TDL', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=0)
    style_td_c = ParagraphStyle('TDC', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=1)

    story = [
        Paragraph("ARYAVARTA AUTOMATION", ParagraphStyle('H1', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"))),
        Paragraph("<b>DELIVERY CHALLAN & DISPATCH NOTE</b> (For Goods Transport / Gate Pass)", ParagraphStyle('H2', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor("#2B6CB0"))),
        Paragraph("Gat No. 1610, Dehu Alandi Road, Chikhali, Pune - 411062 | GSTIN: 27ABOFA4930E1ZH", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=8, leading=10)),
        Spacer(1, 8)
    ]
    meta = [
        [Paragraph(f"<b>Challan No:</b> {dc_no}", style_td_l), Paragraph(f"<b>Date:</b> {pd.Timestamp.now().strftime('%d-%b-%Y')}", style_td_l)],
        [Paragraph(f"<b>Consignee / Client:</b> {client_name}", style_td_l), Paragraph(f"<b>Vehicle / LR No:</b> {vehicle_no}", style_td_l)],
        [Paragraph(f"<b>Delivery Address:</b> {address}", style_td_l), Paragraph("<b>HSN Code:</b> 85371010", style_td_l)]
    ]
    t_meta = Table(meta, colWidths=[270, 270])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    
    t_items_data = [
        [Paragraph("S.No", style_th_c), Paragraph("Item Description", style_th_l), Paragraph("HSN Code", style_th_c), Paragraph("Qty", style_th_c), Paragraph("Gross Wt (kg)", style_th_c), Paragraph("Status", style_th_c)],
        [Paragraph("1", style_td_c), Paragraph(f"<b>{panel_type} Industrial Control Panel</b>", style_td_l), Paragraph("85371010", style_td_c), Paragraph("1 Set", style_td_c), Paragraph(f"{gross_weight} kg", style_td_c), Paragraph("Supply Only", style_td_c)]
    ]
    t_items = Table(t_items_data, colWidths=[35, 215, 70, 50, 85, 85])
    t_items.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.extend([t_meta, Spacer(1, 10), t_items, Spacer(1, 15), Paragraph("<b>Declaration:</b> Material dispatched in good condition for installation & commissioning. Not for sale.", ParagraphStyle('Dec', parent=styles['Normal'], fontSize=8, leading=10)), Spacer(1, 20), Paragraph("<b>Receiver's Stamp & Signature:</b> ____________________ &nbsp;&nbsp;&nbsp;&nbsp; <b>For ARYAVARTA AUTOMATION</b>", ParagraphStyle('Sig', parent=styles['Normal'], fontSize=9, leading=12))])
    doc.build(story)
    buf.seek(0)
    return buf

def generate_tds_pdf(client_name, panel_type, motor_kw, brand):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    style_th_l = ParagraphStyle('THL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    style_td_l = ParagraphStyle('TDL', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=0)
    style_td_b = ParagraphStyle('TDB', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, alignment=0)

    story = [
        Paragraph("ARYAVARTA AUTOMATION", ParagraphStyle('H1', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"))),
        Paragraph("<b>TECHNICAL DATA SHEET (TDS) & TENDER COMPLIANCE</b>", ParagraphStyle('H2', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor("#2B6CB0"))),
        Paragraph("Gat No. 1610, Dehu Alandi Road, Chikhali, Pune - 411062 | GSTIN: 27ABOFA4930E1ZH", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=8, leading=10)),
        Spacer(1, 10)
    ]
    raw_data = [
        ["Parameter", "Specification / Technical Compliance"],
        ["Equipment Scope", f"{panel_type} ({motor_kw})"],
        ["Target Client", client_name],
        ["Operating Voltage & Freq.", "415V AC ± 10%, 3-Phase 4-Wire, 50 Hz ± 5%"],
        ["Design Standard Compliance", "IEC 61439-1 / IS 8623 / IS 13947"],
        ["Switchgear Component Brand", f"{brand} (Form-2B Compartmentalized)"],
        ["Enclosure Protection Rating", "IP54 / IP55 Powder Coated Floor / Wall Mount"],
        ["Sheet Metal Construction", "1.6mm / 2.0mm CRCA Sheet Steel"],
        ["Surface Treatment & Shade", "7-Tank Pre-treatment & Polyester Powder Coating (RAL 7035)"],
        ["Main Busbar Conductor", "EC Grade High-Conductivity Copper (1.5 A/mm²)"],
        ["Short Circuit Fault Level", "25 kA for 1 Second (IS 8623)"],
        ["Control Circuit Voltage", "24V DC SMPS / 230V AC Isolated Control Tx"]
    ]
    
    table_data = [[Paragraph(f"<b>{raw_data[0][0]}</b>", style_th_l), Paragraph(f"<b>{raw_data[0][1]}</b>", style_th_l)]]
    for row in raw_data[1:]:
        table_data.append([Paragraph(f"<b>{row[0]}</b>", style_td_b), Paragraph(row[1], style_td_l)])
        
    t = Table(table_data, colWidths=[170, 370])
    t.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.extend([t, Spacer(1, 15), Paragraph("<b>For ARYAVARTA AUTOMATION</b><br/><br/>Authorized Technical Signatory", ParagraphStyle('Sig', parent=styles['Normal'], fontSize=9, leading=12))])
    doc.build(story)
    buf.seek(0)
    return buf

st.sidebar.title("⚡ Aryavarta Automation")
st.sidebar.markdown("**Sales & Engineering Suite**")
st.sidebar.success("🟢 100% Offline & Free | Chikhali Plant")

menu = st.sidebar.radio("Navigation", [
    "Create Panel Quote", 
    "MCC Multi-Feeder Panel Sizer",
    "Motor Switchgear Master Chart",
    "MPCB & Motor Protection Sizer",
    "ACB Selection & Protection Release Sizer",
    "Motor Starting Voltage Dip Calculator",
    "Quotation Register & Pipeline",
    "PLC & HMI Automation Sizer",
    "Star-Delta & DOL Switchgear Sizing",
    "Soft Starter & Bypass Contactor Sizer",
    "Enclosure Fabrication Costing", 
    "VFD Energy Savings & Payback", 
    "VFD Dynamic Braking Resistor (DBR) Sizer",
    "Cable & Busbar Calculator", 
    "Cable Gland & Terminal Sizer",
    "Control Wire Color & Ferrule Guide",
    "IP Protection Rating & Gasket Guide",
    "Short-Circuit kA & Copper Weight", 
    "Transformer & DG Sizing", 
    "APFC Capacitor Sizing & MSEDCL ROI", 
    "Harmonics & AHF Sizing",
    "Panel Thermal & Fan Sizing", 
    "Control Tx & 24V SMPS Sizing",
    "Control Panel UPS & Battery Sizer",
    "CT Ratio, Class & Burden Sizer",
    "Earthing Pit Resistance (IS 3043)",
    "Voltage Unbalance & NEMA Derating (NEMA MG-1)",
    "Neutral Busbar & Triplen Harmonics (IEC 60364)",
    "FAT Quality Certificate", 
    "Dispatch Delivery Challan (DC)",
    "VFD Diagnostic Assistant", 
    "Manage Price Database"
])

if menu == "Create Panel Quote":
    st.header("📋 Panel Estimator, PDF & WhatsApp Proposal")
    with st.expander("🤖 Auto-Extract Specs from Customer Text or Upload Documents (PDF, Excel, Word, TXT, CSV)", expanded=True):
        up_col1, up_col2 = st.columns([1, 1])
        with up_col1:
            uploaded_inquiry_files = st.file_uploader("📁 Upload Inquiry Files - Multi-File Supported", type=["pdf", "xlsx", "xls", "csv", "txt", "docx"], accept_multiple_files=True)
            if uploaded_inquiry_files:
                combined_texts = []
                for uf in uploaded_inquiry_files:
                    txt = extract_text_from_file(uf)
                    combined_texts.append(f"--- File: {uf.name} ---\n{txt}")
                extracted_file_text = "\n\n".join(combined_texts)
                st.session_state["raw_inquiry_text"] = extracted_file_text
                st.success(f"Loaded {len(uploaded_inquiry_files)} file(s) successfully!")

        with up_col2:
            raw_rfq = st.text_area("Paste or Review customer inquiry text:", value=st.session_state.get("raw_inquiry_text", ""), placeholder="e.g. Need complete MCC panel with 2x 22kW VFD, 3x 37kW Star-Delta, 4x 7.5kW DOL feeders with Siemens switchgear.", height=120)

        if st.button("⚡ Process Inquiry with AI / Smart Extractor"):
            if raw_rfq:
                parsed = query_local_ollama(f"Extract specs from: '{raw_rfq}'. Keys: 'motor_kw', 'panel_type', 'preferred_brand'")
                if not parsed:
                    parsed = parse_inquiry_text_heuristically(raw_rfq)
                
                if parsed:
                    if parsed.get("panel_type"): st.session_state["sel_panel"] = parsed["panel_type"]
                    if parsed.get("motor_kw"): st.session_state["sel_kw"] = parsed["motor_kw"]
                    if parsed.get("preferred_brand"): st.session_state["sel_brand"] = parsed["brand"] if "brand" in parsed else parsed.get("preferred_brand")
                    st.success(f"Parsed Specs: Panel: {parsed.get('panel_type', 'N/A')} | Motor: {parsed.get('motor_kw', 'N/A')} | Brand: {parsed.get('preferred_brand', 'N/A')}")
                    st.rerun()

    quote_mode = st.radio("Estimation Mode", ["⚡ Single Feeder Panel", "🏢 Full Multi-Feeder Panel / MCC / PCC Board"], horizontal=True)

    if quote_mode == "⚡ Single Feeder Panel":
        c1, c2, c3, c4 = st.columns([3, 3, 3, 3])
        p_types = ["VFD Panel", "Star-Delta Control Panel", "APFC Panel", "LT Distribution Panel"]
        m_kw = ["7.5 kW", "15 kW", "22 kW", "37 kW", "55 kW", "75 kW"]
        b_list = ["Siemens", "Schneider", "L&T", "ABB", "Danfoss", "Delta"]
        inc_ratings = ["Auto-Size by kW", "100A 3P MCCB (25kA)", "250A 3P MCCB (36kA)", "400A 3P MCCB (36kA)", "630A 3P MCCB (50kA)"]
        
        with c1:
            panel_type = st.selectbox("Panel Type", p_types, index=p_types.index(st.session_state.get("sel_panel", "VFD Panel")) if st.session_state.get("sel_panel") in p_types else 0)
            client_name = st.text_input("Client Name", "Maharashtra Water Works Ltd")
        with c2:
            motor_kw = st.selectbox("Motor Rating", m_kw, index=m_kw.index(st.session_state.get("sel_kw", "15 kW")) if st.session_state.get("sel_kw") in m_kw else 1)
            preferred_brand = st.selectbox("Switchgear Brand", b_list, index=b_list.index(st.session_state.get("sel_brand", "Siemens")) if st.session_state.get("sel_brand") in b_list else 0)
        with c3:
            selected_incomer = st.selectbox("Main Incomer Specification", inc_ratings, index=0)
            margin_pct = st.slider("Margin (%)", 5, 40, 18)
        with c4:
            labor_cost_inr = st.number_input("Wiring & Assembly Labor (₹)", value=4500, step=500)
            curr_opt = st.selectbox("Currency", ["INR (₹)", "USD ($)", "EUR (€)", "AED (Dh)"])
            ex_rate = st.number_input("Exchange Rate (1 FX = X INR)", value=83.5 if "USD" in curr_opt else (91.0 if "EUR" in curr_opt else (22.7 if "AED" in curr_opt else 1.0)), step=0.1)

        curr_sym = "₹" if "INR" in curr_opt else ("$" if "USD" in curr_opt else ("€" if "EUR" in curr_opt else "AED "))
        labor_cost = labor_cost_inr / ex_rate

        st.subheader("🔌 Optional Panel Accessories & Add-Ons")
        a1, a2, a3, a4 = st.columns(4)
        add_choke = a1.checkbox("3% Input AC Line Choke", value=True if "VFD" in panel_type else False)
        add_heater = a2.checkbox("Space Heater + Thermostat", value=True)
        add_mfm = a3.checkbox("Digital MFM Meter", value=True)
        add_tower = a4.checkbox("Signal Tower Indicator", value=False)

        f_opt = st.selectbox("Freight & Transit Packing Logistics", [
            "Ex-Factory (No Freight Included)",
            "Local Pune / PCMC Delivery & Bubble Wrap",
            "Interstate Transport & Sea-Worthy Wooden Crate"
        ])
        freight_price_inr = 0.0 if "Ex-Factory" in f_opt else (1500.0 if "Local" in f_opt else 6500.0)
        freight_price = freight_price_inr / ex_rate

        if "custom_bom_items" not in st.session_state:
            st.session_state["custom_bom_items"] = []

        with st.expander("➕ Add Custom / Non-Standard Component to BOM", expanded=False):
            c_i1, c_i2, c_i3, c_i4, c_i5 = st.columns([3, 3, 2, 2, 2])
            ci_name = c_i1.text_input("Item Name", "Phase Preventer Relay")
            ci_spec = c_i2.text_input("Specification", "415V 3-Phase Protection")
            ci_brand = c_i3.text_input("Brand", "Minilec")
            ci_qty = c_i4.number_input("Qty", value=1, min_value=1)
            ci_price_inr = c_i5.number_input("Unit Price (₹)", value=1800.0, step=100.0)
            
            ca1, ca2 = st.columns([2, 8])
            if ca1.button("➕ Add Item"):
                st.session_state["custom_bom_items"].append({
                    "Item": ci_name, "Specification": ci_spec, "Brand": ci_brand,
                    "Qty": int(ci_qty), "Unit Price": float(ci_price_inr / ex_rate)
                })
                st.rerun()
            if ca2.button("🗑️ Clear Custom Items") and st.session_state["custom_bom_items"]:
                st.session_state["custom_bom_items"] = []
                st.rerun()

        def build_bom(brand_name):
            bom = []
            enc_r = df_prices[df_prices["Category"]=="Enclosure"].iloc[0] if not df_prices[df_prices["Category"]=="Enclosure"].empty else {}
            bom.append({"Item": "Control Panel Enclosure", "Specification": enc_r.get("Specification","IP54 Floor Mount"), "Brand": enc_r.get("Brand","Standard"), "Qty": 1, "Unit Price": enc_r.get("Unit_Price_INR",14000) / ex_rate})
            
            if "Auto" in selected_incomer:
                kw_val = float(motor_kw.split()[0]) if motor_kw and motor_kw.split()[0].replace('.','',1).isdigit() else 15.0
                flc = (kw_val * 1000) / (1.732 * 415 * 0.85 * 0.88)
                inc_target = "630A" if flc > 250 else ("400A" if flc > 125 else ("250A" if flc > 63 else "100A"))
            else:
                inc_target = selected_incomer

            incomer_kit = get_incomer_default_components(inc_target, brand_name, ex_rate)
            bom.extend(incomer_kit)
            
            feeder_subcomponents = get_feeder_subcomponents(panel_type, motor_kw, brand_name, ex_rate)
            bom.extend(feeder_subcomponents)

            if add_choke and not any("Line Reactor" in item["Item"] for item in bom): 
                bom.append({"Item": "3% AC Line Reactor Choke", "Specification": f"{motor_kw} Harmonic Filter", "Brand": "Elcon", "Qty": 1, "Unit Price": 3800 / ex_rate})
            if add_heater: 
                bom.append({"Item": "Panel Anti-Condensation Heater", "Specification": "80W + Thermostat", "Brand": "Generic", "Qty": 1, "Unit Price": 1400 / ex_rate})
            if add_mfm and not any("Metering CT" in item["Item"] for item in bom): 
                bom.append({"Item": "Digital Multifunction Meter (MFM)", "Specification": "3-Phase V/A/kW/PF", "Brand": "Rishabh", "Qty": 1, "Unit Price": 2800 / ex_rate})
            if add_tower: 
                bom.append({"Item": "3-Color Signal LED Tower Lamp", "Specification": "24VDC / 230VAC", "Brand": "Generic", "Qty": 1, "Unit Price": 1200 / ex_rate})

            bb_r = df_prices[df_prices["Category"]=="Busbar & Wire"].iloc[0]
            acc_r = df_prices[df_prices["Category"]=="Accessories"].iloc[0]
            bom.append({"Item": "Internal Wiring Harness & Copper Busbars", "Specification": bb_r.get("Specification","Copper Harness"), "Brand": bb_r.get("Brand","Polycab"), "Qty": 1, "Unit Price": bb_r.get("Unit_Price_INR",5500) / ex_rate})
            bom.append({"Item": "Control Accessories & Relays", "Specification": acc_r.get("Specification","Relays & Terminal Blocks"), "Brand": acc_r.get("Brand","Generic"), "Qty": 1, "Unit Price": acc_r.get("Unit_Price_INR",3500) / ex_rate})

            if freight_price > 0:
                bom.append({"Item": "Transit Packing & Freight Logistics", "Specification": f_opt.split('(')[0].strip(), "Brand": "Logistics", "Qty": 1, "Unit Price": freight_price})

            if st.session_state.get("custom_bom_items"):
                for item in st.session_state["custom_bom_items"]:
                    bom.append(item.copy())

            res_df = pd.DataFrame(bom)
            res_df["Total Material Cost"] = res_df["Qty"] * res_df["Unit Price"]
            return res_df

        bom_df = build_bom(preferred_brand)

    else:
        st.subheader("🏢 Full Multi-Feeder Panel / MCC Board Configurator")
        mc1, mc2, mc3 = st.columns(3)
        client_name = mc1.text_input("Client Name", "Maharashtra Water Works Ltd")
        preferred_brand = mc2.selectbox("Primary Switchgear Brand", ["Siemens", "Schneider", "L&T", "ABB", "Danfoss", "Delta"])
        panel_type = mc3.text_input("Panel Board Title", "Multi-Feeder Motor Control Center (MCC)")

        mc4, mc5, mc6 = st.columns(3)
        margin_pct = mc4.slider("Margin (%)", 5, 40, 18)
        labor_cost_inr = mc5.number_input("Total Assembly, Busbar & Wiring Labor (₹)", value=18500, step=1000)
        incomer_type = mc6.selectbox("Main Incomer Breaker Rating", ["100A 3P MCCB (25kA)", "250A 3P MCCB (36kA)", "400A 3P MCCB (36kA)", "630A 3P MCCB (50kA)", "800A 4P Drawout ACB", "1250A 4P Drawout ACB", "1600A 4P Drawout ACB"], index=3)

        curr_opt = "INR (₹)"
        curr_sym = "₹"
        ex_rate = 1.0
        labor_cost = labor_cost_inr
        motor_kw = "Multi-Feeder"

        if "multi_feeder_schedule" not in st.session_state:
            st.session_state["multi_feeder_schedule"] = [
                {"Feeder Name": "Feeder 1 - Raw Water Pump", "Starter Type": "VFD Feeder", "Rating": "22 kW", "Brand": preferred_brand, "Qty": 2, "Unit Price (₹)": 48000.0},
                {"Feeder Name": "Feeder 2 - High Lift Pump", "Starter Type": "Star-Delta", "Rating": "37 kW", "Brand": preferred_brand, "Qty": 2, "Unit Price (₹)": 28000.0},
                {"Feeder Name": "Feeder 3 - Agitator Motor", "Starter Type": "DOL Starter", "Rating": "7.5 kW", "Brand": preferred_brand, "Qty": 4, "Unit Price (₹)": 8500.0},
                {"Feeder Name": "Feeder 4 - Auxiliary Distribution", "Starter Type": "Feeder MCCB", "Rating": "100A 3P", "Brand": preferred_brand, "Qty": 3, "Unit Price (₹)": 5200.0}
            ]

        st.markdown("#### 📝 Edit Feeder Schedule & Quantities")
        edited_feeders = st.data_editor(st.session_state["multi_feeder_schedule"], num_rows="dynamic", key="mcc_feeder_editor")

        bom_list = [
            {"Item": "Multi-Bay MCC Panel Enclosure Frame", "Specification": "IP55 Floor Mount Dual-Column", "Brand": "Standard Sheet Metal", "Qty": 1, "Unit Price": 45000.0}
        ]

        incomer_components = get_incomer_default_components(incomer_type, preferred_brand, ex_rate=1.0)
        bom_list.extend(incomer_components)

        for f in edited_feeders:
            f_type = f.get('Starter Type','Starter')
            f_rating = f.get('Rating','15 kW')
            f_brand = f.get('Brand', preferred_brand)
            f_qty = int(f.get('Qty', 1))
            
            feeder_sub_items = get_feeder_subcomponents(f_type, f_rating, f_brand, ex_rate=1.0)
            if feeder_sub_items:
                for sub in feeder_sub_items:
                    sub_copy = sub.copy()
                    sub_copy["Qty"] = sub_copy["Qty"] * f_qty
                    bom_list.append(sub_copy)
            else:
                bom_list.append({
                    "Item": f"{f.get('Feeder Name','Feeder')} ({f_type})",
                    "Specification": f"{f_rating} ({f_type})",
                    "Brand": f_brand,
                    "Qty": f_qty,
                    "Unit Price": float(f.get('Unit Price (₹)', 10000.0))
                })

        bom_list.append({"Item": "Main Copper Busbar & Power Distribution Harness", "Specification": "EC Grade Copper Busbars", "Brand": "Polycab", "Qty": 1, "Unit Price": 24000.0})
        bom_list.append({"Item": "Control Transformers, SMPS & Interlocks", "Specification": "24V DC / 230V AC", "Brand": "Generic", "Qty": 1, "Unit Price": 12500.0})

        bom_df = pd.DataFrame(bom_list)
        bom_df["Total Material Cost"] = bom_df["Qty"] * bom_df["Unit Price"]

    st.dataframe(bom_df)

    mat_total = bom_df["Total Material Cost"].sum()
    factory_cost = mat_total + labor_cost
    sell_price = factory_cost * (1 + margin_pct / 100.0)
    ship_res = calculate_panel_shipping_weight(2000, 1600, 600, "75 kW" if quote_mode != "⚡ Single Feeder Panel" else motor_kw)

    st.divider()
    m1, m2, m3, m4, m5 = st.columns([1, 1, 1, 1.1, 1.3])
    m1.metric("Raw Material", f"{curr_sym} {mat_total:,.2f}")
    m2.metric("Labor & Assembly", f"{curr_sym} {labor_cost:,.2f}")
    m3.metric("Est. Gross Weight", f"{ship_res['gross_weight_kg']} kg")
    m4.metric("Total Factory Cost", f"{curr_sym} {factory_cost:,.2f}")
    m5.metric(f"Commercial Quote ({margin_pct}%)", f"{curr_sym} {sell_price:,.2f}")

    if quote_mode == "⚡ Single Feeder Panel":
        with st.expander("⚡ Instant Multi-Brand Price Comparison Matrix (All Brands)", expanded=False):
            comp_data = []
            for b in b_list:
                b_df = build_bom(b)
                b_mat = b_df["Total Material Cost"].sum()
                b_tot = (b_mat + labor_cost) * (1 + margin_pct / 100.0)
                comp_data.append({
                    "Switchgear Brand": b,
                    f"Raw Material ({curr_sym})": f"{curr_sym} {b_mat:,.2f}",
                    f"Total Quote Price ({curr_sym})": f"{curr_sym} {b_tot:,.2f}",
                    "Difference vs Selected": f"{curr_sym} {b_tot - sell_price:+,.2f}"
                })
            st.dataframe(pd.DataFrame(comp_data))

    b1, b2, b3, b4, b5 = st.columns(5)
    
    excel_proposal_bytes = generate_excel_proposal(client_name, panel_type, motor_kw, preferred_brand, curr_opt, sell_price, curr_sym, bom_df)
    b1.download_button("📊 Excel Proposal", excel_proposal_bytes, f"Proposal_{client_name.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    pdf_bytes = generate_pdf_quotation(client_name, panel_type, motor_kw, preferred_brand, bom_df, labor_cost, margin_pct, sell_price, curr_sym)
    b2.download_button("📄 PDF Quote", pdf_bytes.getvalue(), f"Aryavarta_Quote_{client_name.replace(' ', '_')}.pdf", mime="application/pdf")
    
    pi_bytes = generate_proforma_invoice_pdf(client_name, panel_type, motor_kw, sell_price, curr_sym)
    b3.download_button("🧾 Proforma Invoice", pi_bytes.getvalue(), f"Proforma_Invoice_{client_name.replace(' ', '_')}.pdf", mime="application/pdf")
    
    tds_bytes = generate_tds_pdf(client_name, panel_type, motor_kw, preferred_brand)
    b4.download_button("📋 Technical Specs (TDS)", tds_bytes.getvalue(), f"TDS_{client_name.replace(' ', '_')}.pdf", mime="application/pdf")

    buf_zip = io.BytesIO()
    with zipfile.ZipFile(buf_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"Aryavarta_Quote_{client_name}.pdf", pdf_bytes.getvalue())
        zf.writestr(f"Proforma_Invoice_{client_name}.pdf", pi_bytes.getvalue())
        zf.writestr(f"Technical_Data_Sheet_{client_name}.pdf", tds_bytes.getvalue())
        zf.writestr(f"BOM_{client_name}.xlsx", excel_proposal_bytes)
    b5.download_button("📦 Full Proposal (.ZIP)", buf_zip.getvalue(), f"Proposal_Package_{client_name.replace(' ', '_')}.zip", mime="application/zip")

    st.subheader("📐 Auto-Generated Interactive CAD GA Drawing & Power SLD Diagram Engine")
    with st.expander("🛠️ Interactive Panel Sizing & Component Dimensions Customizer", expanded=True):
        ga_c1, ga_c2, ga_c3, ga_c4 = st.columns(4)
        panel_h_input = ga_c1.number_input("Panel Frame Height (mm)", value=2000, step=100)
        bay_w_input = ga_c2.number_input("Bay Width (mm)", value=800, step=100)
        panel_d_input = ga_c3.number_input("Panel Depth (mm)", value=600, step=50)
        cable_alley_w_input = ga_c4.number_input("Cable Alley Width (mm)", value=200, step=50)
        
        busbar_pos_input = st.radio("Main Busbar Chamber Position", ["Top", "Bottom"], horizontal=True)

    # Perform multi-bay spatial overflow auto-layout calculation
    computed_bays = auto_layout_panel_components(bom_df, panel_h_input, bay_w_input, busbar_pos_input, cable_alley_w_input)

    tab_int_ga, tab_outer_ga, tab_sld, tab_spatial = st.tabs([
        "🖼️ Internal GA (Interior Mounting)", 
        "🚪 Outer GA (Front Door & Meters)", 
        "⚡ Power Single Line Diagram (SLD)",
        "📐 Spatial Footprint & Bay Analysis"
    ])

    with tab_int_ga:
        int_ga_svg = generate_internal_ga_svg(computed_bays, panel_h_input, bay_w_input, busbar_pos_input, cable_alley_w_input)
        st.markdown(int_ga_svg, unsafe_allow_html=True)
        st.download_button("📥 Download Internal GA (.SVG)", int_ga_svg, f"Internal_GA_{client_name.replace(' ', '_')}.svg", mime="image/svg+xml")

    with tab_outer_ga:
        outer_ga_svg = generate_outer_ga_svg(computed_bays, panel_h_input, bay_w_input, preferred_brand, panel_type)
        st.markdown(outer_ga_svg, unsafe_allow_html=True)
        st.download_button("📥 Download Outer GA (.SVG)", outer_ga_svg, f"Outer_GA_{client_name.replace(' ', '_')}.svg", mime="image/svg+xml")

    with tab_sld:
        sld_svg = generate_detailed_sld_svg(selected_incomer if quote_mode=="⚡ Single Feeder Panel" else incomer_type, bom_df, preferred_brand)
        st.markdown(sld_svg, unsafe_allow_html=True)
        st.download_button("📥 Download Power SLD (.SVG)", sld_svg, f"Power_SLD_{client_name.replace(' ', '_')}.svg", mime="image/svg+xml")

    with tab_spatial:
        st.markdown(f"#### 📊 Spatial Overflow Summary: **{len(computed_bays)} Panel Column / Bay(s) Required**")
        st.info(f"💡 Total Panel Enclosure Shipping Footprint: **{panel_h_input}mm (H) x {(len(computed_bays)*bay_w_input)+cable_alley_w_input}mm (W) x {panel_d_input}mm (D)**.")
        
        flat_comps = []
        for bay_idx, bay in enumerate(computed_bays):
            for c in bay["components"]:
                flat_comps.append({
                    "Assigned Bay": f"Bay #{bay_idx+1}",
                    "Component Name": c["name"],
                    "Position (X,Y mm)": f"X: {c['x']}mm, Y: {c['y']}mm",
                    "Width (mm)": c["w"],
                    "Height (mm)": c["h"],
                    "Component Type": c["type"]
                })
        st.dataframe(pd.DataFrame(flat_comps))

    if st.button("💾 Save Quote to Quotation Register & Pipeline"):
        save_quote_to_history(client_name, panel_type, motor_kw, preferred_brand, sell_price * ex_rate, margin_pct)
        st.success(f"Quote for {client_name} ({curr_sym} {sell_price:,.2f}) saved to Quotation Register!")

elif menu == "MCC Multi-Feeder Panel Sizer":
    st.header("🏭 MCC Multi-Feeder Load & Incomer Sizer")
    st.markdown("Sizes main incomer ACB/MCCB, busbars, and frame dimensions for multi-motor Control Centers.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("1. VFD Feeders")
        v_kw = st.number_input("VFD Motor Power (kW)", value=22.0, step=2.5)
        v_q = st.number_input("VFD Feeders Qty", value=2, min_value=0)
    with c2:
        st.subheader("2. Star-Delta Feeders")
        sd_kw_v = st.number_input("Star-Delta Motor Power (kW)", value=37.0, step=2.5)
        sd_q = st.number_input("Star-Delta Feeders Qty", value=2, min_value=0)
    with c3:
        st.subheader("3. DOL Feeders & Diversity")
        dol_kw_v = st.number_input("DOL Motor Power (kW)", value=7.5, step=2.5)
        dol_q = st.number_input("DOL Feeders Qty", value=4, min_value=0)
        div = st.slider("Plant Diversity Factor", 0.4, 1.0, 0.75, step=0.05)

    mcc_res = calculate_mcc_multi_feeder(v_kw, v_q, sd_kw_v, sd_q, dol_kw_v, dol_q, div)
    st.divider()
    mc1, mc2, mc3, mc4 = st.columns([1, 1, 1.2, 1.5])
    mc1.metric("Total Connected Load", f"{mcc_res['tot_kw']} kW")
    mc2.metric("Demand Operating Load", f"{mcc_res['demand_kw']} kW")
    mc3.metric("Plant FLC Current", f"{mcc_res['flc']} A")
    mc4.metric("Recommended Main Incomer", mcc_res['incomer'])
    st.info(f"💡 **Recommended Main Copper Busbar:** {mcc_res['busbar']} | **Suggested Enclosure Frame:** {mcc_res['frame']}")

elif menu == "Cable & Busbar Calculator":
    st.header("📐 Electrical Cable & Busbar Sizing (IS 3961 / IEC 61439)")
    c1, c2, c3 = st.columns(3)
    kw = c1.number_input("Motor Rating (kW)", min_value=0.1, max_value=300.0, value=15.0, step=0.1, format="%.1f")
    dist = c1.number_input("Cable Run Distance (Meters)", 30, step=5)
    cond = c2.radio("Material", ["Copper", "Aluminum"])
    p_type = c3.selectbox("Panel Type", ["VFD Panel", "Star-Delta Panel", "DOL Panel"])
    amb_t = c2.slider("Ambient Temperature (°C)", 30, 55, 45)
    tray_cables = c3.selectbox("Cables Grouped in Tray", [1, 3, 6])
    
    res = calculate_cable_and_busbar(kw, dist, cond, amb_t, tray_cables)
    m1, m2, m3, m4, m5 = st.columns([1, 1.1, 1.3, 1.3, 1.3])
    m1.metric("Motor FLC", f"{res['flc']} A")
    m2.metric("Recommended MCCB", res['rec_mccb'])
    m3.metric("Cable Cross-Section", res['rec_cable'])
    m4.metric("Busbar Size", res['rec_busbar'])
    m5.metric("Max Run (3% ΔV Limit)", f"{res['max_len_3pct']} Meters")
    st.info(f"💡 **Thermal, Derating & Current Density Check:** Derating Factor: **{res['derating']}** ({amb_t}°C Ambient + {tray_cables} Cable Tray Grouping). Design Current: {res['design_current']}A. Busbar Current Density: **{res['curr_density']} A/mm²**. Est. ΔT: **{res['temp_rise']}°C** (Passes IS 8623 < 35°C rise limit). Voltage Drop over {dist}m: {res['v_drop']}V.")

elif menu == "Quotation Register & Pipeline":
    st.header("📈 Quotation Register & Sales Pipeline Tracker")
    st.markdown("Track generated customer quotes, follow-up statuses, and total sales pipeline value offline.")
    if os.path.exists(QUOTES_FILE):
        q_df = pd.read_excel(QUOTES_FILE)
        st.divider()
        q1, q2, q3, q4 = st.columns(4)
        q1.metric("Total Quotes Issued", len(q_df))
        q2.metric("Total Pipeline Value", f"₹ {q_df['Quote_Price_INR'].sum():,.2f}")
        won_val = q_df[q_df['Status']=='Won']['Quote_Price_INR'].sum() if 'Won' in q_df['Status'].values else 0
        q3.metric("Won Orders Value", f"₹ {won_val:,.2f}")
        won_count = len(q_df[q_df['Status']=='Won']) if 'Won' in q_df['Status'].values else 0
        won_pct = round((won_count / len(q_df)) * 100, 1) if len(q_df)>0 else 0
        q4.metric("Conversion Rate", f"{won_pct}%")

        reg_buf = io.BytesIO()
        with pd.ExcelWriter(reg_buf, engine="openpyxl") as reg_w:
            q_df.to_excel(reg_w, index=False, sheet_name="Quotes_Register")
        reg_buf.seek(0)
        st.download_button("📥 Export Full Quotes Register (.XLSX)", reg_buf.getvalue(), f"Quotes_Register_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.subheader("📊 Quotations Log & Status Manager")
        edited_q = st.data_editor(q_df, num_rows="dynamic")
        if st.button("💾 Update Register Status"):
            edited_q.to_excel(QUOTES_FILE, index=False)
            st.success("Quotation Register updated successfully!")
    else:
        st.info("No quotes logged yet. Create a quote in 'Create Panel Quote' and click 'Save Quote to Quotation Register'.")

elif menu == "Manage Price Database":
    st.header("⚙️ Local Price Database Manager")
    edited = st.data_editor(df_prices, num_rows="dynamic")
    if st.button("💾 Save Changes"):
        edited.to_excel(DB_FILE, index=False)
        st.cache_data.clear()
        st.success("Database updated successfully!")
