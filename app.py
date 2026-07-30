import io
import json
import math
import os
import re
import html
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
    "Power Contactor (32A AC3)": {"Height_mm": 85, "Width_mm": 65, "Depth_mm": 95, "Type": "Contactor"},
    "Power Contactor (65A AC3)": {"Height_mm": 125, "Width_mm": 85, "Depth_mm": 125, "Type": "Contactor"},
    "Power Contactor (110A AC3)": {"Height_mm": 160, "Width_mm": 120, "Depth_mm": 150, "Type": "Contactor"},
    "Thermal Overload Relay (TOR)": {"Height_mm": 90, "Width_mm": 65, "Depth_mm": 85, "Type": "Relay"},
    "Star-Delta Timer Relay": {"Height_mm": 90, "Width_mm": 45, "Depth_mm": 75, "Type": "Relay"},
    "PLC Base CPU Unit": {"Height_mm": 100, "Width_mm": 140, "Depth_mm": 75, "Type": "PLC"},
    "PLC I/O Expansion Module": {"Height_mm": 100, "Width_mm": 60, "Depth_mm": 75, "Type": "PLC"},
    "Touchscreen HMI (7 Inch)": {"Height_mm": 140, "Width_mm": 200, "Depth_mm": 45, "Type": "HMI"},
    "Touchscreen HMI (10 Inch)": {"Height_mm": 200, "Width_mm": 270, "Depth_mm": 50, "Type": "HMI"},
    "24V DC SMPS (10A)": {"Height_mm": 125, "Width_mm": 70, "Depth_mm": 125, "Type": "PowerSupply"},
    "Control Tx (500VA)": {"Height_mm": 140, "Width_mm": 130, "Depth_mm": 110, "Type": "Transformer"},
    "3% AC Line Reactor Choke": {"Height_mm": 210, "Width_mm": 180, "Depth_mm": 150, "Type": "Reactor"},
    "Digital MFM Meter": {"Height_mm": 96, "Width_mm": 96, "Depth_mm": 60, "Type": "Meter"},
    "Metering CT Set (Set of 3)": {"Height_mm": 110, "Width_mm": 240, "Depth_mm": 60, "Type": "CT"},
    "Space Heater + Thermostat": {"Height_mm": 120, "Width_mm": 90, "Depth_mm": 50, "Type": "Heater"},
    "Panel Cooling Filter Fan Unit": {"Height_mm": 250, "Width_mm": 250, "Depth_mm": 120, "Type": "Fan"},
    "Power DIN Terminal Block Strip": {"Height_mm": 60, "Width_mm": 250, "Depth_mm": 45, "Type": "Terminal"}
}

def xml_escape(text):
    """Safely escapes XML string entities to prevent SVG render errors."""
    return html.escape(str(text))

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

def calculate_panel_shipping_weight(h_mm, w_mm, d_mm, kw_str):
    """Calculates estimated gross panel shipping weight and cubic volume."""
    enc_w = calculate_enclosure_fabrication(h_mm, w_mm, d_mm, 1.6, 82, 220)["weight_kg"]
    kw_val = float(kw_str.split()[0]) if kw_str and kw_str.split()[0].replace('.','',1).isdigit() else 15.0
    sw_w = 8.0 + (kw_val * 0.55)
    cu_w = 4.0 + (kw_val * 0.25)
    tot_w = round(enc_w + sw_w + cu_w, 1)
    vol_m3 = round(((h_mm + 100) * (w_mm + 100) * (d_mm + 100)) / 1_000_000_000.0, 2)
    return {"gross_weight_kg": tot_w, "volume_m3": vol_m3}

def calculate_cable_and_busbar(kw_val, distance_m, conductor_type, amb_temp=40, cables_in_tray=1):
    flc = (kw_val * 1000) / (1.732 * 415 * 0.85 * 0.88)
    k_temp = 1.0 if amb_temp <= 40 else (0.91 if amb_temp <= 45 else 0.82)
    k_group = 1.0 if cables_in_tray <= 1 else (0.82 if cables_in_tray <= 3 else 0.73)
    total_derating = k_temp * k_group
    design_current = (flc * 1.25) / total_derating
    
    cables = [(2.5,24),(4,32),(6,41),(10,57),(16,76),(25,101),(35,125),(50,150),(70,190),(95,235),(120,270),(150,310),(185,355),(240,420),(300,480)] if conductor_type=="Copper" else [(4,25),(6,32),(10,44),(16,59),(25,78),(35,97),(50,116),(70,147),(95,182),(120,210),(150,240),(185,275),(240,325),(300,375)]
    sqmm = next((s for s, a in cables if a >= design_current), cables[-1][0])
    req_mm2 = design_current / (1.5 if conductor_type=="Copper" else 0.9)
    busbars = [(15,3),(20,3),(20,5),(25,5),(30,5),(40,5),(50,6),(60,8),(80,10),(100,10)]
    bsize = next((f"{w}x{t} mm" for w, t in busbars if (w*t) >= req_mm2), "100x10 mm")
    mccb_r = max(32, math.ceil(flc * 1.5 / 10.0) * 10)
    
    rho_val = 0.0175 if conductor_type == "Copper" else 0.0282
    p_loss_wm = 3 * (design_current ** 2) * (rho_val / req_mm2)
    temp_rise_c = round(min(65.0, 12.0 + (p_loss_wm * 0.45)), 1)
    max_len_m = round((12.45 * sqmm) / (1.732 * design_current * rho_val), 1) if design_current > 0 else 0
    curr_density = round(design_current / (float(bsize.split('x')[0]) * float(bsize.split('x')[1].split()[0])), 2) if "x" in bsize else 1.5

    return {"flc": round(flc,2), "design_current": round(design_current,2), "rec_mccb": f"{mccb_r}A MCCB", "rec_cable": f"{sqmm} sq.mm ({conductor_type})", "rec_busbar": f"{bsize} ({conductor_type})", "v_drop": round((1.732*design_current*distance_m*rho_val)/sqmm, 2), "temp_rise": temp_rise_c, "derating": round(total_derating, 2), "max_len_3pct": max_len_m, "curr_density": curr_density}

def calculate_mcc_multi_feeder(vfd_kw, vfd_qty, sd_kw, sd_qty, dol_kw, dol_qty, diversity):
    tot_kw = (vfd_kw * vfd_qty) + (sd_kw * sd_qty) + (dol_kw * dol_qty)
    demand_kw = tot_kw * diversity
    flc = (demand_kw * 1000) / (1.732 * 415 * 0.85 * 0.88) if demand_kw > 0 else 0
    incomer_amps = max(63, math.ceil(flc * 1.25 / 10.0) * 10)
    incomer_type = f"{incomer_amps}A MCCB" if incomer_amps <= 630 else f"{incomer_amps}A Air Circuit Breaker (ACB)"
    req_mm2 = (flc * 1.25) / 1.5
    busbars = [(25,5),(30,5),(40,5),(50,6),(60,8),(80,10),(100,10)]
    rec_busbar = next((f"{w}x{t} mm Copper" for w, t in busbars if (w*t) >= req_mm2), "2x 100x10 mm Copper")
    tot_feeders = vfd_qty + sd_qty + dol_qty
    frame_size = "Single Column (1800x800x600)" if tot_feeders <= 4 else "Dual Column (2000x1600x600)" if tot_feeders <= 8 else "Multi-Bay MCC Panel (2000x2400x800)"
    return {"tot_kw": round(tot_kw, 1), "demand_kw": round(demand_kw, 1), "flc": round(flc, 1), "incomer": incomer_type, "busbar": rec_busbar, "frame": frame_size}

def calculate_star_delta_starter(motor_kw):
    flc = (motor_kw * 1000.0) / (1.732 * 415.0 * 0.85 * 0.88)
    i_phase = flc / 1.732
    main_delta_a = math.ceil(i_phase * 1.15)
    star_a = math.ceil((flc / 3.0) * 1.15)
    tor_set = round(flc * 0.58, 1)
    mccb_r = max(32, math.ceil(flc * 1.5 / 10.0) * 10)
    return {"flc": round(flc, 1), "i_phase": round(i_phase, 1), "main_delta_contactor": f"{main_delta_a}A AC3 (Qty 2)", "star_contactor": f"{star_a}A AC3 (Qty 1)", "tor_range": f"{round(tor_set*0.85, 1)}A - {round(tor_set*1.15, 1)}A (Set @ {tor_set}A)", "mccb": f"{mccb_r}A MCCB"}

def calculate_mpcb_and_protection(motor_kw, load_type="Standard Load (Class 10)"):
    flc = (motor_kw * 1000.0) / (1.732 * 415.0 * 0.85 * 0.88)
    mpcb_min = round(flc * 0.9, 1)
    mpcb_max = round(flc * 1.15, 1)
    mag_trip_a = round(flc * 13.0, 1)
    trip_class = "Class 10 (Trips in 4-10s @ 7.2x FLC)" if "Class 10" in load_type else ("Class 20 (Trips in 10-20s @ 7.2x FLC)" if "Class 20" in load_type else "Class 30 (Trips in 20-30s @ 7.2x FLC)")
    rec_type = "MPCB (Type 2 Coordination)" if flc <= 100 else "MCCB + Electronic Overload Relay (EOR)"
    return {"flc": round(flc, 1), "thermal_range": f"{mpcb_min}A - {mpcb_max}A (Set @ {round(flc, 1)}A)", "mag_trip": f"{mag_trip_a} A (Instantaneous)", "trip_class": trip_class, "protection_device": rec_type}

def calculate_enclosure_fabrication(h_mm, w_mm, d_mm, thickness_mm, crca_rate_per_kg, coating_rate_per_sqm, insulator_qty=12, lock_qty=2):
    area_m2 = 2.0 * ((h_mm*w_mm) + (w_mm*d_mm) + (h_mm*d_mm)) / 1_000_000.0 * 1.35
    weight_kg = area_m2 * (thickness_mm / 1000.0) * 7850.0
    crca_cost = weight_kg * crca_rate_per_kg
    coating_cost = area_m2 * coating_rate_per_sqm
    hardware_cost = (insulator_qty * 85.0) + (lock_qty * 350.0) + 650.0
    total_cost = crca_cost + coating_cost + hardware_cost
    return {
        "area_m2": round(area_m2, 2), "weight_kg": round(weight_kg, 1),
        "crca_cost": round(crca_cost, 2), "coating_cost": round(coating_cost, 2),
        "hardware_cost": round(hardware_cost, 2), "total_enc_cost": round(total_cost, 2)
    }

def get_incomer_default_components(incomer_type, brand="Siemens", ex_rate=1.0):
    inc_lower = str(incomer_type).lower()
    if "630" in inc_lower:
        return [
            {"Item": "Main Incomer 630A 3P MCCB", "Specification": f"630A 3P 50kA Microprocessor/TM ({brand} 3VA/3VM)", "Brand": brand, "Qty": 1, "Unit Price": 28500.0 / ex_rate},
            {"Item": "Extended Door Rotary Handle Kit (ROM)", "Specification": "630A Door Operating Mechanism + Shaft", "Brand": brand, "Qty": 1, "Unit Price": 3200.0 / ex_rate},
            {"Item": "Terminal Spreader Links Kit", "Specification": "630A Busbar/Cable Spreader Extension (Set of 3)", "Brand": brand, "Qty": 1, "Unit Price": 2800.0 / ex_rate},
            {"Item": "Phase Barriers / Insulating Shrouds", "Specification": "630A Inter-Phase Barrier Set", "Brand": brand, "Qty": 1, "Unit Price": 950.0 / ex_rate},
            {"Item": "Auxiliary & Alarm Contact Block", "Specification": "1NO+1NC Aux + 1 Trip Alarm Switch", "Brand": brand, "Qty": 1, "Unit Price": 1850.0 / ex_rate},
            {"Item": "Shunt Trip Release Coil", "Specification": "230V AC Remote Emergency Trip Coil", "Brand": brand, "Qty": 1, "Unit Price": 2400.0 / ex_rate},
            {"Item": "Incomer Metering CT Set", "Specification": "600/5A Class 0.5 Measuring CTs (Set of 3)", "Brand": "Rishabh", "Qty": 1, "Unit Price": 3600.0 / ex_rate},
            {"Item": "Incomer Control & Meter Protection MCB", "Specification": "6A 3P 10kA C-Curve Control MCB (Siemens 5SY)", "Brand": brand, "Qty": 1, "Unit Price": 850.0 / ex_rate}
        ]
    elif "400" in inc_lower:
        return [
            {"Item": "Main Incomer 400A 3P MCCB", "Specification": f"400A 3P 36kA/50kA TM/Microprocessor ({brand})", "Brand": brand, "Qty": 1, "Unit Price": 18500.0 / ex_rate},
            {"Item": "Extended Door Rotary Handle Kit (ROM)", "Specification": "400A Door Operating Mechanism + Shaft", "Brand": brand, "Qty": 1, "Unit Price": 2600.0 / ex_rate},
            {"Item": "Terminal Spreader Links Kit", "Specification": "400A Cable/Busbar Spreader Extension (Set of 3)", "Brand": brand, "Qty": 1, "Unit Price": 2200.0 / ex_rate},
            {"Item": "Phase Barriers / Insulating Shrouds", "Specification": "400A Inter-Phase Barrier Set", "Brand": brand, "Qty": 1, "Unit Price": 750.0 / ex_rate},
            {"Item": "Auxiliary & Alarm Contact Block", "Specification": "1NO+1NC Aux Contact", "Brand": brand, "Qty": 1, "Unit Price": 1450.0 / ex_rate},
            {"Item": "Shunt Trip Release Coil", "Specification": "230V AC Remote Emergency Trip Coil", "Brand": brand, "Qty": 1, "Unit Price": 2100.0 / ex_rate},
            {"Item": "Incomer Metering CT Set", "Specification": "400/5A Class 0.5 Measuring CTs (Set of 3)", "Brand": "Rishabh", "Qty": 1, "Unit Price": 2900.0 / ex_rate},
            {"Item": "Incomer Protection MCB", "Specification": "6A 3P 10kA Control MCB", "Brand": brand, "Qty": 1, "Unit Price": 850.0 / ex_rate}
        ]
    elif "250" in inc_lower:
        return [
            {"Item": "Main Incomer 250A 3P MCCB", "Specification": f"250A 3P 36kA Thermal-Magnetic ({brand})", "Brand": brand, "Qty": 1, "Unit Price": 11500.0 / ex_rate},
            {"Item": "Extended Door Rotary Handle Kit (ROM)", "Specification": "250A Door Operating Mechanism + Shaft", "Brand": brand, "Qty": 1, "Unit Price": 2100.0 / ex_rate},
            {"Item": "Terminal Spreader Links Kit", "Specification": "250A Terminal Spreader Extension (Set of 3)", "Brand": brand, "Qty": 1, "Unit Price": 1600.0 / ex_rate},
            {"Item": "Incomer Metering CT Set", "Specification": "250/5A Class 0.5 Measuring CTs (Set of 3)", "Brand": "Rishabh", "Qty": 1, "Unit Price": 2400.0 / ex_rate}
        ]
    elif "100" in inc_lower:
        return [
            {"Item": "Main Incomer 100A 3P MCCB", "Specification": f"100A 3P 25kA Thermal-Magnetic ({brand})", "Brand": brand, "Qty": 1, "Unit Price": 4800.0 / ex_rate},
            {"Item": "Extended Door Rotary Handle Kit (ROM)", "Specification": "100A Door Operating Mechanism + Shaft", "Brand": brand, "Qty": 1, "Unit Price": 1600.0 / ex_rate}
        ]
    else: # ACB
        rating_str = "800A" if "800" in inc_lower else ("1250A" if "1250" in inc_lower else "1600A")
        acb_base_price = 95000.0 if "800" in inc_lower else (135000.0 if "1250" in inc_lower else 175000.0)
        return [
            {"Item": f"Main Incomer {rating_str} 4P Drawout ACB", "Specification": f"{rating_str} 4P 50kA Microprocessor ETU ({brand} 3WA/3WL)", "Brand": brand, "Qty": 1, "Unit Price": acb_base_price / ex_rate},
            {"Item": "ACB Motorized Racking & Shunt Trip Release", "Specification": "230V AC Motor Mechanism + Shunt Coil", "Brand": brand, "Qty": 1, "Unit Price": 18500.0 / ex_rate},
            {"Item": "Incomer Metering CT Set & Protection MCB", "Specification": f"{rating_str}/5A Class 0.2S CT Set + 6A 3P Control MCB", "Brand": "Rishabh", "Qty": 1, "Unit Price": 6800.0 / ex_rate}
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
            {"Item": f"Star-Delta Incomer MPCB / MCCB ({mccb_a}A)", "Specification": f"{mccb_a}A 3P 25kA/36kA Motor Duty Breaker ({brand} 3RV/3VM)", "Brand": brand, "Qty": 1, "Unit Price": (3500 + mccb_a * 18) / ex_rate},
            {"Item": f"Main Power Contactor ({main_a}A AC-3)", "Specification": f"{main_a}A 3P 230V AC Coil Power Contactor ({brand} 3RT)", "Brand": brand, "Qty": 1, "Unit Price": (1800 + main_a * 25) / ex_rate},
            {"Item": f"Delta Power Contactor ({main_a}A AC-3)", "Specification": f"{main_a}A 3P 230V AC Coil Power Contactor ({brand} 3RT)", "Brand": brand, "Qty": 1, "Unit Price": (1800 + main_a * 25) / ex_rate},
            {"Item": f"Star Power Contactor ({star_a}A AC-3)", "Specification": f"{star_a}A 3P 230V AC Coil Power Contactor ({brand} 3RT)", "Brand": brand, "Qty": 1, "Unit Price": (1200 + star_a * 22) / ex_rate},
            {"Item": "Mechanical & Electrical Interlock Block Set", "Specification": "Star-Delta Contactor Interlock Kit", "Brand": brand, "Qty": 1, "Unit Price": 1250.0 / ex_rate},
            {"Item": "Electronic Star-Delta Timer Relay", "Specification": "0.1s - 30s 230VAC Electronic Timer ({brand} 3RP)", "Brand": brand, "Qty": 1, "Unit Price": 2200.0 / ex_rate},
            {"Item": "Thermal Overload Relay (TOR)", "Specification": f"Class 10 Adjustable Bimetallic Relay ({brand} 3RU)", "Brand": brand, "Qty": 1, "Unit Price": 2400.0 / ex_rate},
            {"Item": "Door Start/Stop Buttons & Signal LEDs", "Specification": "Flush Pushbuttons + Red/Green/Yellow LEDs ({brand} 3SB)", "Brand": brand, "Qty": 1, "Unit Price": 1400.0 / ex_rate}
        ]
    elif "DOL" in feeder_type:
        dol_cont_a = math.ceil(flc * 1.15)
        return [
            {"Item": f"DOL Incomer MPCB ({dol_cont_a}A)", "Specification": f"Adjustable Thermal-Magnetic MPCB ({brand} 3RV)", "Brand": brand, "Qty": 1, "Unit Price": 3200.0 / ex_rate},
            {"Item": f"DOL Power Contactor ({dol_cont_a}A AC-3)", "Specification": f"{dol_cont_a}A 3P Power Contactor ({brand} 3RT)", "Brand": brand, "Qty": 1, "Unit Price": (1400 + dol_cont_a * 20) / ex_rate},
            {"Item": "Auxiliary Contact Block", "Specification": "1NO+1NC Front Snap Auxiliary Block", "Brand": brand, "Qty": 1, "Unit Price": 650.0 / ex_rate},
            {"Item": "Start/Stop Pushbuttons & Status Lamps", "Specification": "Green Start, Red Stop, Amber Trip LED ({brand} 3SB)", "Brand": brand, "Qty": 1, "Unit Price": 1100.0 / ex_rate}
        ]
    elif "VFD" in feeder_type:
        mccb_a = max(32, math.ceil(flc * 1.25 / 10.0) * 10)
        return [
            {"Item": f"VFD Power Unit ({rating_kw_str})", "Specification": f"{rating_kw_str} 415V Heavy Duty Drive ({brand} Sinamics G120C/V20)", "Brand": brand, "Qty": 1, "Unit Price": (22000 + kw_val * 950) / ex_rate},
            {"Item": f"VFD Incomer MCCB / MPCB ({mccb_a}A)", "Specification": f"{mccb_a}A 3P 25kA Incomer Breaker ({brand} 3VM)", "Brand": brand, "Qty": 1, "Unit Price": (3200 + mccb_a * 15) / ex_rate},
            {"Item": "Semiconductor Fast-Acting Fuses (aR)", "Specification": f"Fast Semiconductor Fuse Set ({brand} 3NE)", "Brand": brand, "Qty": 3, "Unit Price": 1800.0 / ex_rate},
            {"Item": "3% Input AC Line Reactor Choke", "Specification": "Harmonic Mitigation Reactor", "Brand": "Elcon", "Qty": 1, "Unit Price": (2500 + kw_val * 120) / ex_rate},
            {"Item": "Line Isolation Power Contactor", "Specification": f"{math.ceil(flc)}A 3P Line Contactor ({brand} 3RT)", "Brand": brand, "Qty": 1, "Unit Price": (1600 + flc * 20) / ex_rate},
            {"Item": "Door Keypad BOP & Potentiometer Unit", "Specification": "Display BOP + Speed Adjuster Set", "Brand": brand, "Qty": 1, "Unit Price": 3200.0 / ex_rate},
            {"Item": "Panel Cooling Louver Fan Unit", "Specification": "Filter Fan Unit (> 250 m³/h)", "Brand": "Generic", "Qty": 1, "Unit Price": 2800.0 / ex_rate}
        ]
    return []

def get_short_comp_name(full_name):
    """Generates a clean short display name for CAD boxes."""
    fn = str(full_name)
    if "MCCB" in fn: return re.search(r'\d+A', fn).group(0) + " MCCB" if re.search(r'\d+A', fn) else "MCCB"
    if "ACB" in fn: return re.search(r'\d+A', fn).group(0) + " ACB" if re.search(r'\d+A', fn) else "ACB"
    if "VFD Power Unit" in fn or "VFD Drive" in fn: return "VFD " + (re.search(r'\d+(\.\d+)?\s*kW', fn).group(0) if re.search(r'\d+(\.\d+)?\s*kW', fn) else "Drive")
    if "Contactor" in fn: return "Contactor " + (re.search(r'\d+A', fn).group(0) if re.search(r'\d+A', fn) else "")
    if "Timer" in fn: return "SD Timer"
    if "Thermal Overload" in fn or "TOR" in fn: return "TOR Relay"
    if "CT Set" in fn: return "CT Set"
    if "MCB" in fn: return "Control MCB"
    if "Fan" in fn: return "Filter Fan"
    if "Heater" in fn: return "Heater"
    if "SMPS" in fn: return "24V SMPS"
    if "Transformer" in fn: return "Control Tx"
    if "Reactor" in fn or "Choke" in fn: return "AC Line Choke"
    if "Enclosure" in fn or "Frame" in fn: return "Enclosure"
    if "Busbar" in fn or "Harness" in fn: return "Busbars"
    return fn[:14]

def auto_layout_panel_components(bom_df, panel_h_mm=2000, bay_w_mm=800, busbar_pos="Top", cable_alley_w_mm=200):
    """Calculates spatial coordinates (X, Y, W, H) in millimeters for each component with multi-bay auto-overflow."""
    busbar_h_mm = 220
    gland_h_mm = 120
    usable_h_mm = panel_h_mm - busbar_h_mm - gland_h_mm
    
    bays_data = []
    current_bay = 1
    start_y = busbar_h_mm + 30 if busbar_pos == "Top" else 40
    
    cur_x = 40
    cur_y = start_y
    row_max_h = 0
    
    components_placed = []
    
    for _, row in bom_df.iterrows():
        item_name = str(row["Item"])
        qty = int(row["Qty"])
        
        # Exclude structure-level frame/busbars from din-rail mounting plate coordinates
        if any(w in item_name.lower() for w in ["frame", "enclosure", "busbar & power", "wiring harness", "transit packing"]):
            continue
            
        dim = None
        for k, v in COMPONENT_DIMENSIONS_LIBRARY.items():
            if k.lower() in item_name.lower():
                dim = v
                break
        if not dim:
            dim = {"Height_mm": 110, "Width_mm": 90, "Depth_mm": 80, "Type": "General"}
            
        comp_h = dim["Height_mm"]
        comp_w = dim["Width_mm"]
        comp_type = dim["Type"]
        
        for q in range(qty):
            if cur_x + comp_w + 20 > (bay_w_mm - 40):
                cur_x = 40
                cur_y += row_max_h + 30
                row_max_h = 0
                
            max_y_limit = (panel_h_mm - gland_h_mm - 40) if busbar_pos == "Top" else (panel_h_mm - busbar_h_mm - gland_h_mm - 40)
            if cur_y + comp_h > max_y_limit:
                current_bay += 1
                cur_y = start_y
                cur_x = 40
                row_max_h = 0
                
            components_placed.append({
                "Bay": current_bay,
                "Item Name": item_name,
                "Short Name": get_short_comp_name(item_name),
                "Type": comp_type,
                "X (mm)": float(cur_x),
                "Y (mm)": float(cur_y),
                "Width (mm)": float(comp_w),
                "Height (mm)": float(comp_h)
            })
            
            cur_x += comp_w + 20
            if comp_h > row_max_h:
                row_max_h = comp_h
                
    return pd.DataFrame(components_placed)

def generate_internal_ga_svg_from_df(layout_df, panel_h_mm=2000, bay_w_mm=800, busbar_pos="Top", cable_alley_w=200):
    num_bays = int(layout_df["Bay"].max()) if not layout_df.empty else 1
    total_w_mm = (num_bays * bay_w_mm) + cable_alley_w
    
    svg = [f'<svg width="100%" height="auto" viewBox="0 0 {total_w_mm + 80} {panel_h_mm + 120}" xmlns="http://www.w3.org/2000/svg" style="background:#090d16; border-radius:8px; font-family:sans-serif;">']
    svg.append(f'<text x="{(total_w_mm+80)/2}" y="35" fill="#38bdf8" font-size="22" font-weight="bold" text-anchor="middle">INTERNAL GENERAL ARRANGEMENT (GA) - {num_bays} BAY PANEL ({panel_h_mm}x{total_w_mm}mm)</text>')
    
    start_x = 40
    start_y = 60
    
    # Cable Alley Rendering
    svg.append(f'<rect x="{start_x}" y="{start_y}" width="{cable_alley_w}" height="{panel_h_mm}" fill="#1e293b" stroke="#64748b" stroke-width="3"/>')
    svg.append(f'<text x="{start_x + cable_alley_w/2}" y="{start_y + panel_h_mm/2}" fill="#94a3b8" font-size="16" font-weight="bold" text-anchor="middle" transform="rotate(-90,{start_x + cable_alley_w/2},{start_y + panel_h_mm/2})">CABLE ALLEY ({cable_alley_w}mm)</text>')
    
    cur_x = start_x + cable_alley_w
    for b_idx in range(1, num_bays + 1):
        svg.append(f'<rect x="{cur_x}" y="{start_y}" width="{bay_w_mm}" height="{panel_h_mm}" fill="#020617" stroke="#38bdf8" stroke-width="4"/>')
        
        # Busbar Chamber
        bb_h = 200
        bb_y = start_y if busbar_pos == "Top" else (start_y + panel_h_mm - bb_h)
        svg.append(f'<rect x="{cur_x+10}" y="{bb_y}" width="{bay_w_mm-20}" height="{bb_h}" fill="#1e293b" stroke="#cbd5e1" stroke-width="2"/>')
        for i, color in enumerate(["#ef4444", "#eab308", "#3b82f6", "#000000"]):
            bar_y = bb_y + 20 + (i * 38)
            svg.append(f'<rect x="{cur_x+25}" y="{bar_y}" width="{bay_w_mm-50}" height="18" fill="{color}" stroke="#ffffff" stroke-width="1"/>')
        svg.append(f'<text x="{cur_x + bay_w_mm/2}" y="{bb_y + bb_h - 15}" fill="#f8fafc" font-size="14" font-weight="bold" text-anchor="middle">MAIN BUSBAR CHAMBER (R-Y-B-N)</text>')
        
        # Component Rendering in this Bay
        bay_comps = layout_df[layout_df["Bay"] == b_idx] if not layout_df.empty else pd.DataFrame()
        for _, c in bay_comps.iterrows():
            cx = cur_x + float(c["X (mm)"])
            cy = start_y + float(c["Y (mm)"])
            cw = float(c["Width (mm)"])
            ch = float(c["Height (mm)"])
            c_name = xml_escape(c["Item Name"])
            c_short = xml_escape(c.get("Short Name", get_short_comp_name(c["Item Name"])))
            c_type = str(c.get("Type", "General"))
            
            c_color = "#2563eb" if c_type == "VFD" else ("#059669" if c_type in ["MCCB","ACB"] else ("#d97706" if c_type == "Contactor" else "#475569"))
            
            svg.append(f'<g><title>{c_name} ({cw:.0f}x{ch:.0f}mm)</title>')
            svg.append(f'<rect x="{cx}" y="{cy}" width="{cw}" height="{ch}" rx="6" fill="{c_color}" stroke="#f8fafc" stroke-width="2"/>')
            svg.append(f'<text x="{cx + cw/2}" y="{cy + ch/2 + 5}" fill="#ffffff" font-size="12" font-weight="bold" text-anchor="middle">{c_short}</text>')
            svg.append('g>')
            
        svg.append(f'<text x="{cur_x + bay_w_mm/2}" y="{start_y + panel_h_mm - 20}" fill="#38bdf8" font-size="16" font-weight="bold" text-anchor="middle">BAY #{b_idx} ({bay_w_mm}mm)</text>')
        cur_x += bay_w_mm

    svg.append('</svg>')
    return "".join(svg)

def generate_outer_ga_svg_from_df(layout_df, panel_h_mm=2000, bay_w_mm=800, brand="Siemens", panel_type="MCC Panel"):
    num_bays = int(layout_df["Bay"].max()) if not layout_df.empty else 1
    total_w_mm = (num_bays * bay_w_mm) + 200
    
    svg = [f'<svg width="100%" height="auto" viewBox="0 0 {total_w_mm + 80} {panel_h_mm + 120}" xmlns="http://www.w3.org/2000/svg" style="background:#0b0f19; border-radius:8px; font-family:sans-serif;">']
    svg.append(f'<text x="{(total_w_mm+80)/2}" y="35" fill="#38bdf8" font-size="22" font-weight="bold" text-anchor="middle">OUTER ELEVATION GA VIEW - {xml_escape(panel_type)} ({num_bays} BAY)</text>')
    
    start_x = 40
    start_y = 60
    alley_w = 200
    
    # Cable Alley Door
    svg.append(f'<rect x="{start_x}" y="{start_y}" width="{alley_w}" height="{panel_h_mm}" fill="#1e293b" stroke="#64748b" stroke-width="3"/>')
    svg.append(f'<circle cx="{start_x + alley_w - 20}" cy="{start_y + panel_h_mm/2}" r="10" fill="#475569" stroke="#cbd5e1" stroke-width="2"/>')
    
    cur_x = start_x + alley_w
    for b_idx in range(1, num_bays + 1):
        svg.append(f'<rect x="{cur_x}" y="{start_y}" width="{bay_w_mm}" height="{panel_h_mm}" fill="#0f172a" stroke="#38bdf8" stroke-width="4"/>')
        svg.append(f'<rect x="{cur_x + 10}" y="{start_y + 80}" width="12" height="35" fill="#64748b"/>')
        svg.append(f'<rect x="{cur_x + 10}" y="{start_y + panel_h_mm - 120}" width="12" height="35" fill="#64748b"/>')
        svg.append(f'<rect x="{cur_x + bay_w_mm - 30}" y="{start_y + panel_h_mm/2 - 25}" width="18" height="50" rx="4" fill="#334155" stroke="#94a3b8" stroke-width="2"/>')
        
        # Metering Door Compartment
        svg.append(f'<rect x="{cur_x + 40}" y="{start_y + 40}" width="{bay_w_mm - 80}" height="140" rx="6" fill="#020617" stroke="#0ea5e9" stroke-width="2"/>')
        svg.append(f'<rect x="{cur_x + 60}" y="{start_y + 60}" width="120" height="90" rx="4" fill="#0f172a" stroke="#38bdf8" stroke-width="2"/>')
        svg.append(f'<text x="{cur_x + 120}" y="{start_y + 112}" fill="#38bdf8" font-family="monospace" font-size="18" font-weight="bold" text-anchor="middle">415.2 V</text>')
        
        # Phase Indicator LEDs
        for idx, color in enumerate(["#ef4444", "#eab308", "#3b82f6"]):
            svg.append(f'<circle cx="{cur_x + bay_w_mm - 80}" cy="{start_y + 70 + (idx*30)}" r="10" fill="{color}"/>')
            
        # Pushbuttons & Control Lamps Section
        svg.append(f'<rect x="{cur_x + 40}" y="{start_y + 210}" width="{bay_w_mm - 80}" height="180" rx="6" fill="#1e293b" stroke="#64748b" stroke-width="2"/>')
        svg.append(f'<circle cx="{cur_x + 100}" cy="{start_y + 270}" r="16" fill="#22c55e"/>')
        svg.append(f'<circle cx="{cur_x + 180}" cy="{start_y + 270}" r="16" fill="#ef4444"/>')
        svg.append(f'<circle cx="{cur_x + 260}" cy="{start_y + 270}" r="16" fill="#eab308"/>')
        
        # Manufacturer Plate
        svg.append(f'<rect x="{cur_x + bay_w_mm/2 - 120}" y="{start_y + panel_h_mm - 50}" width="240" height="35" rx="4" fill="#f8fafc" stroke="#475569"/>')
        svg.append(f'<text x="{cur_x + bay_w_mm/2}" y="{start_y + panel_h_mm - 28}" fill="#0f172a" font-size="13" font-weight="bold" text-anchor="middle">ARYAVARTA ({xml_escape(brand)})</text>')
        cur_x += bay_w_mm

    svg.append('</svg>')
    return "".join(svg)

def generate_detailed_sld_svg(incomer_info="630A 3P MCCB", feeder_items=None, brand="Siemens"):
    svg_w = 600
    svg_h = 360
    svg = [f'<svg width="100%" height="auto" viewBox="0 0 {svg_w} {svg_h}" xmlns="http://www.w3.org/2000/svg" style="background:#0f172a; border-radius:8px; font-family:sans-serif;">']
    svg.append(f'<text x="{svg_w/2}" y="30" fill="#38bdf8" font-size="16" font-weight="bold" text-anchor="middle">POWER SINGLE LINE DIAGRAM (SLD)</text>')
    
    svg.append('<line x1="40" y1="60" x2="560" y2="60" stroke="#ef4444" stroke-width="4"/>')
    svg.append('<line x1="40" y1="68" x2="560" y2="68" stroke="#eab308" stroke-width="4"/>')
    svg.append('<line x1="40" y1="76" x2="560" y2="76" stroke="#3b82f6" stroke-width="4"/>')
    svg.append('<text x="300" y="50" fill="#94a3b8" font-family="monospace" font-size="11" text-anchor="middle">415V 3-PHASE 50Hz MAIN BUSBAR</text>')
    
    svg.append('<line x1="300" y1="76" x2="300" y2="110" stroke="#cbd5e1" stroke-width="3"/>')
    svg.append('<rect x="230" y="110" width="140" height="45" rx="5" fill="#1e293b" stroke="#38bdf8" stroke-width="2"/>')
    svg.append(f'<text x="300" y="137" fill="#38bdf8" font-size="12" font-weight="bold" text-anchor="middle">{xml_escape(str(incomer_info)[:18])}</text>')
    
    svg.append('<line x1="300" y1="155" x2="300" y2="200" stroke="#cbd5e1" stroke-width="3"/>')
    svg.append('<circle cx="300" cy="178" r="14" fill="none" stroke="#eab308" stroke-width="2"/>')
    svg.append('<text x="328" y="182" fill="#eab308" font-size="10">CT 600/5A</text>')
    
    feeder_x_positions = [130, 300, 470]
    f_titles = ["VFD Drive Feeder", "Star-Delta Starter", "Auxiliary DOL Feeder"]
    
    svg.append('<line x1="130" y1="200" x2="470" y2="200" stroke="#cbd5e1" stroke-width="3"/>')
    
    for idx, fx in enumerate(feeder_x_positions):
        svg.append(f'<line x1="{fx}" y1="200" x2="{fx}" y2="235" stroke="#cbd5e1" stroke-width="3"/>')
        svg.append(f'<rect x="{fx-55}" y="235" width="110" height="45" rx="4" fill="#020617" stroke="#22c55e" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="262" fill="#22c55e" font-size="10" font-weight="bold" text-anchor="middle">{xml_escape(f_titles[idx])}</text>')
        svg.append(f'<line x1="{fx}" y1="280" x2="{fx}" y2="305" stroke="#cbd5e1" stroke-width="3"/>')
        svg.append(f'<circle cx="{fx}" cy="322" r="16" fill="#1e293b" stroke="#e2e8f0" stroke-width="2"/>')
        svg.append(f'<text x="{fx}" y="327" fill="#ffffff" font-size="11" font-weight="bold" text-anchor="middle">M 3~</text>')
        
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

def query_local_ollama(prompt, format_json=True):
    payload = {"model": "llama3.2", "prompt": prompt, "stream": False}
    if format_json: payload["format"] = "json"
    try:
        req = urllib.request.Request("http://localhost:11434/api/generate", data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=12) as res:
            resp = json.loads(res.read().decode("utf-8"))["response"]
            return json.loads(resp) if format_json else resp
    except Exception: return None

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

def parse_inquiry_text_heuristically(text):
    text_lower = text.lower()
    kw_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:kw|hp)', text_lower)
    found_kw = None
    if kw_match:
        val = float(kw_match.group(1))
        if "hp" in kw_match.group(0): val = val * 0.7457
        std_kws = [7.5, 15, 22, 37, 55, 75]
        closest = min(std_kws, key=lambda x: abs(x - val))
        found_kw = f"{closest if isinstance(closest, int) or closest.is_integer() else closest} kW".replace('.0 kW', ' kW')

    found_panel = None
    if "apfc" in text_lower or "power factor" in text_lower: found_panel = "APFC Panel"
    elif "star" in text_lower or "delta" in text_lower: found_panel = "Star-Delta Control Panel"
    elif "lt" in text_lower or "distribution" in text_lower or "pcc" in text_lower: found_panel = "LT Distribution Panel"
    elif "vfd" in text_lower or "drive" in text_lower or "inverter" in text_lower: found_panel = "VFD Panel"

    found_brand = None
    for b in ["Siemens", "Schneider", "L&T", "ABB", "Danfoss", "Delta"]:
        if b.lower() in text_lower:
            found_brand = b
            break

    return {"motor_kw": found_kw, "panel_type": found_panel, "preferred_brand": found_brand}

st.sidebar.title("⚡ Aryavarta Automation")
st.sidebar.markdown("**Sales & Engineering Suite**")
st.sidebar.success("🟢 100% Offline & Free | Chikhali Plant")

menu = st.sidebar.radio("Navigation", [
    "Create Panel Quote", 
    "MCC Multi-Feeder Panel Sizer",
    "Motor Switchgear Master Chart",
    "MPCB & Motor Protection Sizer",
    "Quotation Register & Pipeline",
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
    
    with st.expander("🛠️ Interactive Panel Sizing & Enclosure Bay Configurator", expanded=True):
        ga_c1, ga_c2, ga_c3, ga_c4 = st.columns(4)
        panel_h_input = ga_c1.number_input("Panel Frame Height (mm)", value=2000, step=100)
        bay_w_input = ga_c2.number_input("Bay Width (mm)", value=800, step=100)
        panel_d_input = ga_c3.number_input("Panel Depth (mm)", value=600, step=50)
        cable_alley_w_input = ga_c4.number_input("Cable Alley Width (mm)", value=200, step=50)
        busbar_pos_input = st.radio("Main Busbar Chamber Position", ["Top", "Bottom"], horizontal=True)

    # Initialize or reset component layout in session state
    if "panel_layout_df" not in st.session_state or st.button("🔄 Reset Component Coordinates to Default Auto-Layout"):
        st.session_state["panel_layout_df"] = auto_layout_panel_components(
            bom_df, panel_h_input, bay_w_input, busbar_pos_input, cable_alley_w_input
        )

    st.markdown("#### 🎛️ Interactive Component Coordinates & Placement Editor")
    st.info("💡 **Live Interactive CAD:** Edit **Bay**, **X (mm)**, **Y (mm)**, **Width (mm)**, **Height (mm)**, or **Item Name** directly in the table below to reposition components inside the panel in real time!")
    
    edited_layout_df = st.data_editor(
        st.session_state["panel_layout_df"],
        num_rows="dynamic",
        key="cad_layout_editor",
        column_config={
            "Bay": st.column_config.NumberColumn("Bay #", min_value=1, max_value=10, step=1),
            "X (mm)": st.column_config.NumberColumn("Pos X (mm)", step=10.0),
            "Y (mm)": st.column_config.NumberColumn("Pos Y (mm)", step=10.0),
            "Width (mm)": st.column_config.NumberColumn("Width (mm)", step=5.0),
            "Height (mm)": st.column_config.NumberColumn("Height (mm)", step=5.0)
        }
    )
    
    st.session_state["panel_layout_df"] = edited_layout_df

    tab_int_ga, tab_outer_ga, tab_sld, tab_spatial = st.tabs([
        "🖼️ Internal GA (Interior Mounting)", 
        "🚪 Outer GA (Front Door & Meters)", 
        "⚡ Power Single Line Diagram (SLD)",
        "📐 Spatial Footprint & Bay Analysis"
    ])

    with tab_int_ga:
        int_ga_svg = generate_internal_ga_svg_from_df(
            st.session_state["panel_layout_df"], panel_h_input, bay_w_input, busbar_pos_input, cable_alley_w_input
        )
        st.markdown(int_ga_svg, unsafe_allow_html=True)
        st.download_button(
            "📥 Download Internal GA (.SVG)", 
            int_ga_svg, 
            f"Internal_GA_{client_name.replace(' ', '_')}.svg", 
            mime="image/svg+xml"
        )

    with tab_outer_ga:
        outer_ga_svg = generate_outer_ga_svg_from_df(
            st.session_state["panel_layout_df"], panel_h_input, bay_w_input, preferred_brand, panel_type
        )
        st.markdown(outer_ga_svg, unsafe_allow_html=True)
        st.download_button(
            "📥 Download Outer GA (.SVG)", 
            outer_ga_svg, 
            f"Outer_GA_{client_name.replace(' ', '_')}.svg", 
            mime="image/svg+xml"
        )

    with tab_sld:
        sld_svg = generate_detailed_sld_svg(
            selected_incomer if quote_mode=="⚡ Single Feeder Panel" else incomer_type, bom_df, preferred_brand
        )
        st.markdown(sld_svg, unsafe_allow_html=True)
        st.download_button(
            "📥 Download Power SLD (.SVG)", 
            sld_svg, 
            f"Power_SLD_{client_name.replace(' ', '_')}.svg", 
            mime="image/svg+xml"
        )

    with tab_spatial:
        num_bays_val = int(st.session_state["panel_layout_df"]["Bay"].max()) if not st.session_state["panel_layout_df"].empty else 1
        st.markdown(f"#### 📊 Spatial Footprint Summary: **{num_bays_val} Panel Column / Bay(s) Configured**")
        st.info(f"💡 Total Panel Enclosure Shipping Footprint: **{panel_h_input}mm (H) x {(num_bays_val*bay_w_input)+cable_alley_w_input}mm (W) x {panel_d_input}mm (D)**.")
        st.dataframe(st.session_state["panel_layout_df"])

    if st.button("💾 Save Quote to Quotation Register & Pipeline"):
        save_quote_to_history(client_name, panel_type, motor_kw, preferred_brand, sell_price * ex_rate, margin_pct)
        st.success(f"Quote for {client_name} ({curr_sym} {sell_price:,.2f}) saved to Quotation Register!")

elif menu == "MCC Multi-Feeder Panel Sizer":
    st.header("🏭 MCC Multi-Feeder Load & Incomer Sizer")
    c1, c2, c3 = st.columns(3)
    with c1:
        v_kw = st.number_input("VFD Motor Power (kW)", value=22.0, step=2.5)
        v_q = st.number_input("VFD Feeders Qty", value=2, min_value=0)
    with c2:
        sd_kw_v = st.number_input("Star-Delta Motor Power (kW)", value=37.0, step=2.5)
        sd_q = st.number_input("Star-Delta Feeders Qty", value=2, min_value=0)
    with c3:
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

elif menu == "Motor Switchgear Master Chart":
    st.header("⚡ 3-Phase Motor Switchgear & Cable Master Reference Chart")
    std_ratings = [7.5, 11.0, 15.0, 18.5, 22.0, 30.0, 37.0, 45.0, 55.0, 75.0, 90.0, 110.0, 132.0, 160.0]
    chart_data = [calculate_motor_master_chart(k) for k in std_ratings]
    st.dataframe(pd.DataFrame(chart_data), use_container_width=True)

elif menu == "MPCB & Motor Protection Sizer":
    st.header("🛡️ MPCB & Overload Trip Class Sizer (IEC 60947-4-1)")
    c1, c2 = st.columns(2)
    m_kw_p = c1.number_input("Motor Power Rating (kW)", value=15.0, step=2.5)
    l_type = c2.selectbox("Application / Duty Inertia", [
        "Standard Load (Class 10 - Pumps, Fans, Blowers)",
        "Medium Inertia (Class 20 - Crushers, Agitators)",
        "High Inertia (Class 30 - Centrifuges, Heavy Flywheel Fans)"
    ])
    p_res = calculate_mpcb_and_protection(m_kw_p, l_type)
    st.divider()
    pr1, pr2, pr3, pr4 = st.columns(4)
    pr1.metric("Motor FLC Current", f"{p_res['flc']} A")
    pr2.metric("MPCB Thermal Set Range", p_res["thermal_range"])
    pr3.metric("Magnetic Instantaneous Trip", p_res["mag_trip"])
    pr4.metric("Recommended Protection", p_res["protection_device"])

elif menu == "Quotation Register & Pipeline":
    st.header("📈 Quotation Register & Sales Pipeline Tracker")
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
        st.download_button("📥 Export Full Quotes Register (.XLSX)", reg_buf.getvalue(), f"Quotes_Register_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.subheader("📊 Quotations Log & Status Manager")
        edited_q = st.data_editor(q_df, num_rows="dynamic")
        if st.button("💾 Update Register Status"):
            edited_q.to_excel(QUOTES_FILE, index=False)
            st.success("Quotation Register updated successfully!")

elif menu == "Manage Price Database":
    st.header("⚙️ Local Price Database Manager")
    edited = st.data_editor(df_prices, num_rows="dynamic")
    if st.button("💾 Save Changes"):
        edited.to_excel(DB_FILE, index=False)
        st.cache_data.clear()
        st.success("Database updated successfully!")
