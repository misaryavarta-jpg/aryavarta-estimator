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
    new_data = pd.DataFrame([{
        "Date": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "Client": client, "Scope": f"{panel_type} ({kw})", "Brand": brand,
        "Quote_Price_INR": price, "Margin_Pct": margin, "Status": "Pending"
    }])
    df_updated = pd.concat([pd.read_excel(QUOTES_FILE), new_data], ignore_index=True) if os.path.exists(QUOTES_FILE) else new_data
    df_updated.to_excel(QUOTES_FILE, index=False)

def init_database():
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

def calculate_cu_vs_al_busbar(current_amps, busbar_len_m, cu_rate=850.0, al_rate=260.0):
    area_cu = current_amps / 1.5
    area_al = current_amps / 0.9
    vol_cu = (area_cu / 1e6) * busbar_len_m * 3.45
    vol_al = (area_al / 1e6) * busbar_len_m * 3.45
    wt_cu = round(vol_cu * 8960.0, 2)
    wt_al = round(vol_al * 2700.0, 2)
    cost_cu = round(wt_cu * cu_rate, 2)
    cost_al = round(wt_al * al_rate, 2)
    cost_saved = round(cost_cu - cost_al, 2)
    wt_saved = round(wt_cu - wt_al, 2)
    pct_cost_saved = round((cost_saved / cost_cu) * 100, 1) if cost_cu > 0 else 0
    return {
        "area_cu": round(area_cu, 1), "area_al": round(area_al, 1),
        "wt_cu": wt_cu, "wt_al": wt_al,
        "cost_cu": cost_cu, "cost_al": cost_al,
        "cost_saved": cost_saved, "wt_saved": wt_saved, "pct_cost_saved": pct_cost_saved
    }

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

def calculate_acb_and_protection_settings(kva_rating, pf=0.85, required_icu=50, execution="Drawout"):
    i_nom = (kva_rating * 1000.0) / (1.732 * 415.0)
    acb_ratings = [800, 1000, 1250, 1600, 2000, 2500, 3200, 4000, 5000, 6300]
    rec_rating = next((r for r in acb_ratings if r >= i_nom * 1.25), acb_ratings[-1])
    frame = "Frame 1 (up to 1600A)" if rec_rating <= 1600 else ("Frame 2 (2000A - 3200A)" if rec_rating <= 3200 else "Frame 3 (4000A - 6300A)")
    ir_set = round(i_nom, 1)
    isd_set = round(ir_set * 4.0, 1)
    ii_set = round(rec_rating * 10.0, 1)
    ig_set = round(rec_rating * 0.3, 1)
    return {
        "i_nom": round(i_nom, 1), "rec_acb": f"{rec_rating}A {execution} ACB ({required_icu} kA)",
        "frame": frame, "ir_setting": f"Ir = {round(ir_set/rec_rating, 2)}x In ({ir_set} A, tr = 10s)",
        "isd_setting": f"Isd = 4.0x Ir ({isd_set} A, tsd = 0.2s)",
        "ii_setting": f"Ii = 10x In ({ii_set} A, Instantaneous)",
        "ig_setting": f"Ig = 0.3x In ({ig_set} A, tg = 0.2s)"
    }

def calculate_motor_starting_vdrop(kw_val, starting_method, cable_len_m, cable_sqmm, conductor_type="Copper"):
    flc = (kw_val * 1000.0) / (1.732 * 415.0 * 0.85 * 0.88)
    mult = {"DOL": 6.0, "Star-Delta": 2.5, "Soft Starter": 2.0, "VFD": 1.15}.get(starting_method, 6.0)
    i_start = flc * mult
    rho = 0.0175 if conductor_type == "Copper" else 0.0282
    r_per_m = rho / cable_sqmm
    x_per_m = 0.00008
    pf_start = 0.35
    sin_phi = math.sqrt(1 - pf_start**2)
    v_drop_volts = 1.732 * i_start * cable_len_m * (r_per_m * pf_start + x_per_m * sin_phi)
    v_drop_pct = round((v_drop_volts / 415.0) * 100.0, 2)
    term_voltage = round(415.0 - v_drop_volts, 1)
    status = "PASS ✅ (Dip < 10% - Standard)" if v_drop_pct <= 10.0 else ("WARNING ⚠️ (10% - 15% Dip - Check contactor hold-in)" if v_drop_pct <= 15.0 else "FAIL ❌ (>15% Dip - High risk of motor stalling/chatter)")
    return {"flc": round(flc, 1), "i_start": round(i_start, 1), "v_drop_v": round(v_drop_volts, 1), "v_drop_pct": v_drop_pct, "term_v": term_voltage, "status": status}

def calculate_vfd_payback(motor_kw, hrs_per_day, rate_per_kwh, speed_reduction_pct, vfd_cost_inr):
    speed_ratio = (100.0 - speed_reduction_pct) / 100.0
    power_vfd_kw = motor_kw * (speed_ratio ** 3)
    kw_saved_per_hr = motor_kw - power_vfd_kw
    annual_kwh_saved = kw_saved_per_hr * hrs_per_day * 365.0
    annual_savings_inr = annual_kwh_saved * rate_per_kwh
    payback_months = (vfd_cost_inr / annual_savings_inr * 12.0) if annual_savings_inr > 0 else 0
    return {"kw_saved_hr": round(kw_saved_per_hr, 2), "annual_kwh": round(annual_kwh_saved, 0), "annual_inr": round(annual_savings_inr, 2), "payback_months": round(payback_months, 1)}

def calculate_dbr_sizing(motor_kw, load_type="Standard Machine (10% Duty)", dc_bus_v=750):
    duty_map = {
        "Light Stopping (5% Duty)": (0.05, 1.0),
        "Standard Machine (10% Duty)": (0.10, 1.15),
        "High Inertia Centrifuge / Fan (20% Duty)": (0.20, 1.25),
        "Heavy Crane / Hoist Overhauling Load (50% Duty)": (0.50, 1.50)
    }
    duty_cycle, power_mult = duty_map.get(load_type, (0.10, 1.15))
    peak_p_kw = motor_kw * power_mult
    res_ohms = round((dc_bus_v ** 2) / (peak_p_kw * 1000.0), 1)
    cont_watts = round(peak_p_kw * 1000.0 * duty_cycle, 0)
    rec_type = "Aluminum Encased Wirewound Resistor" if cont_watts <= 1200 else "Stainless Steel Grid Resistor Bank"
    return {"peak_kw": round(peak_p_kw, 1), "res_ohms": res_ohms, "cont_watts": cont_watts, "duty_pct": f"{int(duty_cycle * 100)}%", "rec_type": rec_type}

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

def calculate_short_circuit_and_cu_weight(tx_kva, tx_z_pct, bus_w_mm, bus_t_mm, bus_len_m, cu_rate_per_kg):
    i_sc_ka = (tx_kva * 100.0) / (math.sqrt(3) * 0.415 * tx_z_pct * 10)
    rec_ka = 10 if i_sc_ka <= 10 else 16 if i_sc_ka <= 16 else 25 if i_sc_ka <= 25 else 36 if i_sc_ka <= 36 else 50
    volume_m3 = (bus_w_mm / 1000.0) * (bus_t_mm / 1000.0) * bus_len_m * 3.45
    cu_weight_kg = volume_m3 * 8960.0
    actual_area = bus_w_mm * bus_t_mm
    req_area_1s = (i_sc_ka * 1000.0 * 1.0) / 176.0
    thermal_status = "PASS ✅" if actual_area >= req_area_1s else "FAIL ❌ (Under-sized for 1s Short Circuit)"
    req_gi_earth = (i_sc_ka * 1000.0 * 1.0) / 80.0
    gi_strips = [(25,3), (25,6), (32,6), (50,6), (50,10), (75,10)]
    rec_gi_earth = next((f"{w}x{t} mm GI Earth Strip" for w, t in gi_strips if (w*t) >= req_gi_earth), "75x10 mm GI Earth Strip")
    
    i_pk_ka = i_sc_ka * 2.1
    d_m = max(0.05, (bus_w_mm + 25) / 1000.0)
    f_peak_nm = round(0.2 * (i_pk_ka ** 2) / d_m, 1)
    spacing_mm = 250 if i_sc_ka > 36 else (300 if i_sc_ka > 25 else (400 if i_sc_ka > 16 else 500))
    
    return {"i_sc_ka": round(i_sc_ka, 2), "i_pk_ka": round(i_pk_ka, 1), "f_peak_nm": f_peak_nm, "rec_spacing_mm": f"{spacing_mm} mm", "rec_mccb_ka": f"{rec_ka} kA Rating", "cu_weight_kg": round(cu_weight_kg, 2), "cu_cost_inr": round(cu_weight_kg * cu_rate_per_kg, 2), "actual_area": actual_area, "req_area_1s": round(req_area_1s, 1), "thermal_status": thermal_status, "rec_gi_earth": rec_gi_earth}

def calculate_panel_thermal(h_mm, w_mm, d_mm, vfd_kw, vfd_qty, amb_t, max_int_t, sealed, outdoor=False):
    area_m2 = 1.8 * (h_mm/1000) * ((w_mm+d_mm)/1000) + 1.4 * (w_mm/1000) * (d_mm/1000)
    solar_w = (area_m2 * 125.0) if outdoor else 0.0
    tot_w = (vfd_kw * 30 * vfd_qty) + 150 + (vfd_kw * 10 * vfd_qty) + solar_w
    dt = max(5.0, max_int_t - amb_t)
    nat_w = 5.5 * area_m2 * dt
    net_w = max(0.0, tot_w - nat_w)
    m3h = (3.3 * net_w) / dt if net_w > 0 else 0
    heater_w = 40 if area_m2 <= 2.5 else (80 if area_m2 <= 5.0 else 150)
    rec = "Natural radiation sufficient." if net_w <= 0 else (f"⚠️ Sealed Panel AC Required: {round(net_w/1000, 1)} kW" if sealed else f"Install Filter Fan Unit (> {round(m3h)} m³/h)")
    return {"tot_w": round(tot_w,1), "nat_w": round(nat_w,1), "net_w": round(net_w,1), "m3h": round(m3h,1), "rec": rec, "heater_w": f"{heater_w}W + Thermostat", "solar_w": round(solar_w, 1)}

def calculate_dg_and_transformer_sizing(running_kw, max_start_kw, start_method, diversity):
    pf = 0.85
    base_running_kva = (running_kw * diversity) / pf
    start_mult = {"DOL": 6.0, "Star-Delta": 3.0, "Soft Starter": 2.5, "VFD": 1.15}.get(start_method, 1.15)
    peak_kva = base_running_kva + (((max_start_kw / pf) * start_mult) * 0.4)
    tx_ratings = [100, 160, 250, 315, 500, 630, 750, 1000, 1250, 1600, 2000, 2500]
    dg_ratings = [62.5, 82.5, 125, 160, 200, 250, 320, 500, 625, 750, 1010, 1250]
    return {"running_kva": round(base_running_kva, 1), "peak_kva": round(peak_kva, 1), "rec_tx": f"{next((r for r in tx_ratings if r >= peak_kva * 1.20), tx_ratings[-1])} kVA", "rec_dg": f"{next((r for r in dg_ratings if r >= peak_kva * 1.25), dg_ratings[-1])} kVA"}

def calculate_apfc_and_msedcl_rebate(active_kw, c_pf, t_pf, harmonics, monthly_bill_inr, panel_cost_inr):
    if c_pf >= t_pf: return {"req_kvar": 0, "steps": "None", "monthly_savings": 0, "payback_months": 0, "status": "Current PF is optimal."}
    kvar = math.ceil(active_kw * (math.tan(math.acos(c_pf)) - math.tan(math.acos(t_pf))))
    steps = [5,5,5,10] if kvar<=25 else [5,10,10,25] if kvar<=50 else [10,15,25,25,25] if kvar<=100 else [25,25,50,50] if kvar<=200 else [50]*math.ceil(kvar/50)
    current_rebate_pct = 0.0 if c_pf >= 0.95 else -0.03 if c_pf < 0.90 else 0.0
    target_rebate_pct = 0.07 if t_pf >= 0.98 else 0.05 if t_pf >= 0.95 else 0.0
    net_rebate_pct = target_rebate_pct - current_rebate_pct
    monthly_savings = monthly_bill_inr * net_rebate_pct
    payback_m = (panel_cost_inr / monthly_savings) if monthly_savings > 0 else 0
    return {"req_kvar": kvar, "steps": f"{len(steps)} Steps: " + ", ".join([f"{s}kVAR" for s in steps]), "monthly_savings": round(monthly_savings, 2), "payback_months": round(payback_m, 1), "pct": round((1 - c_pf/t_pf)*100, 1)}

def calculate_detuned_reactor_and_capacitor(req_kvar, sys_voltage=415.0, detuning_pct=7.0):
    p = detuning_pct / 100.0
    v_cap_operating = sys_voltage / (1.0 - p)
    cap_ratings = [440, 480, 525, 690]
    rec_v_rating = next((v for v in cap_ratings if v >= v_cap_operating * 1.05), cap_ratings[-1])
    nameplate_kvar = req_kvar * ((rec_v_rating / v_cap_operating) ** 2)
    reactor_kvar = req_kvar * p
    tuning_freq = round(50.0 / math.sqrt(p), 1)
    
    step_kvar = min(req_kvar, 25.0) if req_kvar > 0 else 25.0
    c_phase = (step_kvar * 1000.0) / (3.0 * 2.0 * math.pi * 50.0 * (sys_voltage ** 2))
    r_dis_kohms = round((50.0 / (c_phase * 2.463)) / 1000.0, 1) if c_phase > 0 else 120.0
    p_dis_w = max(5.0, round(((sys_voltage ** 2) / (r_dis_kohms * 1000.0)) * 2.5, 1)) if r_dis_kohms > 0 else 5.0

    return {
        "v_cap_op": round(v_cap_operating, 1),
        "rec_v_rating": f"{rec_v_rating}V Heavy Duty Capacitor",
        "nameplate_kvar": round(nameplate_kvar, 1),
        "reactor_kvar": round(reactor_kvar, 1),
        "tuning_freq": f"{tuning_freq} Hz ({int(detuning_pct)}% Detuned)",
        "r_dis_kohms": f"{r_dis_kohms} kΩ",
        "p_dis_w": f"{p_dis_w}W Wirewound"
    }

def calculate_harmonics_and_ahf(vfd_kw, vfd_qty, has_choke):
    vfd_current = ((vfd_kw * 1000) / (1.732 * 415 * 0.85 * 0.95)) * vfd_qty
    thdi_pct = 30.0 if has_choke else 42.0
    harmonic_amps = vfd_current * (thdi_pct / 100.0)
    req_ahf_amps = max(0.0, vfd_current * ((thdi_pct - 8.0) / 100.0))
    standard_ahf = [30, 50, 75, 100, 150, 200, 300]
    rec_ahf = next((f"{a}A Active Harmonic Filter (AHF)" for a in standard_ahf if a >= req_ahf_amps), f"{round(req_ahf_amps)}A AHF System") if req_ahf_amps > 15 else "No AHF required (3% Line Reactor sufficient)."
    return {"vfd_current": round(vfd_current, 1), "thdi_pct": thdi_pct, "harmonic_amps": round(harmonic_amps, 1), "req_ahf_amps": round(req_ahf_amps, 1), "rec_ahf": rec_ahf}

def calculate_emc_and_dvdt_filter(vfd_kw, cable_len_m, environment="Industrial (Class C3)"):
    flc = (vfd_kw * 1000.0) / (1.732 * 415.0 * 0.85 * 0.88)
    emc_class = "Class C3 (Industrial Environment)" if "Industrial" in environment else "Class C2 (Commercial/Public)"
    dvdt_rec = "None Required (Safe < 30m)" if cable_len_m <= 30 else ("3% Output dv/dt AC Reactor" if cable_len_m <= 100 else "Sine-Wave Filter (LC Filter)")
    return {
        "rec_emc_filter": f"{math.ceil(flc * 1.25)}A RFI/EMC Filter ({emc_class})",
        "dvdt_rec": dvdt_rec,
        "cable_spec": "Shielded / Armored Cable (360° EMC Glands)"
    }

def calculate_smps_and_control_tx(contactor_qty, relay_qty, hmi_present, heater_w, sensor_qty):
    ac_va = (contactor_qty * 15) + (relay_qty * 3) + heater_w + 30
    tx_ratings = [100, 150, 250, 500, 750, 1000, 1500, 2000]
    dc_amps = 0.5 + (0.8 if hmi_present else 0) + (relay_qty * 0.05) + (sensor_qty * 0.03)
    smps_ratings = [2.5, 5.0, 10.0, 20.0, 40.0]
    return {"ac_va": round(ac_va, 1), "rec_tx": f"{next((r for r in tx_ratings if r >= ac_va * 1.3), tx_ratings[-1])} VA Control Tx", "dc_amps": round(dc_amps, 2), "rec_smps": f"{next((s for s in smps_ratings if s >= dc_amps * 1.25), smps_ratings[-1])}A / 24V DC SMPS"}

def calculate_dc_control_vdrop(dc_amps, wire_sqmm, wire_len_m):
    rho = 0.0178
    v_drop_dc = round((2.0 * dc_amps * wire_len_m * rho) / wire_sqmm, 2)
    end_v = round(24.0 - v_drop_dc, 2)
    max_dist_m = round((3.6 * wire_sqmm) / (2.0 * dc_amps * rho), 1) if dc_amps > 0 else 0
    status = "PASS ✅ Safe (> 20.4V DC)" if end_v >= 20.4 else "FAIL ❌ Drop > 15% (Risk of Relay Chatter / PLC I/O Drop)"
    return {"v_drop_dc": v_drop_dc, "end_v": end_v, "max_dist_m": max_dist_m, "status": status}

def calculate_ups_and_battery(load_va, backup_mins, dc_bus_v=24):
    ups_ratings = [0.5, 1.0, 2.0, 3.0, 5.0, 6.0, 10.0]
    req_ups = next((r for r in ups_ratings if r >= (load_va * 1.25) / 1000.0), ups_ratings[-1])
    req_ah = (load_va * (backup_mins / 60.0)) / (dc_bus_v * 0.85)
    ah_std = [7, 12, 18, 26, 42, 65, 100, 120, 150, 200]
    rec_ah = next((a for a in ah_std if a >= req_ah), ah_std[-1])
    batt_qty = max(1, int(dc_bus_v / 12))
    return {"rec_ups": f"{req_ups} kVA Online UPS", "req_ah": round(req_ah, 1), "rec_battery": f"{batt_qty}x 12V {rec_ah} Ah SMF Batteries ({dc_bus_v}V DC Bus)", "chg_amps": round(rec_ah * 0.1, 1)}

def calculate_plc_hmi_sizing(di, do, ai, ao, hmi_size, plc_brand):
    extra_di = max(0, di - 8); extra_do = max(0, do - 6)
    d_cards = math.ceil(max(extra_di, extra_do) / 8.0)
    ai_cards = math.ceil(ai / 4.0); ao_cards = math.ceil(ao / 2.0)
    brand_mult = {"Siemens": 1.5, "Schneider": 1.3, "Delta": 0.85}.get(plc_brand, 1.0)
    cpu_cost = 18000 * brand_mult; card_cost = (d_cards * 6500 + ai_cards * 9500 + ao_cards * 8500) * brand_mult
    hmi_cost = {"4.3 Inch": 12000, "7.0 Inch": 22000, "10.0 Inch": 42000, "None": 0}.get(hmi_size, 0)
    total_hw = cpu_cost + card_cost + hmi_cost
    
    card_list = []
    if d_cards > 0: card_list.append(f"{d_cards}x 8-Ch DI/DO")
    if ai_cards > 0: card_list.append(f"{ai_cards}x 4-Ch AI")
    if ao_cards > 0: card_list.append(f"{ao_cards}x 2-Ch AO")
    cards_desc = ", ".join(card_list) if card_list else "None (Embedded I/O Sufficient)"
    
    return {
        "tot_io": di + do + ai + ao,
        "cpu": f"{plc_brand} Base CPU (8 DI / 6 DO)",
        "cards": cards_desc,
        "hmi": f"{hmi_size} Touchscreen" if hmi_size != "None" else "No HMI",
        "est_hw_cost": round(total_hw, 2)
    }

def calculate_panel_shipping_weight(h_mm, w_mm, d_mm, kw_str):
    enc_w = calculate_enclosure_fabrication(h_mm, w_mm, d_mm, 1.6, 82, 220)["weight_kg"]
    kw_val = float(kw_str.split()[0]) if kw_str and kw_str.split()[0].replace('.','',1).isdigit() else 15.0
    sw_w = 8.0 + (kw_val * 0.55)
    cu_w = 4.0 + (kw_val * 0.25)
    tot_w = round(enc_w + sw_w + cu_w, 1)
    vol_m3 = round(((h_mm + 100) * (w_mm + 100) * (d_mm + 100)) / 1_000_000_000.0, 2)
    return {"gross_weight_kg": tot_w, "volume_m3": vol_m3}

def calculate_gland_and_terminals(cable_sqmm, cores, control_signals):
    gland_map = [
        (4, "M20 (Single/Double Comp)"),
        (10, "M25 (Single/Double Comp)"),
        (25, "M32 (Double Compression)"),
        (50, "M40 (Double Compression)"),
        (95, "M50 (Double Compression)"),
        (185, "M63 (Double Compression)"),
        (300, "M75 (Double Compression)")
    ]
    gland_size = next((g for s, g in gland_map if cable_sqmm <= s), "M75 (Double Compression)")
    p_tb_type = f"{cable_sqmm} sq.mm Din-Rail / Stud TB ({cores}P)"
    ctrl_tbs = math.ceil(control_signals * 1.2)
    
    p_od = 12 if cable_sqmm <= 10 else (18 if cable_sqmm <= 35 else (26 if cable_sqmm <= 95 else (38 if cable_sqmm <= 185 else 48)))
    tray_req_mm = (2 * p_od + (control_signals / 2) * 8) * 1.35
    trays = [100, 150, 300, 450, 600, 750, 900]
    rec_tray = next((f"{t} mm Perforated/Ladder Tray" for t in trays if t >= tray_req_mm * 1.25), "900 mm Heavy Duty Tray")
    
    return {"gland": gland_size, "p_tb": p_tb_type, "ctrl_tbs": f"{ctrl_tbs} Nos (Includes 20% Spares)", "rec_tray": rec_tray}

def get_control_wiring_and_ferrule_schedule(scheme_type):
    colors_info = [
        {"Circuit Type": "3-Phase Power", "Standard Color": "Red, Yellow, Blue", "Wire Size": "2.5 to 300+ sq.mm", "Ferrule Scheme": "R, Y, B / L1, L2, L3"},
        {"Circuit Type": "Neutral Conductor", "Standard Color": "Black", "Wire Size": "Phase size or 50% reduced", "Ferrule Scheme": "N / N1"},
        {"Circuit Type": "Protective Earth (PE)", "Standard Color": "Green / Green-Yellow", "Wire Size": "Min 2.5 sq.mm / GI Strip", "Ferrule Scheme": "PE / E"},
        {"Circuit Type": "230V AC Control", "Standard Color": "Red or Grey (IS 375)", "Wire Size": "1.0 sq.mm / 1.5 sq.mm", "Ferrule Scheme": "C01, C02, L, N, A1, A2"},
        {"Circuit Type": "24V DC Positive (+)", "Standard Color": "Dark Blue", "Wire Size": "0.75 sq.mm / 1.0 sq.mm", "Ferrule Scheme": "+24V, P24"},
        {"Circuit Type": "24V DC Negative (0V)", "Standard Color": "Blue-White / Light Blue", "Wire Size": "0.75 sq.mm / 1.0 sq.mm", "Ferrule Scheme": "0V, N24, GND"},
        {"Circuit Type": "Analog (4-20mA/0-10V)", "Standard Color": "Shielded Twisted Pair", "Wire Size": "0.5 sq.mm / 0.75 sq.mm", "Ferrule Scheme": "AI1+, AI1-, AO1+, AO1-"},
        {"Circuit Type": "External Interlock", "Standard Color": "Orange (IS 375)", "Wire Size": "1.0 sq.mm", "Ferrule Scheme": "X01, X02 (Live when Main OFF)"}
    ]
    ferrule_sample = {
        "VFD Panel": [("Incomer Power", "L1, L2, L3"), ("VFD Output Motor", "U, V, W"), ("Run/Stop Command", "DI1, COM"), ("Speed Potentiometer", "+10V, AI1, GND"), ("Fault Relay Contact", "R1A, R1B, R1C")],
        "Star-Delta Panel": [("Main Contactor Coil", "A1-M, A2-M"), ("Delta Contactor Coil", "A1-D, A2-D"), ("Star Contactor Coil", "A1-S, A2-S"), ("Timer Control", "15, 18, 28"), ("Thermal Overload", "95, 96, 97, 98")],
        "APFC Panel": [("CT Secondary Current", "S1-R, S2-R, S1-Y, S2-Y, S1-B, S2-B"), ("Capacitor Steps 1-6", "C1, C2, C3, C4, C5, C6"), ("Controller Power", "L1, L2 / 230V AUX"), ("Discharge Resistors", "D1, D2")]
    }.get(scheme_type, [])
    return colors_info, ferrule_sample

def get_ip_rating_recommendation(location, dust_level, water_level):
    if "Outdoor" in location:
        ip = "IP65 / IP66" if "Jets" in water_level else "IP55"
        gasket = "CNC Polyurethane / EPDM Seamless"
        canopy = "Required (Slanted Rain Canopy)"
    else:
        if "Jets" in water_level:
            ip = "IP65"; gasket = "Neoprene / EPDM Gasket"; canopy = "Drip Top Optional"
        elif "Splashing" in water_level or "Heavy" in dust_level:
            ip = "IP54"; gasket = "Neoprene Gasket Strip"; canopy = "Not Required"
        elif "Moderate" in dust_level or "Dripping" in water_level:
            ip = "IP42 / IP52"; gasket = "Standard Rubber Gasket"; canopy = "Not Required"
        else:
            ip = "IP20 / IP31"; gasket = "None / Dust Lip Only"; canopy = "Not Required"
    
    fan_filter = "IP54/IP55 Louver + Filter Pad" if "54" in ip or "55" in ip else ("IP65 Sealed Panel AC" if "65" in ip or "66" in ip else "Standard Ventilation Grille")
    return {"ip_rating": ip, "gasket": gasket, "canopy": canopy, "ventilation": fan_filter}

def calculate_ct_sizing(load_amps, wire_len_m, application="Panel MFM Metering"):
    stds = [50, 75, 100, 150, 200, 250, 300, 400, 500, 600, 800, 1000, 1250, 1600, 2000, 2500, 3200, 4000]
    pri = next((s for s in stds if s >= load_amps * 1.15), stds[-1])
    sec = 1 if wire_len_m > 30 else 5
    r_wire = (0.0178 * wire_len_m * 2) / 2.5
    va_wire = (sec ** 2) * r_wire
    va_meter = 0.5 if "Metering" in application or "MFM" in application else 1.5
    tot_va = (va_wire + va_meter) * 1.3
    burdens = [2.5, 5.0, 10.0, 15.0, 30.0]
    rec_va = next((b for b in burdens if b >= tot_va), 30.0)
    cls_map = {"Utility Billing": "Class 0.2S", "Panel MFM Metering": "Class 0.5", "Analog Door Meter": "Class 1.0", "Motor Protection Relay": "Class 5P10"}
    rec_cls = cls_map.get(application, "Class 0.5")
    return {"ct_ratio": f"{pri}/{sec}A", "rec_va": f"{rec_va} VA", "rec_class": rec_cls, "wire_loss_va": round(va_wire, 2)}

def calculate_class_ps_knee_point_voltage(ifault_ka, ct_pri, ct_sec, r_ct_ohms, wire_len_m, wire_sqmm=2.5, r_relay_ohms=0.2):
    i_sec_fault = (ifault_ka * 1000.0) / (ct_pri / ct_sec) if ct_pri > 0 else 0
    rho = 0.0178
    r_wire_single = (rho * wire_len_m) / wire_sqmm
    r_loop = 2.0 * r_wire_single
    v_k_min = round(2.0 * i_sec_fault * (r_ct_ohms + r_loop + r_relay_ohms), 1)
    i_mag_max = round(i_sec_fault * 0.03, 2)
    return {
        "i_sec_fault": round(i_sec_fault, 1),
        "r_loop": round(r_loop, 2),
        "v_k_min": v_k_min,
        "i_mag_max": f"{i_mag_max} A (3% of I_sec_fault)"
    }

def calculate_earthing_pit_resistance(soil_type, pit_type, pit_qty=2):
    rho = {"Moist Clay / Black Cotton Soil": 30.0, "Sandy Soil / Alluvial": 80.0, "Gravel / Rocky Soil": 250.0, "Hard Rock / Granite": 800.0}.get(soil_type, 80.0)
    if "Pipe" in pit_type:
        l_cm = 300.0; d_cm = 5.0
        r_single = (100.0 * rho / (2 * math.pi * l_cm)) * math.log((4 * l_cm) / d_cm)
    else:
        area_m2 = 0.36
        r_single = (rho / 4.0) * math.sqrt(math.pi / area_m2)
    r_net = round(r_single / (pit_qty * 0.85), 2) if pit_qty > 0 else round(r_single, 2)
    r_single = round(r_single, 2)
    status = "PASS ✅ Excellent Grounding (< 1.0 Ω - Safe for PLC/SCADA)" if r_net <= 1.0 else ("PASS ✅ Safe Grounding (< 5.0 Ω - LT Panel Standard)" if r_net <= 5.0 else "WARNING ⚠️ High Resistance (> 5.0 Ω - Add Chemical Gel / Extra Pit)")
    rec_strip = "50x6 mm GI Strip / 25x3 mm Copper" if r_net <= 1.0 else "25x6 mm GI Strip / 25x3 mm Copper"
    return {"rho": rho, "r_single": r_single, "r_net": r_net, "status": status, "rec_strip": rec_strip}

def calculate_voltage_unbalance(v_r, v_y, v_b):
    v_avg = (v_r + v_y + v_b) / 3.0
    max_dev = max(abs(v_r - v_avg), abs(v_y - v_avg), abs(v_b - v_avg))
    vuf_pct = round((max_dev / v_avg) * 100.0, 2) if v_avg > 0 else 0
    curr_unbalance_pct = round(vuf_pct * 7.5, 1)
    derating_factor = 1.0 if vuf_pct <= 1.0 else (0.95 if vuf_pct <= 2.0 else (0.90 if vuf_pct <= 3.0 else (0.82 if vuf_pct <= 4.0 else 0.75)))
    status = "PASS ✅ Safe (< 1% - No Derating Needed)" if vuf_pct <= 1.0 else ("WARNING ⚠️ Moderate Unbalance (1-3% - Derate Motor)" if vuf_pct <= 3.0 else ("WARNING ⚠️ High Unbalance (3-5% - Severe Derating Required)" if vuf_pct <= 5.0 else "FAIL ❌ Dangerous Unbalance (> 5% - Trip Motor Immediately)"))
    return {"v_avg": round(v_avg, 1), "vuf_pct": vuf_pct, "curr_unbalance_pct": curr_unbalance_pct, "derating_factor": derating_factor, "status": status}

def calculate_neutral_busbar_sizing(phase_current_a, non_linear_pct=40.0, harmonic_thd_pct=30.0):
    triplen_factor = (non_linear_pct / 100.0) * (harmonic_thd_pct / 100.0) * 1.732
    i_neutral = round(max(phase_current_a * 0.5, phase_current_a * (0.5 + triplen_factor)), 1)
    neutral_ratio_pct = round((i_neutral / phase_current_a) * 100.0, 1) if phase_current_a > 0 else 50.0
    
    rec_busbar_pct = "50% Reduced Neutral (Standard LT Panel)" if neutral_ratio_pct <= 50.0 else (
        "100% Full-Capacity Neutral Busbar" if neutral_ratio_pct <= 100.0 else (
            f"⚠️ 150% Oversized Neutral Busbar ({neutral_ratio_pct}%)" if neutral_ratio_pct <= 150.0 else
            f"⚠️ 200% Double-Capacity Neutral Busbar ({neutral_ratio_pct}%)"
        )
    )
    rec_breaker = "3P MCCB / ACB with 50% Neutral" if neutral_ratio_pct <= 50.0 else (
        "4P Breaker with 100% Protected Neutral (4P 4D)" if neutral_ratio_pct <= 100.0 else
        "4P Breaker with Adjustable Over-Sized Neutral Release (4P 4D + Neutral Protection)"
    )
    return {
        "i_neutral": i_neutral,
        "neutral_ratio_pct": neutral_ratio_pct,
        "rec_busbar_pct": rec_busbar_pct,
        "rec_breaker": rec_breaker
    }

def calculate_ieee43_ir_temp_correction(r_meas_momega, test_temp_c):
    k_t = round(0.5 ** ((40.0 - test_temp_c) / 10.0), 3)
    r_40 = round(r_meas_momega * k_t, 1)
    status = "PASS ✅ Excellent (> 100 MΩ @ 40°C Baseline)" if r_40 >= 100.0 else ("PASS ✅ Safe Baseline (> 5.0 MΩ @ 40°C Baseline)" if r_40 >= 5.0 else "FAIL ❌ High Moisture / Insulation Breakdown (< 5.0 MΩ)")
    return {"k_t": k_t, "r_40": r_40, "status": status}

def calculate_soft_starter(motor_kw, connection_type="In-Line"):
    flc = (motor_kw * 1000.0) / (1.732 * 415.0 * 0.85 * 0.88)
    ss_amps = flc if connection_type == "In-Line" else flc / 1.732
    bypass_a = math.ceil(ss_amps * 1.15)
    mccb_a = max(32, math.ceil(flc * 1.25 / 10.0) * 10)
    fuse_a = math.ceil(ss_amps * 1.25)
    return {
        "flc": round(flc, 1),
        "ss_amps": round(ss_amps, 1),
        "bypass_contactor": f"{bypass_a}A AC-1 / AC-3",
        "rec_mccb": f"{mccb_a}A Motor-Duty MCCB",
        "rec_fuse": f"{fuse_a}A aR Fast-Acting Fuse"
    }

def calculate_motor_master_chart(kw_val):
    flc = round((kw_val * 1000.0) / (1.732 * 415.0 * 0.85 * 0.88), 1)
    mccb = f"{max(32, math.ceil(flc * 1.25 / 10.0) * 10)}A MCCB"
    dol = f"{math.ceil(flc * 1.15)}A AC3"
    sd_main = f"{math.ceil((flc / 1.732) * 1.15)}A AC3"
    sd_star = f"{math.ceil((flc / 3.0) * 1.15)}A AC3"
    ss = f"{flc}A Soft Starter"
    cables_cu = [(2.5,24),(4,32),(6,41),(10,57),(16,76),(25,101),(35,125),(50,150),(70,190),(95,235),(120,270),(150,310),(185,355),(240,420),(300,480)]
    cables_al = [(4,25),(6,32),(10,44),(16,59),(25,78),(35,97),(50,116),(70,147),(95,182),(120,210),(150,240),(185,275),(240,325),(300,375)]
    cu = next((s for s, a in cables_cu if a >= flc * 1.25), cables_cu[-1][0])
    al = next((s for s, a in cables_al if a >= flc * 1.25), cables_al[-1][0])
    return {
        "Motor Rating": f"{kw_val} kW ({round(kw_val*1.341,1)} HP)",
        "FLC Current": f"{flc} A",
        "Incomer MCCB": mccb,
        "DOL Contactor": dol,
        "Star-Delta (Main/Delta)": sd_main,
        "Star-Delta (Star)": sd_star,
        "Soft Starter": ss,
        "Cu Cable": f"{cu} sq.mm",
        "Al Cable": f"{al} sq.mm"
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
    busbar_h_mm = 220
    gland_h_mm = 120
    
    current_bay = 1
    start_y = busbar_h_mm + 30 if busbar_pos == "Top" else 40
    
    cur_x = 40
    cur_y = start_y
    row_max_h = 0
    components_placed = []
    
    for _, row in bom_df.iterrows():
        item_name = str(row["Item"])
        qty = int(row["Qty"])
        
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
    
    svg.append(f'<rect x="{start_x}" y="{start_y}" width="{cable_alley_w}" height="{panel_h_mm}" fill="#1e293b" stroke="#64748b" stroke-width="3"/>')
    svg.append(f'<text x="{start_x + cable_alley_w/2}" y="{start_y + panel_h_mm/2}" fill="#94a3b8" font-size="16" font-weight="bold" text-anchor="middle" transform="rotate(-90,{start_x + cable_alley_w/2},{start_y + panel_h_mm/2})">CABLE ALLEY ({cable_alley_w}mm)</text>')
    
    cur_x = start_x + cable_alley_w
    for b_idx in range(1, num_bays + 1):
        svg.append(f'<rect x="{cur_x}" y="{start_y}" width="{bay_w_mm}" height="{panel_h_mm}" fill="#020617" stroke="#38bdf8" stroke-width="4"/>')
        
        bb_h = 200
        bb_y = start_y if busbar_pos == "Top" else (start_y + panel_h_mm - bb_h)
        svg.append(f'<rect x="{cur_x+10}" y="{bb_y}" width="{bay_w_mm-20}" height="{bb_h}" fill="#1e293b" stroke="#cbd5e1" stroke-width="2"/>')
        for i, color in enumerate(["#ef4444", "#eab308", "#3b82f6", "#000000"]):
            bar_y = bb_y + 20 + (i * 38)
            svg.append(f'<rect x="{cur_x+25}" y="{bar_y}" width="{bay_w_mm-50}" height="18" fill="{color}" stroke="#ffffff" stroke-width="1"/>')
        svg.append(f'<text x="{cur_x + bay_w_mm/2}" y="{bb_y + bb_h - 15}" fill="#f8fafc" font-size="14" font-weight="bold" text-anchor="middle">MAIN BUSBAR CHAMBER (R-Y-B-N)</text>')
        
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
            svg.append('</g>')
            
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
    
    svg.append(f'<rect x="{start_x}" y="{start_y}" width="{alley_w}" height="{panel_h_mm}" fill="#1e293b" stroke="#64748b" stroke-width="3"/>')
    svg.append(f'<circle cx="{start_x + alley_w - 20}" cy="{start_y + panel_h_mm/2}" r="10" fill="#475569" stroke="#cbd5e1" stroke-width="2"/>')
    
    cur_x = start_x + alley_w
    for b_idx in range(1, num_bays + 1):
        svg.append(f'<rect x="{cur_x}" y="{start_y}" width="{bay_w_mm}" height="{panel_h_mm}" fill="#0f172a" stroke="#38bdf8" stroke-width="4"/>')
        svg.append(f'<rect x="{cur_x + 10}" y="{start_y + 80}" width="12" height="35" fill="#64748b"/>')
        svg.append(f'<rect x="{cur_x + 10}" y="{start_y + panel_h_mm - 120}" width="12" height="35" fill="#64748b"/>')
        svg.append(f'<rect x="{cur_x + bay_w_mm - 30}" y="{start_y + panel_h_mm/2 - 25}" width="18" height="50" rx="4" fill="#334155" stroke="#94a3b8" stroke-width="2"/>')
        
        svg.append(f'<rect x="{cur_x + 40}" y="{start_y + 40}" width="{bay_w_mm - 80}" height="140" rx="6" fill="#020617" stroke="#0ea5e9" stroke-width="2"/>')
        svg.append(f'<rect x="{cur_x + 60}" y="{start_y + 60}" width="120" height="90" rx="4" fill="#0f172a" stroke="#38bdf8" stroke-width="2"/>')
        svg.append(f'<text x="{cur_x + 120}" y="{start_y + 112}" fill="#38bdf8" font-family="monospace" font-size="18" font-weight="bold" text-anchor="middle">415.2 V</text>')
        
        for idx, color in enumerate(["#ef4444", "#eab308", "#3b82f6"]):
            svg.append(f'<circle cx="{cur_x + bay_w_mm - 80}" cy="{start_y + 70 + (idx*30)}" r="10" fill="{color}"/>')
            
        svg.append(f'<rect x="{cur_x + 40}" y="{start_y + 210}" width="{bay_w_mm - 80}" height="180" rx="6" fill="#1e293b" stroke="#64748b" stroke-width="2"/>')
        svg.append(f'<circle cx="{cur_x + 100}" cy="{start_y + 270}" r="16" fill="#22c55e"/>')
        svg.append(f'<circle cx="{cur_x + 180}" cy="{start_y + 270}" r="16" fill="#ef4444"/>')
        svg.append(f'<circle cx="{cur_x + 260}" cy="{start_y + 270}" r="16" fill="#eab308"/>')
        
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

def generate_fat_certificate_pdf(serial, project, client, inspector, r, y, b, hv_pass):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    style_th_l = ParagraphStyle('THL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    style_td_l = ParagraphStyle('TDL', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=0)
    
    story = [
        Paragraph("ARYAVARTA AUTOMATION", ParagraphStyle('H1', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"))),
        Paragraph("<b>FACTORY ACCEPTANCE TEST (FAT) & QUALITY CERTIFICATE</b>", ParagraphStyle('H2', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor("#2B6CB0"))),
        Paragraph(f"Serial: {serial} | Client: {client} | Date: {pd.Timestamp.now().strftime('%d-%b-%Y')}", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=8, leading=10)),
        Spacer(1, 10), Paragraph("<b>1. Insulation Resistance Test (500V Megger)</b>", ParagraphStyle('H3', parent=styles['Heading3'], fontSize=10, leading=12))
    ]
    t_m = Table([
        [Paragraph("Phase", style_th_l), Paragraph("Measured Value", style_th_l), Paragraph("Min Required", style_th_l), Paragraph("Result", style_th_l)],
        [Paragraph("R-E", style_td_l), Paragraph(f"{r} MΩ", style_td_l), Paragraph("> 2.0 MΩ", style_td_l), Paragraph("PASS" if r>=2 else "FAIL", style_td_l)],
        [Paragraph("Y-E", style_td_l), Paragraph(f"{y} MΩ", style_td_l), Paragraph("> 2.0 MΩ", style_td_l), Paragraph("PASS" if y>=2 else "FAIL", style_td_l)],
        [Paragraph("B-E", style_td_l), Paragraph(f"{b} MΩ", style_td_l), Paragraph("> 2.0 MΩ", style_td_l), Paragraph("PASS" if b>=2 else "FAIL", style_td_l)]
    ], colWidths=[130, 130, 140, 100])
    t_m.setStyle(TableStyle([('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")), ('PADDING', (0,0), (-1,-1), 4)]))
    story.extend([t_m, Spacer(1, 10), Paragraph("<b>2. HV & Safety Checks</b>", ParagraphStyle('H3', parent=styles['Heading3'], fontSize=10, leading=12))])
    t_c = Table([
        [Paragraph("Check Item", style_th_l), Paragraph("Standard", style_th_l), Paragraph("Status", style_th_l)],
        [Paragraph("2.5kV HV Withstand (1 Min)", style_td_l), Paragraph("IEC 61439-1", style_td_l), Paragraph("PASSED" if hv_pass else "FAILED", style_td_l)],
        [Paragraph("Busbar Phase Clearance", style_td_l), Paragraph("IS 8623", style_td_l), Paragraph("PASSED", style_td_l)],
        [Paragraph("Door Earth Continuity", style_td_l), Paragraph("Safety Standard", style_td_l), Paragraph("PASSED", style_td_l)]
    ], colWidths=[240, 180, 120])
    t_c.setStyle(TableStyle([('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#4A5568")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")), ('PADDING', (0,0), (-1,-1), 4)]))
    story.extend([t_c, Spacer(1, 20), Paragraph(f"QA Inspector: {inspector} &nbsp;&nbsp;&nbsp;&nbsp; Quality Manager Stamp: ____________", ParagraphStyle('Sig', parent=styles['Normal'], fontSize=9, leading=12))])
    doc.build(story); buf.seek(0); return buf

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
    doc.build(story); buf.seek(0); return buf

def generate_delivery_challan_pdf(dc_no, client_name, address, vehicle_no, panel_type, gross_weight):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    style_th_l = ParagraphStyle('THL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    style_td_l = ParagraphStyle('TDL', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=0)

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
        ('PADDING', (0,0), (-1,-1), 5)
    ]))
    t_items = Table([
        [Paragraph("S.No", style_th_l), Paragraph("Item Description", style_th_l), Paragraph("HSN Code", style_th_l), Paragraph("Qty", style_th_l), Paragraph("Gross Wt (kg)", style_th_l), Paragraph("Status", style_th_l)],
        [Paragraph("1", style_td_l), Paragraph(f"{panel_type} Industrial Control Panel", style_td_l), Paragraph("85371010", style_td_l), Paragraph("1 Set", style_td_l), Paragraph(f"{gross_weight} kg", style_td_l), Paragraph("Supply Only", style_td_l)]
    ], colWidths=[40, 210, 80, 50, 80, 80])
    t_items.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 4)
    ]))
    story.extend([t_meta, Spacer(1, 10), t_items, Spacer(1, 15), Paragraph("<b>Declaration:</b> Material dispatched in good condition for installation & commissioning. Not for sale.", ParagraphStyle('Dec', parent=styles['Normal'], fontSize=8, leading=10)), Spacer(1, 20), Paragraph("<b>Receiver's Stamp & Signature:</b> ____________________ &nbsp;&nbsp;&nbsp;&nbsp; <b>For ARYAVARTA AUTOMATION</b>", ParagraphStyle('Sig', parent=styles['Normal'], fontSize=9, leading=12))])
    doc.build(story); buf.seek(0); return buf

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
    doc.build(story); buf.seek(0); return buf

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
        if "hp" in kw_match.group(0):
            val = val * 0.7457
        std_kws = [7.5, 15, 22, 37, 55, 75]
        closest = min(std_kws, key=lambda x: abs(x - val))
        found_kw = f"{closest if isinstance(closest, int) or closest.is_integer() else closest} kW".replace('.0 kW', ' kW')

    found_panel = None
    if "apfc" in text_lower or "power factor" in text_lower:
        found_panel = "APFC Panel"
    elif "star" in text_lower or "delta" in text_lower:
        found_panel = "Star-Delta Control Panel"
    elif "lt" in text_lower or "distribution" in text_lower or "pcc" in text_lower:
        found_panel = "LT Distribution Panel"
    elif "vfd" in text_lower or "drive" in text_lower or "inverter" in text_lower:
        found_panel = "VFD Panel"

    found_brand = None
    for b in ["Siemens", "Schneider", "L&T", "ABB", "Danfoss", "Delta"]:
        if b.lower() in text_lower:
            found_brand = b
            break

    return {
        "motor_kw": found_kw,
        "panel_type": found_panel,
        "preferred_brand": found_brand
    }

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
    with st.expander("🤖 Auto-Extract Specs from Customer Text or Upload Document (PDF, Excel, Word, TXT, CSV)", expanded=True):
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
    
    with st.expander("🛠️ Interactive Panel Sizing & Enclosure Bay Configurator", expanded=True):
        ga_c1, ga_c2, ga_c3, ga_c4 = st.columns(4)
        panel_h_input = ga_c1.number_input("Panel Frame Height (mm)", value=2000, step=100)
        bay_w_input = ga_c2.number_input("Bay Width (mm)", value=800, step=100)
        panel_d_input = ga_c3.number_input("Panel Depth (mm)", value=600, step=50)
        cable_alley_w_input = ga_c4.number_input("Cable Alley Width (mm)", value=200, step=50)
        busbar_pos_input = st.radio("Main Busbar Chamber Position", ["Top", "Bottom"], horizontal=True)

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
    st.dataframe(pd.DataFrame(chart_data), width="stretch")

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

elif menu == "ACB Selection & Protection Release Sizer":
    st.header("⚡ Air Circuit Breaker (ACB) & Microprocessor Release Sizer (IEC 60947-2)")
    c1, c2, c3 = st.columns(3)
    tx_k = c1.number_input("Incomer Transformer / Load Capacity (kVA)", value=1000.0, step=100.0)
    icu_ka = c2.selectbox("Required Fault Capacity (Icu)", [36, 50, 65, 85, 100], index=1)
    exec_type = c3.selectbox("Breaker Execution", ["Drawout Type (4P)", "Drawout Type (3P)", "Fixed Type (4P)", "Fixed Type (3P)"])
    
    acb_res = calculate_acb_and_protection_settings(tx_k, 0.85, icu_ka, exec_type)
    st.divider()
    a1, a2, a3 = st.columns(3)
    a1.metric("Nominal Incomer Load", f"{acb_res['i_nom']} A")
    a2.metric("Recommended ACB Breaker", acb_res['rec_acb'])
    a3.metric("Enclosure Frame Size", acb_res['frame'])
    
    st.subheader("🎛️ Microprocessor Electronic Trip Unit (ETU) L-S-I-G Release Settings")
    es1, es2, es3, es4 = st.columns(4)
    es1.metric("Overload (Long Time - Ir)", acb_res['ir_setting'])
    es2.metric("Selective Short (Isd)", acb_res['isd_setting'])
    es3.metric("Instant Short (Ii)", acb_res['ii_setting'])
    es4.metric("Earth Fault (Ig)", acb_res['ig_setting'])

elif menu == "Motor Starting Voltage Dip Calculator":
    st.header("⚡ Motor Starting Inrush Current & Line Voltage Dip Sizer (IEC 60947)")
    c1, c2, c3 = st.columns(3)
    m_kw_v = c1.number_input("Motor Power Rating (kW)", value=30.0, step=2.5)
    s_meth = c2.selectbox("Starting Method", ["DOL", "Star-Delta", "Soft Starter", "VFD"])
    c_dist = c3.number_input("Cable Length to Motor (Meters)", value=50, step=5)
    
    c4, c5 = st.columns(2)
    c_sq = c4.selectbox("Cable Conductor Size (sq.mm)", [2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240, 300], index=6)
    c_mat = c5.radio("Conductor Material", ["Copper", "Aluminum"])

    vd_res = calculate_motor_starting_vdrop(m_kw_v, s_meth, c_dist, c_sq, c_mat)
    st.divider()
    vd1, vd2, vd3, vd4 = st.columns(4)
    vd1.metric("Motor FLC", f"{vd_res['flc']} A")
    vd2.metric("Peak Starting Inrush", f"{vd_res['i_start']} A")
    vd3.metric("Transient Voltage Drop", f"{vd_res['v_drop_v']} V ({vd_res['v_drop_pct']} %)")
    vd4.metric("Motor Terminal Voltage", f"{vd_res['term_v']} V")

elif menu == "PLC & HMI Automation Sizer":
    st.header("🎛️ PLC & HMI Automation System I/O Sizer")
    c1, c2, c3 = st.columns(3)
    with c1:
        di_cnt = st.number_input("Digital Inputs (Pushbuttons/Sensors)", value=16, min_value=0)
        do_cnt = st.number_input("Digital Outputs (Contactors/Valves)", value=12, min_value=0)
    with c2:
        ai_cnt = st.number_input("Analog Inputs (4-20mA Transmitters)", value=4, min_value=0)
        ao_cnt = st.number_input("Analog Outputs (0-10V VFD Speed)", value=2, min_value=0)
    with c3:
        p_brand = st.selectbox("Preferred PLC Brand", ["Siemens", "Delta", "Schneider"])
        h_size = st.selectbox("Touchscreen HMI Display", ["7.0 Inch", "4.3 Inch", "10.0 Inch", "None"])

    plc_res = calculate_plc_hmi_sizing(di_cnt, do_cnt, ai_cnt, ao_cnt, h_size, p_brand)
    st.divider()
    l1, l2, l3, l4 = st.columns([1, 1.3, 1.6, 1.2])
    l1.metric("Total System I/O", f"{plc_res['tot_io']} Signals")
    l2.metric("Recommended CPU Base", plc_res['cpu'])
    l3.metric("Required Expansion Modules", plc_res['cards'])
    l4.metric("Est. Hardware Cost", f"₹ {plc_res['est_hw_cost']:,.2f}")

elif menu == "Star-Delta & DOL Switchgear Sizing":
    st.header("⚡ Star-Delta & DOL Starter Component Sizing (IEC 60947-4-1)")
    col1, col2 = st.columns(2)
    sd_kw = col1.number_input("Motor Rating (kW)", value=30.0, step=2.5)
    col2.metric("Equivalent Horsepower", f"{round(sd_kw * 1.341, 1)} HP")
    sd_res = calculate_star_delta_starter(sd_kw)
    st.divider()
    s1, s2, s3, s4 = st.columns([1, 1, 1.2, 1.2])
    s1.metric("Motor FLC (Line)", f"{sd_res['flc']} A")
    s2.metric("Phase Current (In-Delta)", f"{sd_res['i_phase']} A")
    s3.metric("Main & Delta Contactors", sd_res['main_delta_contactor'])
    s4.metric("Star Contactor", sd_res['star_contactor'])

elif menu == "Soft Starter & Bypass Contactor Sizer":
    st.header("⚡ Soft Starter & Bypass Contactor Sizer (IEC 60947-4-2)")
    col1, col2 = st.columns(2)
    ss_kw = col1.number_input("Motor Power Rating (kW)", value=45.0, step=5.0)
    ss_conn = col2.radio("Wiring Configuration", ["In-Line Connection", "Inside-Delta Connection"])
    
    ss_res = calculate_soft_starter(ss_kw, "Inside-Delta" if "Inside" in ss_conn else "In-Line")
    st.divider()
    s1, s2, s3, s4 = st.columns([1, 1.1, 1.3, 1.4])
    s1.metric("Motor FLC Current", f"{ss_res['flc']} A")
    s2.metric("Required Soft Starter Rating", f"{ss_res['ss_amps']} A")
    s3.metric("Recommended Bypass Contactor", ss_res["bypass_contactor"])
    s4.metric("Main Incomer MCCB", ss_res["rec_mccb"])

elif menu == "Enclosure Fabrication Costing":
    st.header("🛠️ CRCA Sheet Metal & Powder Coating Estimator")
    c1, c2, c3 = st.columns(3)
    with c1:
        h = st.number_input("Height (mm)", value=1200, step=100)
        w = st.number_input("Width (mm)", value=800, step=100)
        d = st.number_input("Depth (mm)", value=400, step=50)
    with c2:
        gauge = st.selectbox("CRCA Sheet Thickness", [1.6, 2.0, 1.2, 2.5], index=0)
        steel_rate = st.number_input("CRCA Steel Rate (₹ / kg)", value=82.0, step=2.0)
        coating_rate = st.number_input("Powder Coating Rate (₹ / m²)", value=220.0, step=10.0)
    with c3:
        insulators = st.number_input("SMC Busbar Support Insulators (Qty)", value=12, step=2)
        locks = st.number_input("Concealed Door Locks & Handles (Qty)", value=2, step=1)

    f_res = calculate_enclosure_fabrication(h, w, d, gauge, steel_rate, coating_rate, insulators, locks)
    st.divider()
    fc1, fc2, fc3, fc4, fc5 = st.columns(5)
    fc1.metric("Effective Sheet Area", f"{f_res['area_m2']} m²")
    fc2.metric("CRCA Steel Weight", f"{f_res['weight_kg']} kg")
    fc3.metric("Raw Steel Cost", f"₹ {f_res['crca_cost']:,.2f}")
    fc4.metric("Hardware & Insulators", f"₹ {f_res['hardware_cost']:,.2f}")
    fc5.metric("Est. Fabricated Box Cost", f"₹ {f_res['total_enc_cost']:,.2f}")

elif menu == "VFD Energy Savings & Payback":
    st.header("🌱 VFD Energy Savings & Payback Calculator (Affinity Laws)")
    c1, c2 = st.columns(2)
    with c1:
        m_kw_val = st.number_input("Motor Rating (kW)", value=30.0, step=5.0)
        run_hrs = st.number_input("Daily Operating Hours", value=16, max_value=24, min_value=1)
        tariff = st.number_input("Electricity Rate (₹ / kWh)", value=9.5, step=0.5)
    with c2:
        speed_red = st.slider("Average Flow/Speed Reduction (%)", 5, 40, 20)
        vfd_cost = st.number_input("Total VFD Panel Cost (₹)", value=65000, step=5000)

    res_pb = calculate_vfd_payback(m_kw_val, run_hrs, tariff, speed_red, vfd_cost)
    st.divider()
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Power Reduction", f"{res_pb['kw_saved_hr']} kW/hr")
    p2.metric("Annual Energy Saved", f"{res_pb['annual_kwh']:,.0f} kWh")
    p3.metric("Annual Bill Savings", f"₹ {res_pb['annual_inr']:,.2f}")
    p4.metric("Payback Period", f"{res_pb['payback_months']} Months")

elif menu == "VFD Dynamic Braking Resistor (DBR) Sizer":
    st.header("⚡ VFD Dynamic Braking Resistor (DBR) Sizer")
    c1, c2, c3 = st.columns(3)
    d_kw = c1.number_input("Motor Power Rating (kW)", value=22.0, step=2.5)
    d_load = c2.selectbox("Application & Braking Duty Cycle", [
        "Light Stopping (5% Duty)",
        "Standard Machine (10% Duty)",
        "High Inertia Centrifuge / Fan (20% Duty)",
        "Heavy Crane / Hoist Overhauling Load (50% Duty)"
    ])
    d_v = c3.selectbox("VFD Nominal DC Bus Voltage", [750, 680, 310], index=0)
    
    dbr_res = calculate_dbr_sizing(d_kw, d_load, d_v)
    st.divider()
    db1, db2, db3, db4 = st.columns(4)
    db1.metric("Peak Braking Power", f"{dbr_res['peak_kw']} kW")
    db2.metric("Min Resistance Value", f"{dbr_res['res_ohms']} Ω")
    db3.metric("Continuous Power Rating", f"{dbr_res['cont_watts']} W")
    db4.metric("Braking Duty Cycle", dbr_res['duty_pct'])

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

    with st.expander("⚖️ Copper vs. Aluminum Busbar Weight & Cost Optimization Analyzer", expanded=False):
        c_u1, c_u2, c_u3 = st.columns(3)
        b_len_m = c_u1.number_input("Total Busbar Run Length (Meters)", value=3.0, step=0.5)
        cu_p = c_u2.number_input("Copper Rate (₹ / kg)", value=850.0, step=10.0)
        al_p = c_u3.number_input("Aluminum Rate (₹ / kg)", value=260.0, step=10.0)
        
        ca_res = calculate_cu_vs_al_busbar(res['design_current'], b_len_m, cu_p, al_p)
        ca1, ca2, ca3, ca4 = st.columns(4)
        ca1.metric("Copper Busbar Weight / Cost", f"{ca_res['wt_cu']} kg / ₹ {ca_res['cost_cu']:,.2f}")
        ca2.metric("Aluminum Equivalent Wt / Cost", f"{ca_res['wt_al']} kg / ₹ {ca_res['cost_al']:,.2f}")
        ca3.metric("Weight Saved with Aluminum", f"{ca_res['wt_saved']} kg ({round((ca_res['wt_saved']/ca_res['wt_cu'])*100,1) if ca_res['wt_cu']>0 else 0}%)")
        ca4.metric("Cost Saved with Aluminum", f"₹ {ca_res['cost_saved']:,.2f} ({ca_res['pct_cost_saved']}%)")

    with st.expander("🔩 Busbar Joint Overlap & Bolt Tightening Torque Guide (IS 8623 / IEC 61439)", expanded=False):
        t_df = pd.DataFrame([
            {"Bolt Size": "M6 Bolt", "Tightening Torque (N·m)": "9.0 N·m", "Min Contact Overlap": "1.0 x Busbar Width (W)", "Spring Washer Spec": "Belleville / Disc Spring + Flat"},
            {"Bolt Size": "M8 Bolt", "Tightening Torque (N·m)": "20.0 N·m", "Min Contact Overlap": "1.0 x Busbar Width (W)", "Spring Washer Spec": "Belleville / Disc Spring + Flat"},
            {"Bolt Size": "M10 Bolt", "Tightening Torque (N·m)": "40.0 N·m", "Min Contact Overlap": "1.0 x Busbar Width (W)", "Spring Washer Spec": "Belleville / Disc Spring + Flat"},
            {"Bolt Size": "M12 Bolt", "Tightening Torque (N·m)": "70.0 N·m", "Min Contact Overlap": "1.0 x Busbar Width (W)", "Spring Washer Spec": "Belleville / Disc Spring + Flat"}
        ])
        st.dataframe(t_df)

elif menu == "Cable Gland & Terminal Sizer":
    st.header("🔌 Cable Gland, Terminal Block & Cable Tray Sizer")
    c1, c2, c3 = st.columns(3)
    c_sqmm = c1.number_input("Power Cable Conductor Size (sq.mm)", value=35.0, step=5.0)
    c_cores = c2.selectbox("Cable Cores", ["3.5 Core Armored", "4 Core Armored", "3 Core Armored", "Single Core"])
    c_io = c3.number_input("Control Wire Signals / Terminals", value=14, min_value=0)
    
    g_res = calculate_gland_and_terminals(c_sqmm, 4 if "4" in c_cores else 3, c_io)
    st.divider()
    g1, g2, g3, g4 = st.columns([1.2, 1.2, 1.3, 1.3])
    g1.metric("Recommended Cable Gland", g_res["gland"])
    g2.metric("Power Terminal Block", g_res["p_tb"])
    g3.metric("Control Terminals (DIN)", g_res["ctrl_tbs"])
    g4.metric("Cable Tray Width", g_res["rec_tray"])

elif menu == "Control Wire Color & Ferrule Guide":
    st.header("🎨 Control Wiring Color Code & Ferrule Schedule (IS 375 / IEC 60446)")
    scheme = st.selectbox("Select Panel Control Scheme", ["VFD Panel", "Star-Delta Panel", "APFC Panel"])
    cols_df, ferr_list = get_control_wiring_and_ferrule_schedule(scheme)
    st.subheader("1. Standard Wire Color Code & Cross-Section Guide")
    st.dataframe(pd.DataFrame(cols_df))
    st.subheader(f"2. Recommended Ferrule Tag Scheme ({scheme})")
    st.table(pd.DataFrame(ferr_list, columns=["Circuit Function", "Standard Ferrule Tags"]))

elif menu == "IP Protection Rating & Gasket Guide":
    st.header("🛡️ Ingress Protection (IP Rating) & Gasket Guide (IEC 60529 / IS 13947)")
    c1, c2, c3 = st.columns(3)
    loc = c1.selectbox("Installation Location", ["Indoor Substation / Control Room", "Indoor Plant Floor (Dust/Moisture)", "Outdoor / Exposed Environment"])
    dust = c2.selectbox("Dust & Particle Exposure", ["Low / Normal Ambient", "Moderate Dust", "Heavy Industrial Dust / Cement / Flyash"])
    water = c3.selectbox("Liquid / Water Exposure", ["Dry / No Water", "Dripping Water / Condensation", "Splashing Water", "High Pressure Jets / Washdown"])
    
    ip_res = get_ip_rating_recommendation(loc, dust, water)
    st.divider()
    i1, i2, i3, i4 = st.columns([1.2, 1.5, 1.5, 1.4])
    i1.metric("Recommended Rating", ip_res["ip_rating"])
    i2.metric("Door Sealing Gasket", ip_res["gasket"])
    i3.metric("Rain Canopy Roof Shield", ip_res["canopy"])
    i4.metric("Louver & Cooling Plan", ip_res["ventilation"])

elif menu == "Short-Circuit kA & Copper Weight":
    st.header("💥 Short-Circuit Fault Level (kA) & Copper Weight Estimator")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Upstream Transformer & Fault Level")
        tx_kva = st.number_input("Transformer Capacity (kVA)", value=1000, step=100)
        tx_z = st.number_input("Transformer Impedance Z (%)", value=5.0, step=0.5)
    with col2:
        st.subheader("2. Panel Copper Busbar Sizing")
        b_w = st.number_input("Busbar Width (mm)", value=50, step=5)
        b_t = st.number_input("Busbar Thickness (mm)", value=6, step=1)
        b_len = st.number_input("Panel Length / Run (Meters)", value=2.5, step=0.5)
        cu_rate = st.number_input("Copper Rate (₹ / kg)", value=850, step=10)

    sc_res = calculate_short_circuit_and_cu_weight(tx_kva, tx_z, b_w, b_t, b_len, cu_rate)
    st.divider()
    k1, k2, k3, k4 = st.columns([1, 1.1, 1.2, 1.2])
    k1.metric("Est. Short-Circuit Current", f"{sc_res['i_sc_ka']} kA")
    k2.metric("Required MCCB Rating", sc_res['rec_mccb_ka'])
    k3.metric("Total Copper Busbar Weight", f"{sc_res['cu_weight_kg']} kg")
    k4.metric("Est. Raw Copper Material Cost", f"₹ {sc_res['cu_cost_inr']:,.2f}")

elif menu == "Transformer & DG Sizing":
    st.header("⚡ Substation Transformer & DG Set Sizing Calculator")
    c1, c2 = st.columns(2)
    with c1:
        tot_kw = st.number_input("Total Connected Panel Load (kW)", value=150.0, step=10.0)
        largest_kw = st.number_input("Largest Single Motor Rating (kW)", value=45.0, step=5.0)
    with c2:
        s_method = st.selectbox("Largest Motor Starting Method", ["VFD", "Soft Starter", "Star-Delta", "DOL"])
        div_factor = st.slider("Plant Diversity Factor", 0.4, 1.0, 0.75, step=0.05)
    
    dg_res = calculate_dg_and_transformer_sizing(tot_kw, largest_kw, s_method, div_factor)
    st.divider()
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Operating Running Load", f"{dg_res['running_kva']} kVA")
    d2.metric("Peak Starting Demand", f"{dg_res['peak_kva']} kVA")
    d3.metric("Recommended Transformer", dg_res['rec_tx'])
    d4.metric("Recommended DG Set", dg_res['rec_dg'])

elif menu == "APFC Capacitor Sizing & MSEDCL ROI":
    st.header("⚡ APFC Capacitor Bank Sizing & MSEDCL Bill Savings Calculator")
    c1, c2, c3 = st.columns(3)
    with c1:
        kw = st.number_input("Plant Active Load (kW)", 150.0, step=10.0)
        monthly_bill = st.number_input("Avg. Monthly Electricity Bill (₹)", value=180000, step=10000)
    with c2:
        pf1 = st.slider("Current PF (cos φ1)", 0.50, 0.95, 0.78)
        pf2 = st.slider("Target PF (cos φ2)", 0.90, 0.99, 0.98)
    with c3:
        apfc_cost = st.number_input("Estimated APFC Panel Price (₹)", value=125000, step=5000)
        harm = st.checkbox("Harmonics Present (VFD / UPS)", value=True)

    apfc_res = calculate_apfc_and_msedcl_rebate(kw, pf1, pf2, harm, monthly_bill, apfc_cost)
    st.divider()
    ap1, ap2, ap3, ap4 = st.columns([1.1, 1, 1.3, 1])
    ap1.metric("Required Capacitor Bank", f"{apfc_res['req_kvar']} kVAR")
    ap2.metric("Line Current Savings", f"{apfc_res['pct']} %")
    ap3.metric("Est. Monthly Bill Savings", f"₹ {apfc_res['monthly_savings']:,.2f}")
    ap4.metric("Payback Period", f"{apfc_res['payback_months']} Months")
    st.success(f"**Recommended Steps:** {apfc_res['steps']}")

    with st.expander("⚡ Detuned Series Reactor & Capacitor Voltage Rating Sizer (IEC 61431)", expanded=False):
        d_p = st.selectbox("Detuned Reactor Tuning Factor", [7.0, 5.67, 14.0], format_func=lambda x: f"{x}% Detuned ({'5th/7th' if x<=7 else '3rd'} Harmonics)")
        det_res = calculate_detuned_reactor_and_capacitor(apfc_res['req_kvar'], 415.0, d_p)
        dr1, dr2, dr3, dr4 = st.columns(4)
        dr1.metric("Capacitor Terminal Voltage", f"{det_res['v_cap_op']} V")
        dr2.metric("Capacitor Duty Rating", det_res['rec_v_rating'])
        dr3.metric("Nameplate Rating Needed", f"{det_res['nameplate_kvar']} kVAR")
        dr4.metric("Series Reactor Sizing", f"{det_res['reactor_kvar']} kVAR ({det_res['tuning_freq']})")

elif menu == "Harmonics & AHF Sizing":
    st.header("⚡ VFD Harmonics & Active Filter (AHF) Calculator (IEEE 519)")
    c1, c2, c3 = st.columns(3)
    vfd_kw_val = c1.number_input("VFD Motor Power Rating (kW)", value=45.0, step=5.0)
    vfd_qty_val = c2.number_input("Quantity of VFD Drives", value=2, min_value=1)
    choke_val = c3.checkbox("3% AC Line Reactor Installed", value=True)

    h_res = calculate_harmonics_and_ahf(vfd_kw_val, vfd_qty_val, choke_val)
    st.divider()
    h1, h2, h3, h4 = st.columns([1, 1, 1, 1.4])
    h1.metric("Total VFD Load Current", f"{h_res['vfd_current']} A")
    h2.metric("Est. Harmonic Distortion", f"{h_res['thdi_pct']} % THDi")
    h3.metric("Harmonic Current Spikes", f"{h_res['harmonic_amps']} A")
    h4.metric("Required AHF Capacity", f"{h_res['req_ahf_amps']} A")

    with st.expander("⚡ VFD Motor Cable Length, EMC/RFI Filter & dv/dt Spike Sizer (IEC 61800-3)", expanded=False):
        e_col1, e_col2 = st.columns(2)
        v_dist_m = e_col1.number_input("Motor Cable Distance (Meters)", value=45.0, step=5.0)
        e_env = e_col2.selectbox("Installation Environment", ["Industrial (Class C3)", "Commercial / Public (Class C2)"])
        emc_res = calculate_emc_and_dvdt_filter(vfd_kw_val, v_dist_m, e_env)
        em1, em2, em3 = st.columns(3)
        em1.metric("EMC/RFI Noise Filter", emc_res['rec_emc_filter'])
        em2.metric("Output dv/dt Spike Protection", emc_res['dvdt_rec'])
        em3.metric("Motor Wiring Specification", emc_res['cable_spec'])

elif menu == "Panel Thermal & Fan Sizing":
    st.header("🌡️ Thermal Dissipation & Fan Airflow (IEC 60890)")
    c1, c2, c3 = st.columns(3)
    h = c1.number_input("Height (mm)", 1200); w = c1.number_input("Width (mm)", 800); d = c1.number_input("Depth (mm)", 400)
    vfd_kw = c2.number_input("VFD Power (kW)", 15.0); vfd_q = c2.number_input("VFD Quantity", 1); sealed = c2.checkbox("IP55 Sealed Panel")
    outdoor = c2.checkbox("Outdoor Sun Exposure (Solar Heat Load)")
    amb_t = c3.slider("Ambient Temp (°C)", 25, 50, 38); int_t = c3.slider("Max Internal Temp (°C)", 35, 55, 45)
    res = calculate_panel_thermal(h, w, d, vfd_kw, vfd_q, amb_t, int_t, sealed, outdoor)
    t1, t2, t3, t4 = st.columns(4)
    t1.metric("Internal Heat", f"{res['tot_w']} W"); t2.metric("Radiation", f"{res['nat_w']} W"); t3.metric("Net Heat", f"{res['net_w']} W"); t4.metric("Airflow", f"{res['m3h']} m³/h")

elif menu == "Control Tx & 24V SMPS Sizing":
    st.header("🔌 Control Transformer & 24V DC SMPS Power Calculator")
    c1, c2 = st.columns(2)
    with c1:
        cnt_q = st.number_input("Power Contactors / Starters Qty", value=4, min_value=0)
        rel_q = st.number_input("24V / 230V Auxiliary Relays Qty", value=6, min_value=0)
        htr_w = st.number_input("Panel Space Heater Rating (W)", value=80, step=10)
    with c2:
        hmi_on = st.checkbox("HMI / PLC Touchscreen Display Present", value=True)
        sens_q = st.number_input("4-20mA Transmitters / Sensors Qty", value=4, min_value=0)

    tx_res = calculate_smps_and_control_tx(cnt_q, rel_q, hmi_on, htr_w, sens_q)
    st.divider()
    s1, s2, s3, s4 = st.columns([1, 1.3, 1, 1.3])
    s1.metric("AC Control Burden", f"{tx_res['ac_va']} VA")
    s2.metric("Recommended Control Tx", tx_res['rec_tx'])
    s3.metric("24V DC Control Current", f"{tx_res['dc_amps']} A")
    s4.metric("Recommended 24V DC SMPS", tx_res['rec_smps'])

    with st.expander("⚡ 24V DC Field Control Wiring Voltage Drop & Distance Sizer", expanded=False):
        d1, d2 = st.columns(2)
        dc_wire_sq = d1.selectbox("Control Wire Size (sq.mm)", [0.5, 0.75, 1.0, 1.5, 2.5], index=1)
        dc_dist_m = d2.number_input("Field Wiring Distance to Sensors/Relays (Meters)", value=30.0, step=5.0)
        dc_vd_res = calculate_dc_control_vdrop(tx_res['dc_amps'], dc_wire_sq, dc_dist_m)
        dv1, dv2, dv3, dv4 = st.columns(4)
        dv1.metric("Voltage Drop", f"{dc_vd_res['v_drop_dc']} V DC")
        dv2.metric("Terminal Voltage", f"{dc_vd_res['end_v']} V DC")
        dv3.metric("Max Run (15% Drop Limit)", f"{dc_vd_res['max_dist_m']} Meters")
        dv4.metric("Status Check", dc_vd_res['status'])

elif menu == "Control Panel UPS & Battery Sizer":
    st.header("🔋 Control Panel UPS & Battery Backup Sizer (IEC 62040)")
    c1, c2, c3 = st.columns(3)
    load_va = c1.number_input("Total Control Load (VA / Watts)", value=600, step=50)
    bk_mins = c2.number_input("Required Backup Time (Minutes)", value=30, step=10)
    dc_v = c3.selectbox("DC Battery Bus Voltage", [24, 48, 12, 120], index=0)
    
    ups_res = calculate_ups_and_battery(load_va, bk_mins, dc_v)
    st.divider()
    u1, u2, u3, u4 = st.columns(4)
    u1.metric("Recommended Online UPS", ups_res["rec_ups"])
    u2.metric("Calculated Ah Capacity", f"{ups_res['req_ah']} Ah")
    u3.metric("Recommended Battery Bank", ups_res["rec_battery"])
    u4.metric("Min Charger Current", f"{ups_res['chg_amps']} A")

elif menu == "CT Ratio, Class & Burden Sizer":
    st.header("⚡ Current Transformer (CT) Ratio, Class & Burden Sizer (IS 2705 / IEC 61869)")
    c1, c2, c3 = st.columns(3)
    i_load = c1.number_input("Full Load / Incomer Current (Amps)", value=250.0, step=10.0)
    dist_m = c2.number_input("One-Way Wire Distance to Meter (Meters)", value=5.0, step=1.0)
    app_type = c3.selectbox("Application Type", ["Panel MFM Metering", "Utility Billing", "Analog Door Meter", "Motor Protection Relay"])
    
    ct_res = calculate_ct_sizing(i_load, dist_m, app_type)
    st.divider()
    ct1, ct2, ct3, ct4 = st.columns(4)
    ct1.metric("Recommended CT Ratio", ct_res["ct_ratio"])
    ct2.metric("Required CT Burden", ct_res["rec_va"])
    ct3.metric("Accuracy Class", ct_res["rec_class"])
    ct4.metric("Secondary Cable Loss", f"{ct_res['wire_loss_va']} VA")

    with st.expander(r"⚡ Protection Class PS CT Knee Point Voltage ($V_k$) & Saturation Sizer (IS 2705 / IEC 61869-2)", expanded=False):
        p_col1, p_col2, p_col3 = st.columns(3)
        i_f_ka = p_col1.number_input("Max Fault Current (kA)", value=25.0, step=5.0)
        ct_pri_val = p_col2.number_input("CT Primary Ratio (A)", value=400, step=50)
        ct_sec_val = p_col3.selectbox("CT Secondary Rating", [1, 5], index=0)
        
        p_col4, p_col5 = st.columns(2)
        r_ct_v = p_col4.number_input("CT Secondary Resistance R_ct (Ω)", value=1.5, step=0.1)
        r_rel_v = p_col5.number_input("Relay Input Impedance R_relay (Ω)", value=0.2, step=0.05)
        
        ps_res = calculate_class_ps_knee_point_voltage(i_f_ka, ct_pri_val, ct_sec_val, r_ct_v, dist_m, 2.5, r_rel_v)
        vk1, vk2, vk3, vk4 = st.columns(4)
        vk1.metric("Secondary Fault Current", f"{ps_res['i_sec_fault']} A")
        vk2.metric("Lead Loop Resistance (2xL)", f"{ps_res['r_loop']} Ω")
        vk3.metric("Min Knee Point Voltage (Vk)", f"{ps_res['v_k_min']} V")
        vk4.metric("Max Magnetizing Current (Imag)", ps_res['i_mag_max'])

elif menu == "Earthing Pit Resistance (IS 3043)":
    st.header("⚡ Earthing Pit Resistance & Neutral-to-Earth Sizer (IS 3043 / IEEE 80)")
    c1, c2, c3 = st.columns(3)
    s_type = c1.selectbox("Soil Type & Location", ["Moist Clay / Black Cotton Soil", "Sandy Soil / Alluvial", "Gravel / Rocky Soil", "Hard Rock / Granite"], index=1)
    e_type = c2.selectbox("Earthing Electrode Type", ["50mm x 3m GI Pipe Electrode", "600x600x6mm Copper Plate Electrode"])
    p_qty = c3.number_input("Number of Parallel Earth Pits", value=2, min_value=1, step=1)
    
    e_res = calculate_earthing_pit_resistance(s_type, e_type, p_qty)
    st.divider()
    er1, er2, er3, er4 = st.columns(4)
    er1.metric("Soil Resistivity (ρ)", f"{e_res['rho']} Ω·m")
    er2.metric("Single Pit Resistance", f"{e_res['r_single']} Ω")
    er3.metric("Net Grid Resistance", f"{e_res['r_net']} Ω")
    er4.metric("Earthing Conductor Strip", e_res['rec_strip'])

elif menu == "Voltage Unbalance & NEMA Derating (NEMA MG-1)":
    st.header("⚡ 3-Phase Voltage Unbalance & Motor Derating (NEMA MG-1 / IEC 60034-26)")
    c1, c2, c3 = st.columns(3)
    vr = c1.number_input("R-Phase Voltage (V)", value=415.0, step=1.0)
    vy = c2.number_input("Y-Phase Voltage (V)", value=405.0, step=1.0)
    vb = c3.number_input("B-Phase Voltage (V)", value=422.0, step=1.0)
    
    vu_res = calculate_voltage_unbalance(vr, vy, vb)
    st.divider()
    u1, u2, u3, u4 = st.columns(4)
    u1.metric("Average Phase Voltage", f"{vu_res['v_avg']} V")
    u2.metric("Voltage Unbalance (VUF)", f"{vu_res['vuf_pct']} %")
    u3.metric("Est. Current Unbalance", f"{vu_res['curr_unbalance_pct']} %")
    u4.metric("NEMA Derating Factor", f"{vu_res['derating_factor']} x kW")

elif menu == "Neutral Busbar & Triplen Harmonics (IEC 60364)":
    st.header("⚡ Neutral Busbar & Triplen Harmonics Sizer (IEC 60364-5-52)")
    c1, c2, c3 = st.columns(3)
    p_curr = c1.number_input("3-Phase Line Load Current (Amps)", value=250.0, step=10.0)
    nl_pct = c2.slider("Non-Linear Load Proportion (%)", 0, 100, 45)
    thd_pct = c3.slider("Current Harmonic Distortion (THDi %)", 5, 60, 35)
    
    nb_res = calculate_neutral_busbar_sizing(p_curr, nl_pct, thd_pct)
    st.divider()
    n1, n2, n3, n4 = st.columns([1, 1, 1.4, 1.6])
    n1.metric("Phase Line Current", f"{p_curr} A")
    n2.metric("Calculated Neutral Current", f"{nb_res['i_neutral']} A")
    n3.metric("Neutral Capacity Ratio", nb_res['rec_busbar_pct'])
    n4.metric("Recommended Circuit Breaker", nb_res['rec_breaker'])

elif menu == "FAT Quality Certificate":
    st.header("✅ Factory Acceptance Test (FAT) Certificate")
    c1, c2 = st.columns(2)
    s_no = c1.text_input("Serial No.", "AA-2026-0891"); client = c1.text_input("Customer Name", "Tata AutoComp"); proj = c1.text_input("Project", "Chakan Site")
    insp = c2.text_input("QA Inspector", "Sumit"); hv_pass = c2.checkbox("2.5kV HV Test Passed", value=True)
    m1, m2, m3 = st.columns(3)
    r = m1.number_input("R-Earth (MΩ)", 150.0); y = m2.number_input("Y-Earth (MΩ)", 145.0); b = m3.number_input("B-Earth (MΩ)", 160.0)
    
    with st.expander("🌡️ IEEE 43 Insulation Resistance (IR / Megger) Temperature Normalization to 40°C Baseline", expanded=False):
        ir_col1, ir_col2 = st.columns(2)
        test_t = ir_col1.slider("Site / Testing Ambient Temperature (°C)", 10.0, 55.0, 28.0)
        r_meas = ir_col2.number_input("Measured Megger IR Reading (MΩ)", value=r, step=10.0)
        ir_res = calculate_ieee43_ir_temp_correction(r_meas, test_t)
        ir1, ir2, ir3 = st.columns(3)
        ir1.metric("IEEE 43 Correction Factor (Kt)", ir_res['k_t'])
        ir2.metric("Normalized IR @ 40°C Baseline", f"{ir_res['r_40']} MΩ")
        ir3.metric("IEEE 43 Status", ir_res['status'])

    if st.button("📄 Generate FAT Certificate PDF"):
        pdf = generate_fat_certificate_pdf(s_no, proj, client, insp, r, y, b, hv_pass)
        st.download_button("📥 Download Official FAT QA Certificate", pdf, f"FAT_{s_no}.pdf", mime="application/pdf")

elif menu == "Dispatch Delivery Challan (DC)":
    st.header("🚚 Dispatch Delivery Challan (DC) & Gate Pass Generator")
    c1, c2 = st.columns(2)
    with c1:
        dc_no = st.text_input("Delivery Challan No.", f"AA/DC/{os.urandom(2).hex().upper()}")
        c_name = st.text_input("Consignee / Client Name", "Maharashtra Water Works Ltd")
        p_scope = st.text_input("Panel Scope Description", "22 kW VFD Control Panel (Siemens Drive)")
    with c2:
        v_no = st.text_input("Vehicle / LR Number", "MH 14 HG 4821")
        g_wt = st.number_input("Gross Weight (kg)", value=185.0, step=5.0)
        c_addr = st.text_area("Delivery Site Address", "Plot 42, MIDC Bhosari, Pune - 411026", height=68)

    if st.button("🚛 Generate Official Delivery Challan (PDF)"):
        dc_pdf = generate_delivery_challan_pdf(dc_no, c_name, c_addr, v_no, p_scope, g_wt)
        st.download_button("📥 Download Official Delivery Challan (PDF)", dc_pdf, f"Delivery_Challan_{dc_no.replace('/', '_')}.pdf", mime="application/pdf")

elif menu == "VFD Diagnostic Assistant":
    st.header("🛠️ Field VFD Diagnostic Assistant")
    brand = st.selectbox("VFD Brand", ["Siemens", "Danfoss", "Schneider", "ABB", "Delta"])
    code = st.text_input("Fault Code", placeholder="e.g. F0001, W013, Overcurrent, F0002")
    if st.button("🔍 Diagnose Fault"):
        if code:
            ans = query_local_ollama(f"Provide quick 4-step diagnostic checklist for {brand} VFD showing Fault: '{code}'.", format_json=False)
            if ans: st.markdown(ans)
            else: st.info(f"Ollama offline. **Standard Checklist for {brand} ({code}):**\n1. Check line voltage & DC bus.\n2. Megger motor winding.\n3. Disconnect load & test in scalar mode.")

elif menu == "Manage Price Database":
    st.header("⚙️ Local Price Database Manager")
    edited = st.data_editor(df_prices, num_rows="dynamic")
    if st.button("💾 Save Changes"):
        edited.to_excel(DB_FILE, index=False)
        st.cache_data.clear()
        st.success("Database updated successfully!")

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

        st.subheader("📲 One-Click Customer Follow-up Reminder")
        pending_clients = q_df[q_df["Status"] == "Pending"]["Client"].tolist() if "Client" in q_df.columns and "Status" in q_df.columns else []
        if pending_clients:
            sel_c = st.selectbox("Select Pending Client to Follow Up", pending_clients)
            c_row = q_df[q_df["Client"] == sel_c].iloc[0]
            fu_msg = f"Dear {sel_c},\n\nGreetings from Aryavarta Automation! Following up on our quotation for {c_row.get('Scope','Control Panel')} (Quote Value: ₹ {c_row.get('Quote_Price_INR',0):,.2f}). Please let us know if you need any technical clarifications or commercial adjustments.\n\nBest Regards,\nAryavarta Automation Sales Team"
            st.text_area("Follow-up Draft (Copy for Email / WhatsApp):", fu_msg, height=120)
        else:
            st.info("No pending quotes needing follow-up at the moment!")
    else:
        st.info("No quotes logged yet. Create a quote in 'Create Panel Quote' and click 'Save Quote to Quotation Register'.")
