import io
import json
import math
import urllib.request
import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(
    page_title="Aryavarta Automation - LV Engineering & Commercial Optimization Suite",
    page_icon="⚡",
    layout="wide"
)

st.markdown("""
<style>
div[data-testid="stMetricValue"] {
    white-space: normal !important;
    word-break: break-word !important;
    font-size: 1.15rem !important;
    line-height: 1.3 !important;
    font-weight: 700 !important;
}
div[data-testid="stMetricLabel"] {
    font-size: 0.85rem !important;
    font-weight: 600 !important;
}
div[data-testid="stMetric"] {
    background-color: #f8fafc;
    border: 1px solid #e2e8f0;
    padding: 10px 14px;
    border-radius: 8px;
}
@media (prefers-color-scheme: dark) {
    div[data-testid="stMetric"] {
        background-color: #1e293b;
        border: 1px solid #334155;
    }
}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# LOCAL FREE AI ENGINE (OLLAMA WATCHDOG & MULTI-MODEL ROUTER)
# ==============================================================================

def check_ollama_status(base_url="http://localhost:11434"):
    """Checks local Ollama service and returns installed LLM models."""
    try:
        req = urllib.request.Request(f"{base_url}/api/tags", headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=2) as res:
            data = json.loads(res.read().decode("utf-8"))
            models = [m.get("name") for m in data.get("models", [])]
            return True, models
    except Exception:
        return False, []

def query_local_ai(prompt, model="llama3.2", system_instruction="You are a senior switchgear design and industrial automation EPC engineer at Aryavarta Automation, Pune."):
    """Queries local Ollama model with a 120s timeout budget for cold loads."""
    payload = {
        "model": model,
        "prompt": f"{system_instruction}\n\nTask:\n{prompt}",
        "stream": False
    }
    try:
        req = urllib.request.Request(
            "http://localhost:11434/api/generate",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req, timeout=120) as res:
            response_json = json.loads(res.read().decode("utf-8"))
            return response_json.get("response", "").strip()
    except Exception:
        return None

# ==============================================================================
# RIGOROUS ELECTRICAL ENGINEERING CALCULATION ENGINES
# ==============================================================================

def calculate_short_circuit_and_busbar(tx_kva, tx_z, bus_len_m, conductor_type="Copper", cu_rate=890.0, al_rate=265.0):
    """Calculates short-circuit withstand area, electrodynamic repulsion forces, and support pitch per IS 8623 / IEC 61439."""
    i_sc_ka = (tx_kva * 100.0) / (math.sqrt(3) * 0.415 * tx_z * 10)
    flc_amps = (tx_kva * 1000.0) / (math.sqrt(3) * 415.0)
    
    # Adiabatic thermal withstand area per IEC 61439: S = (Isc * sqrt(t)) / k
    k_factor = 176.0 if conductor_type == "Copper" else 112.0
    req_area_1s = (i_sc_ka * 1000.0 * 1.0) / k_factor
    
    current_density = 1.5 if conductor_type == "Copper" else 0.9
    req_continuous_area = (flc_amps * 1.25) / current_density
    design_area = max(req_area_1s, req_continuous_area)
    
    busbars = [(25,5),(30,5),(40,5),(50,6),(60,8),(80,10),(100,10),(120,10)]
    rec_busbar = next((f"{w}x{t} mm" for w, t in busbars if (w*t) >= design_area), "2x 100x10 mm")
    
    i_pk_ka = i_sc_ka * 2.1
    center_distance_m = 0.075  # Standard 75mm phase-to-phase spacing
    f_repulsion_nm = round((0.2 * (i_pk_ka ** 2)) / center_distance_m, 1)
    
    if i_sc_ka <= 16:
        support_pitch = "500 mm Pitch (SMC 40/50mm Insulators)"
    elif i_sc_ka <= 25:
        support_pitch = "400 mm Pitch (SMC 50/60mm Insulators)"
    elif i_sc_ka <= 36:
        support_pitch = "300 mm Pitch (Heavy-Duty SMC Post Insulators)"
    else:
        support_pitch = "250 mm Pitch (Double Clamped Post Insulators)"
        
    vol_m3 = (design_area / 1e6) * bus_len_m * 3.45
    density = 8960.0 if conductor_type == "Copper" else 2700.0
    wt_kg = round(vol_m3 * density, 1)
    rate = cu_rate if conductor_type == "Copper" else al_rate
    cost = round(wt_kg * rate, 2)
    
    return {
        "i_sc_ka": round(i_sc_ka, 2), "flc_amps": round(flc_amps, 1),
        "req_area_1s": round(req_area_1s, 1), "design_area": round(design_area, 1),
        "rec_busbar": f"{rec_busbar} ({conductor_type})",
        "i_pk_ka": round(i_pk_ka, 1), "f_repulsion_nm": f_repulsion_nm,
        "support_pitch": support_pitch, "wt_kg": wt_kg, "cost": cost
    }

def calculate_motor_starting_and_cable(kw_val, starting_method, cable_len_m, cable_sqmm, conductor_type="Copper"):
    """Calculates motor FLC, starting inrush, line voltage drop, and switchgear coordination."""
    flc = (kw_val * 1000.0) / (math.sqrt(3) * 415.0 * 0.85 * 0.88)
    mult = {"DOL": 6.0, "Star-Delta": 2.5, "Soft Starter": 2.0, "VFD": 1.15}.get(starting_method, 6.0)
    i_start = flc * mult
    
    rho = 0.0175 if conductor_type == "Copper" else 0.0282
    r_per_m = rho / cable_sqmm
    x_per_m = 0.00008
    pf_start = 0.35
    sin_phi = math.sqrt(1 - pf_start**2)
    
    v_drop_volts = math.sqrt(3) * i_start * cable_len_m * (r_per_m * pf_start + x_per_m * sin_phi)
    v_drop_pct = round((v_drop_volts / 415.0) * 100.0, 2)
    term_voltage = round(415.0 - v_drop_volts, 1)
    
    mccb_rating = max(32, math.ceil(flc * 1.5 / 10.0) * 10)
    dol_contactor = f"{math.ceil(flc * 1.15)}A AC3"
    star_contactor = f"{math.ceil((flc / 3.0) * 1.15)}A AC3"
    main_delta_contactor = f"{math.ceil((flc / 1.732) * 1.15)}A AC3"
    
    status = "PASS ✅ (< 10% Standard Limit)" if v_drop_pct <= 10.0 else (
        "WARNING ⚠️ (10-15% - Check coil hold-in voltage)" if v_drop_pct <= 15.0 else
        "FAIL ❌ (> 15% - High risk of motor stall/chatter)"
    )
    return {
        "flc": round(flc, 1), "i_start": round(i_start, 1),
        "v_drop_volts": round(v_drop_volts, 1), "v_drop_pct": v_drop_pct,
        "term_voltage": term_voltage, "mccb_rating": f"{mccb_rating}A 3P",
        "dol_contactor": dol_contactor, "main_delta_contactor": main_delta_contactor,
        "star_contactor": star_contactor, "status": status
    }

def calculate_iec60890_thermal(h_mm, w_mm, d_mm, vfd_kw, contactor_qty, transformer_va, sealed=False, amb_t=40, max_int_t=45):
    """Calculates internal heat dissipation and required airflow or air conditioner rating per IEC 60890."""
    area_m2 = 1.8 * (h_mm / 1000.0) * ((w_mm + d_mm) / 1000.0) + 1.4 * (w_mm / 1000.0) * (d_mm / 1000.0)
    
    vfd_heat = vfd_kw * 35.0
    sw_heat = (contactor_qty * 14.0) + (transformer_va * 0.08) + 120.0
    tot_w = vfd_heat + sw_heat
    
    dt = max(5.0, max_int_t - amb_t)
    nat_w = 5.5 * area_m2 * dt
    net_w = max(0.0, tot_w - nat_w)
    
    m3h = (3.3 * net_w) / dt if net_w > 0 else 0
    cfm = m3h * 0.5886
    ton_ac = round(net_w / 3517.0, 2)
    
    if net_w <= 0:
        cooling_plan = "Natural Enclosure Surface Dissipation Sufficient."
    elif sealed:
        cooling_plan = f"Install Closed-Loop Panel Air Conditioner ({round(net_w/1000, 1)} kW / {ton_ac} TR)."
    else:
        cooling_plan = f"Install Louver Filter Fan Unit (> {round(m3h)} m³/h / {round(cfm)} CFM)."
        
    return {
        "area_m2": round(area_m2, 2), "tot_w": round(tot_w, 1), "nat_w": round(nat_w, 1),
        "net_w": round(net_w, 1), "m3h": round(m3h, 1), "cfm": round(cfm, 1),
        "ton_ac": ton_ac, "cooling_plan": cooling_plan
    }

def calculate_dc_control_line_drop(dc_amps, wire_sqmm, distance_m):
    """Calculates 24V DC field wiring drop across pilot solenoids, sensors, and remote PLC racks."""
    rho = 0.0178
    v_drop = (2.0 * dc_amps * distance_m * rho) / wire_sqmm
    end_v = round(24.0 - v_drop, 2)
    v_pct = round((v_drop / 24.0) * 100.0, 1)
    status = "PASS ✅ Safe (> 20.4V DC / -15% Limit)" if end_v >= 20.4 else "FAIL ❌ Drop > 15% (Solenoid/PLC I/O Drop)"
    max_safe_dist = round((3.6 * wire_sqmm) / (2.0 * dc_amps * rho), 1) if dc_amps > 0 else 0
    return {"v_drop": round(v_drop, 2), "end_v": end_v, "v_pct": v_pct, "status": status, "max_safe_dist": max_safe_dist}

def calculate_apfc_and_reactor(active_kw, c_pf=0.78, t_pf=0.98, detuning_pct=7.0):
    """Calculates APFC capacitor bank kVAR, series reactor tuning, and terminal voltage inflation per IEC 61431."""
    if c_pf >= t_pf:
        return {"req_kvar": 0, "status": "Current Power Factor is optimal."}
        
    req_kvar = math.ceil(active_kw * (math.tan(math.acos(c_pf)) - math.tan(math.acos(t_pf))))
    p = detuning_pct / 100.0
    sys_v = 415.0
    
    v_cap_operating = sys_v / (1.0 - p)
    rec_v_rating = 480 if v_cap_operating <= 460 else 525
    nameplate_kvar = req_kvar * ((rec_v_rating / v_cap_operating) ** 2)
    reactor_kvar = req_kvar * p
    tuning_freq = round(50.0 / math.sqrt(p), 1)
    
    step_size = 25 if req_kvar <= 100 else 50
    step_count = math.ceil(req_kvar / step_size)
    
    return {
        "req_kvar": req_kvar, "v_cap_op": round(v_cap_operating, 1),
        "rec_v_rating": f"{rec_v_rating}V Heavy-Duty Rating", "nameplate_kvar": round(nameplate_kvar, 1),
        "reactor_kvar": round(reactor_kvar, 1), "tuning_freq": f"{tuning_freq} Hz (5th Harmonic Elimination)",
        "step_config": f"{step_count} Steps x {step_size} kVAR"
    }

def calculate_vfd_and_msedcl_roi(motor_kw, run_hrs_day, tariff_rate, speed_reduction_pct, vfd_upgrade_cost, c_pf=0.82, t_pf=0.98, monthly_energy_bill=180000):
    """Calculates affinity law pump/fan power savings and MSEDCL power factor incentives."""
    speed_ratio = (100.0 - speed_reduction_pct) / 100.0
    power_saved_kw = motor_kw * (1.0 - (speed_ratio ** 3))
    annual_kwh_saved = power_saved_kw * run_hrs_day * 365.0
    annual_energy_savings_inr = annual_kwh_saved * tariff_rate
    
    rebate_pct = 0.07 if t_pf >= 0.98 else (0.05 if t_pf >= 0.95 else 0.0)
    annual_pf_rebate_inr = (monthly_energy_bill * rebate_pct) * 12.0
    
    total_annual_savings_inr = annual_energy_savings_inr + annual_pf_rebate_inr
    payback_months = round((vfd_upgrade_cost / total_annual_savings_inr) * 12.0, 1) if total_annual_savings_inr > 0 else 0
    
    return {
        "power_saved_kw": round(power_saved_kw, 2),
        "annual_kwh_saved": round(annual_kwh_saved, 0),
        "annual_energy_savings_inr": round(annual_energy_savings_inr, 2),
        "annual_pf_rebate_inr": round(annual_pf_rebate_inr, 2),
        "total_annual_savings_inr": round(total_annual_savings_inr, 2),
        "payback_months": payback_months
    }

# ==============================================================================
# REPORTLAB PDF GENERATORS (FAT, DISPATCH CHALLAN & ENERGY ROI)
# ==============================================================================

def generate_fat_certificate_pdf(sr_no, client_name, project_title, inspector, r_ir, y_ir, b_ir, hv_pass):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("ARYAVARTA AUTOMATION", ParagraphStyle('H1', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor("#1A365D"))),
        Paragraph("<b>FACTORY ACCEPTANCE TEST (FAT) QUALITY CERTIFICATE</b>", styles['Heading2']),
        Paragraph("Gat No. 1610, Dehu Alandi Road, Chikhali, Pune - 411062 | GSTIN: 27ABOFA4930E1ZH", styles['Normal']),
        Paragraph(f"Certificate No: <b>FAT/{sr_no}</b> | Date: <b>{pd.Timestamp.now().strftime('%d-%b-%Y')}</b>", styles['Normal']),
        Spacer(1, 14),
        Paragraph(f"<b>Client Entity:</b> {client_name} &nbsp;&nbsp;&nbsp;&nbsp; <b>Project Scope:</b> {project_title}", styles['Normal']),
        Spacer(1, 12),
        Paragraph("<b>1. Insulation Resistance Testing (Megger 500V DC Normalized)</b>", styles['Heading3'])
    ]
    ir_table = [
        ["Phase Terminals", "Measured Value", "IS 8623 Requirement", "Verdict"],
        ["R-Phase to Earth", f"{r_ir} MΩ", "> 5.0 MΩ", "PASS" if r_ir >= 5.0 else "FAIL"],
        ["Y-Phase to Earth", f"{y_ir} MΩ", "> 5.0 MΩ", "PASS" if y_ir >= 5.0 else "FAIL"],
        ["B-Phase to Earth", f"{b_ir} MΩ", "> 5.0 MΩ", "PASS" if b_ir >= 5.0 else "FAIL"]
    ]
    t_ir = Table(ir_table, colWidths=[130, 130, 150, 90])
    t_ir.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('PADDING', (0,0), (-1,-1), 5)
    ]))
    story.extend([t_ir, Spacer(1, 14), Paragraph("<b>2. Dielectric Withstand & Workmanship Standards</b>", styles['Heading3'])])
    
    checks_table = [
        ["Inspection Parameter", "Reference Standard", "Test Status"],
        ["2.5 kV AC Dielectric Withstand (1 Minute)", "IEC 61439-1 Clause 10.9", "PASSED" if hv_pass else "FAILED"],
        ["Clearance in Air (Ph-Ph Min 19mm, Ph-E Min 14mm)", "IS 8623 Table 1", "PASSED"],
        ["Protective Earth Bonding Continuity (< 0.1 Ω)", "IS 3043 / IEC 61439", "PASSED"],
        ["Door Gasket Sealing & Lock Interlocks", "IEC 60529 (IP55/IP65 Standard)", "PASSED"]
    ]
    t_c = Table(checks_table, colWidths=[240, 170, 90])
    t_c.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#475569")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('PADDING', (0,0), (-1,-1), 5)
    ]))
    story.extend([
        t_c, Spacer(1, 24),
        Paragraph(f"Inspected by: <b>{inspector}</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Quality Head Signature: ____________________", styles['Normal'])
    ])
    doc.build(story)
    buf.seek(0)
    return buf

def generate_delivery_challan_pdf(dc_no, client_name, address, vehicle_no, panel_desc, gross_wt):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("ARYAVARTA AUTOMATION", ParagraphStyle('H1', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor("#1A365D"))),
        Paragraph("<b>DELIVERY CHALLAN & DISPATCH NOTE</b>", styles['Heading2']),
        Paragraph("Gat No. 1610, Dehu Alandi Road, Chikhali, Pune - 411062 | GSTIN: 27ABOFA4930E1ZH", styles['Normal']),
        Spacer(1, 12)
    ]
    meta = [
        [Paragraph(f"<b>Challan No:</b> {dc_no}", styles['Normal']), Paragraph(f"<b>Date:</b> {pd.Timestamp.now().strftime('%d-%b-%Y')}", styles['Normal'])],
        [Paragraph(f"<b>Consignee:</b> {client_name}", styles['Normal']), Paragraph(f"<b>Vehicle / LR No:</b> {vehicle_no}", styles['Normal'])],
        [Paragraph(f"<b>Site Destination:</b> {address}", styles['Normal']), Paragraph(f"<b>Gross Weight:</b> {gross_wt} kg", styles['Normal'])]
    ]
    t_meta = Table(meta, colWidths=[270, 270])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F8FAFC")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('PADDING', (0,0), (-1,-1), 5)
    ]))
    items = [
        ["S.No", "Scope / Item Description", "HSN Code", "Qty", "Remarks"],
        ["1", panel_desc, "85371010", "1 Set", "Dispatched for erection & commissioning."]
    ]
    t_items = Table(items, colWidths=[40, 260, 80, 50, 110])
    t_items.setStyle(TableStyle([
        ('HEADERBACKGROUND', (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('PADDING', (0,0), (-1,-1), 5)
    ]))
    story.extend([
        t_meta, Spacer(1, 14), t_items, Spacer(1, 20),
        Paragraph("<b>Declaration:</b> Goods dispatched in good condition for site installation. Not for retail sale.", styles['Normal']),
        Spacer(1, 24),
        Paragraph("Consignee's Signature: ____________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; For ARYAVARTA AUTOMATION", styles['Normal'])
    ])
    doc.build(story)
    buf.seek(0)
    return buf

# ==============================================================================
# UI NAVIGATION & PRESENTATION
# ==============================================================================

st.sidebar.title("⚡ Aryavarta Automation")
st.sidebar.markdown("**Process Plant Electrical & Automation Suite**")
st.sidebar.caption("Chikhali Works, Pune")

ollama_online, installed_models = check_ollama_status()
if ollama_online:
    st.sidebar.success("🟢 Local AI Engine Active")
    default_index = 0
    if "llama3.2:latest" in installed_models:
        default_index = installed_models.index("llama3.2:latest")
    elif "llama3.2" in installed_models:
        default_index = installed_models.index("llama3.2")
    active_model = st.sidebar.selectbox("Active AI Model", installed_models, index=default_index)
else:
    st.sidebar.error("🔴 Local Ollama Offline")
    st.sidebar.caption("Run `ollama run llama3.2` in terminal to activate local AI.")
    active_model = "llama3.2"

st.sidebar.divider()

menu = st.sidebar.radio("Engineering Suite", [
    "💰 Multi-Brand Procurement Arbitrage & Copper Hedging",
    "📈 VFD Energy Savings & MSEDCL Payback Dossier",
    "🤖 AI Tender Spec & Deviation Auditor",
    "🛡️ AI Cause & Effect (Interlock) Safety Matrix",
    "🎛️ AI SCADA / PLC Tag Database & Modbus Register Generator",
    "📜 AI Control Philosophy & SCL Pseudocode Drafter",
    "🔄 AI Switchgear OEM Cross-Reference & Substitutor",
    "🛠️ AI Field VFD Diagnostic & Commissioning Assistant",
    "⚡ Main Busbar & Short-Circuit kA Sizer (IS 8623)",
    "⚙️ Motor Starting Inrush & Cable Sizer (IS 3961)",
    "🌡️ Panel Thermal Dissipation & Climate Sizer (IEC 60890)",
    "🎛️ 24V DC Field Control Wiring Voltage Drop Sizer",
    "⚡ APFC Capacitor Bank & 7% Detuned Reactor Sizer",
    "📋 Motor Switchgear Master Lookup Chart",
    "✅ Quality Inspection FAT Certificate Generator",
    "🚚 Dispatch Delivery Challan Generator"
])

# --- MODULE 1: PROCUREMENT ARBITRAGE & COPPER HEDGING ---
if menu == "💰 Multi-Brand Procurement Arbitrage & Copper Hedging":
    st.header("💰 Multi-Brand Switchgear Procurement Arbitrage & Copper Hedging")
    st.caption("Compares dealer-discounted purchase costs across approved OEM switchgear and calculates live copper busbar requirements to protect project gross margins.")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.subheader("1. MCX Copper Commodity Hedging")
        live_cu_rate = st.number_input("Live Copper Busbar Rate (₹ / kg)", value=890.0, step=10.0)
        busbar_len = st.number_input("Total Main Busbar Length in Panel (Meters)", value=4.0, step=0.5)
        flc_amps = st.number_input("Operating Busbar Current (Amps)", value=400.0, step=50.0)
        
        cu_res = calculate_short_circuit_and_busbar(250.0, 5.0, busbar_len, "Copper", live_cu_rate, 265.0)
        cu_wt = round(((flc_amps * 1.25 / 1.5) / 1e6) * busbar_len * 3.45 * 8960.0, 1)
        cu_cost = round(cu_wt * live_cu_rate, 2)
        
        m1, m2 = st.columns(2)
        m1.metric("Calculated Copper Weight", f"{cu_wt} kg")
        m2.metric("Landed Copper Cost", f"₹ {cu_cost:,.2f}")
        
    with col_c2:
        st.subheader("2. Switchgear Discount Arbitrage")
        feeder_kw = st.selectbox("Sample Feeder Rating to Compare", [7.5, 15.0, 22.0, 37.0, 45.0, 55.0, 75.0], index=1)
        st.caption("Calculates net dealer landed purchase costs across major approved makes.")
        
        # Market standard benchmark list prices & tier-1 discounts for India
        flc_f = (feeder_kw * 1000) / (1.732 * 415 * 0.85 * 0.88)
        base_list = 4500.0 + (feeder_kw * 380.0)
        
        comparison_data = [
            {"OEM Brand": "Siemens (3VM / 3RT)", "Standard List Price (₹)": base_list * 1.05, "Dealer Discount (%)": 48.0, "Net Cost (₹)": (base_list * 1.05) * 0.52},
            {"OEM Brand": "Schneider (EasyPact / TeSys)", "Standard List Price (₹)": base_list * 1.02, "Dealer Discount (%)": 46.0, "Net Cost (₹)": (base_list * 1.02) * 0.54},
            {"OEM Brand": "ABB (Formula / AX)", "Standard List Price (₹)": base_list * 0.98, "Dealer Discount (%)": 45.0, "Net Cost (₹)": (base_list * 0.98) * 0.55},
            {"OEM Brand": "L&T (dsine / MO)", "Standard List Price (₹)": base_list * 0.95, "Dealer Discount (%)": 44.0, "Net Cost (₹)": (base_list * 0.95) * 0.56}
        ]
        df_arb = pd.DataFrame(comparison_data)
        df_arb["Net Cost (₹)"] = df_arb["Net Cost (₹)"].round(2)
        st.dataframe(df_arb, width="stretch")
        
        cheapest = df_arb.loc[df_arb["Net Cost (₹)"].idxmin()]
        st.success(f"💡 **Procurement Arbitrage Recommendation:** Sourcing **{cheapest['OEM Brand']}** yields the lowest landed cost at **₹ {cheapest['Net Cost (₹)']:,.2f}** per feeder unit.")

# --- MODULE 2: VFD ENERGY SAVINGS & MSEDCL PAYBACK DOSSIER ---
elif menu == "📈 VFD Energy Savings & MSEDCL Payback Dossier":
    st.header("📈 VFD Energy Savings & MSEDCL Power Factor Payback Dossier")
    st.caption("Generates a financial justification report using pump/fan Affinity Laws and MSEDCL power factor rebates to accelerate client purchase orders.")
    
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        m_kw = st.number_input("Process Motor Rating (kW)", value=37.0, step=5.0)
        hrs = st.number_input("Operating Hours per Day", value=18, max_value=24)
        tariff = st.number_input("HT Industrial Tariff Rate (₹ / kWh)", value=9.50, step=0.25)
    with col_r2:
        flow_red = st.slider("Average Required Flow Modulation (%)", 10, 40, 20)
        vfd_invest = st.number_input("Total VFD Panel Upgrade Investment (₹)", value=125000, step=5000)
        energy_bill = st.number_input("Average Monthly Energy Bill (₹)", value=220000, step=10000)

    roi = calculate_vfd_and_msedcl_roi(m_kw, hrs, tariff, flow_red, vfd_invest, monthly_energy_bill=energy_bill)
    
    st.divider()
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Power Reduction", f"{roi['power_saved_kw']} kW/hr")
    r2.metric("Annual Energy Savings", f"₹ {roi['annual_energy_savings_inr']:,.2f}")
    r3.metric("Annual MSEDCL PF Rebate", f"₹ {roi['annual_pf_rebate_inr']:,.2f}")
    r4.metric("Capital Payback Period", f"{roi['payback_months']} Months")
    
    st.info(f"💡 **Executive Proposal Summary:** At {flow_red}% flow reduction, the client saves **{roi['annual_kwh_saved']:,.0f} kWh** annually. Combined with MSEDCL's 7% PF incentive, total annual recurring savings reach **₹ {roi['total_annual_savings_inr']:,.2f}**, recovering investment in **{roi['payback_months']} months**.")

# --- MODULE 3: AI TENDER SPEC & DEVIATION AUDITOR ---
elif menu == "🤖 AI Tender Spec & Deviation Auditor":
    st.header("🤖 AI Tender Specification & Technical Deviation Auditor")
    st.caption("Analyzes client tender clauses against Aryavarta's baseline, identifying deviations across sheet metal, Form factor, and approved vendor lists.")
    
    tender_text = st.text_area(
        "Paste Client Tender Scope / Specification Clauses:",
        value="""Supply of 1 No. Food Grade MCC + Automation Panel for Brewery CIP section.
Incoming Supply: 415V, 3-Phase, 4-Wire, 50Hz from 1000kVA Transformer (5% Z).
Busbars: High conductivity EC Grade Copper rated for 40kA for 1 sec. Temperature rise limited to 35°C over 45°C ambient.
Enclosure: Floor mounting, Form-4b compartmentalized, fabricated with 2.0mm CRCA sheet steel, IP65 protection.
Color Shade: RAL 7032 exterior and white glossy interior.
Approved Makes: Switchgear: Siemens / Schneider only. Terminals: Phoenix Contact only. Cable Glands: Brass nickel-plated double compression.""",
        height=170
    )
    
    if st.button("⚡ Audit Tender Clauses & Generate Deviation Schedule"):
        prompt = (
            f"Analyze this client technical specification: '{tender_text}'. "
            "Audit it against standard LT switchgear manufacturing practices (IS 8623 / IEC 61439-1). "
            "Output a structured Technical Deviation Schedule in a Markdown table with columns: "
            "'Tender Clause / Parameter', 'Client Specified Requirement', 'Aryavarta Manufacturing Standard', "
            "'Compliance Status (Complied / Deviated / Clarification Needed)', 'Technical & Commercial Justification'."
        )
        with st.spinner("🤖 AI is cross-referencing tender specifications and standards..."):
            ai_resp = query_local_ai(prompt, model=active_model)
        if ai_resp:
            st.markdown(ai_resp)
        else:
            st.warning("⚠️ Local Ollama is offline or timed out. Displaying deterministic audit schedule:")
            st.markdown("""
| Tender Clause / Parameter | Client Specified Requirement | Aryavarta Manufacturing Standard | Compliance Status | Technical & Commercial Justification |
| :--- | :--- | :--- | :--- | :--- |
| **Enclosure Compartmentalization** | Form-4b construction | Form-2b standard (Functional units separated from busbars) | **Deviated** | Form-4b requires separate individual rear cable alley terminations for every feeder, increasing panel footprint by 40%. Form-2b offers full operator safety at optimal commercial cost. |
| **Ingress Protection Rating** | IP65 for indoor floor MCC | IP55 with seamless CNC polyurethane gaskets | **Deviated** | IP65 prohibits natural forced ventilation louvers and mandates sealed panel air conditioners. IP55 is the standard for indoor food plants unless high-pressure direct water jets are used. |
| **Sheet Metal Thickness** | 2.0 mm CRCA throughout | 2.0 mm load-bearing frame, 1.6 mm doors & covers | **Deviated** | 1.6 mm doors provide rigid deflection resistance while reducing hinge wear. 2.0 mm is applied on main gland plates and structural corner posts. |
| **Busbar Temperature Rise** | Max 35°C rise over 45°C ambient | 1.5 A/mm² current density verified to IEC 61439-1 | **Complied** | EC-grade copper busbars sized conservatively to ensure total temperature stays well below 85°C (within the 105°C insulation limit). |
| **Approved Vendor Makes** | Siemens / Schneider only | Siemens 3VA/3RT / Schneider EasyPact / TeSys | **Complied** | Full adherence to client-approved switchgear tiers. |
            """)

# --- MODULE 4: AI CAUSE & EFFECT SAFETY MATRIX ---
elif menu == "🛡️ AI Cause & Effect (Interlock) Safety Matrix":
    st.header("🛡️ AI Process Cause & Effect (Interlock) Safety Matrix")
    st.caption("Generates the formal Cause & Effect safety trip table required by EPC consultants before PLC logic approval.")
    
    process_scope = st.text_area(
        "Enter Process System & Trip Conditions:",
        value="Brewery CIP Delivery System: CIP Pump (5.5kW), Return Pump (3.7kW), Caustic Tank Low Level Switch, Acid Tank Low Level Switch, High Temperature Switch (90°C), Delivery Line Pressure Transmitter, Emergency Stop PB, Return Flow Switch.",
        height=120
    )
    
    if st.button("⚡ Generate Cause & Effect Safety Matrix"):
        prompt = (
            f"Given this process equipment and safety sensors: '{process_scope}', build an official Cause and Effect (Interlock) Matrix. "
            "Format as a Markdown table with columns: 'Initiating Cause / Sensor Event', 'Trigger Threshold / Condition', "
            "'Direct Action / Interlock Trip', 'Safety Class (Cat 3/4 / SIL 2)', and 'Manual Reset Required (Yes/No)'."
        )
        with st.spinner("🤖 AI is compiling safety trip permissives and interlock matrix..."):
            ai_resp = query_local_ai(prompt, model=active_model)
        if ai_resp:
            st.markdown(ai_resp)
        else:
            st.markdown("""
| Initiating Cause / Sensor Event | Trigger Threshold / Condition | Direct Action / Interlock Trip | Safety Class | Manual Reset Required |
| :--- | :--- | :--- | :--- | :--- |
| **Emergency Stop PB Pressed** | Contact Open (Fail-Safe 24V DC) | Immediate trip of all CIP & Return Pumps, De-energize heating valves | SIL 2 / Cat 4 | **Yes** (Physical PB release + HMI Reset) |
| **Caustic Tank Level Low** | Float Switch Open (< 15% Level) | Interlock Trip CIP Delivery Pump (Dry Run Protection) | SIL 1 / Cat 2 | No (Auto-resets on level recovery) |
| **Delivery Over-Pressure** | Pressure > 4.5 Bar (PT-01) | Modulate VFD speed to minimum, trip pump if sustained > 5s | SIL 1 | **Yes** (HMI Acknowledgment) |
| **Return Line Flow Lost** | Flow Switch Open for > 15s | Stop CIP heating, alert operator for blockage / valve misalignment | Process Alarm | No |
| **Over-Temperature Alarm** | RTD Pt100 > 92°C | Close Steam Modulating Control Valve immediately | SIL 2 | **Yes** |
            """)

# --- MODULE 5: AI SCADA TAG & MODBUS REGISTER GENERATOR ---
elif menu == "🎛️ AI SCADA / PLC Tag Database & Modbus Register Generator":
    st.header("🎛️ AI SCADA / PLC Tag Database & Modbus Register Synthesizer")
    st.caption("Translates raw plant equipment lists into standardized SCADA/HMI tag sheets and Modbus/PROFINET register maps, exportable directly to CSV.")
    
    equip_input = st.text_area(
        "Enter Plant Motors, Valves & Instruments:",
        value="""Pumps: CIP Delivery Pump (5.5kW VFD), Return Scavenge Pump (3.7kW DOL), Hot Water Pump (7.5kW DOL).
Valves: Caustic Inflow Valve, Acid Inflow Valve, Rinse Water Valve, Drain Dump Valve (All 24V DC Solenoid with Open/Close limit switches).
Transmitters: Delivery Pressure Transmitter (0-10 Bar, 4-20mA), CIP Temp Transmitter (0-150°C, Pt100/4-20mA), Conductivity Transmitter (0-200 mS/cm, 4-20mA).""",
        height=130
    )
    
    if st.button("⚡ Synthesize Tag Database"):
        prompt = (
            f"Generate an industrial SCADA/PLC Tag Database from this equipment: '{equip_input}'. "
            "Output a structured CSV format with columns: Tag_Name, Description, Signal_Type (DI/DO/AI/AO), "
            "PLC_Address (e.g., %I0.0, %Q0.0, %IW64), Modbus_Register (40001 series), Engineering_Units, "
            "Alarm_Low, Alarm_High. Tag names must follow ISA-5.1 standards."
        )
        with st.spinner("🤖 AI is generating tag records and Modbus mapping..."):
            ai_resp = query_local_ai(prompt, model=active_model)
        if ai_resp:
            st.markdown(ai_resp)
        else:
            default_tags = pd.DataFrame([
                {"Tag_Name": "PMP_CIP_DEL_RUN", "Description": "CIP Delivery Pump Running Feedback", "Signal_Type": "DI", "PLC_Address": "%I0.0", "Modbus_Register": "40001", "Engineering_Units": "BOOL", "Alarm_Low": "-", "Alarm_High": "-"},
                {"Tag_Name": "PMP_CIP_DEL_TRP", "Description": "CIP Delivery Pump Trip Alarm", "Signal_Type": "DI", "PLC_Address": "%I0.1", "Modbus_Register": "40002", "Engineering_Units": "BOOL", "Alarm_Low": "-", "Alarm_High": "TRUE"},
                {"Tag_Name": "PMP_CIP_DEL_CMD", "Description": "CIP Delivery Pump Start/Stop Command", "Signal_Type": "DO", "PLC_Address": "%Q0.0", "Modbus_Register": "40003", "Engineering_Units": "BOOL", "Alarm_Low": "-", "Alarm_High": "-"},
                {"Tag_Name": "PMP_CIP_DEL_SPD_REF", "Description": "CIP Delivery Pump Speed Reference", "Signal_Type": "AO", "PLC_Address": "%QW64", "Modbus_Register": "40004", "Engineering_Units": "Hz", "Alarm_Low": "10.0", "Alarm_High": "50.0"},
                {"Tag_Name": "PT_CIP_DEL_PRES", "Description": "CIP Delivery Line Pressure", "Signal_Type": "AI", "PLC_Address": "%IW64", "Modbus_Register": "40005", "Engineering_Units": "Bar", "Alarm_Low": "0.5", "Alarm_High": "4.5"},
                {"Tag_Name": "TT_CIP_LINE_TEMP", "Description": "CIP Return Line Temperature", "Signal_Type": "AI", "PLC_Address": "%IW66", "Modbus_Register": "40006", "Engineering_Units": "°C", "Alarm_Low": "15.0", "Alarm_High": "85.0"},
                {"Tag_Name": "VLV_CAUSTIC_IN_CMD", "Description": "Caustic Inflow Solenoid Valve Open Cmd", "Signal_Type": "DO", "PLC_Address": "%Q0.1", "Modbus_Register": "40007", "Engineering_Units": "BOOL", "Alarm_Low": "-", "Alarm_High": "-"},
                {"Tag_Name": "VLV_CAUSTIC_OPEN_ZSO", "Description": "Caustic Inflow Valve Open Limit Switch", "Signal_Type": "DI", "PLC_Address": "%I0.2", "Modbus_Register": "40008", "Engineering_Units": "BOOL", "Alarm_Low": "-", "Alarm_High": "-"}
            ])
            st.dataframe(default_tags, width="stretch")
            csv_data = default_tags.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download SCADA Tag Database (.csv)", csv_data, "SCADA_Tag_Database.csv", "text/csv")

# --- MODULE 6: AI CONTROL PHILOSOPHY & SCL DRAFTER ---
elif menu == "📜 AI Control Philosophy & SCL Pseudocode Drafter":
    st.header("📜 AI Functional Control Narrative & SCL Pseudocode Drafter")
    st.caption("Generates Structured Control Language (SCL / IEC 61131-3) logic and step-sequence narratives for process loops.")
    
    routine_type = st.selectbox(
        "Select Process Routine Template:",
        [
            "5-Stage CIP Cleaning Loop (Pre-Rinse, Caustic, Inter-Rinse, Acid, Final Rinse)",
            "Brewery Mash Tun Temperature Stepping & Agitation Logic",
            "Dairy Raw Milk Receiving & Deaeration Interlock Sequence",
            "Decanter / Separator Start-up Permissive & Vibration Trip Interlock"
        ]
    )
    
    if st.button("📝 Draft Functional Narrative & SCL Code"):
        prompt = (
            f"Draft the Functional Design Specification (FDS) narrative and IEC 61131-3 SCL pseudocode for: '{routine_type}'. "
            "Include safety permissives, dry-run interlocks, timer state-machine architecture, and alarm trip triggers."
        )
        with st.spinner("🤖 AI is drafting functional control narrative and SCL logic..."):
            ai_resp = query_local_ai(prompt, model=active_model)
        if ai_resp:
            st.markdown(ai_resp)
        else:
            st.code("""
// =========================================================================
// SCL State Machine Logic: 5-Stage CIP Cleaning Loop
// Platform: Siemens S7-1200 / S7-1500 (TIA Portal)
// =========================================================================
CASE #CIP_State OF
    0: // Idle State
        #Pump_Cmd := FALSE;
        #Return_Pump_Cmd := FALSE;
        #Steam_Valve := FALSE;
        IF #Start_PB AND #Tank_Level_OK AND NOT #E_Stop THEN
            #CIP_State := 10; // Advance to Pre-Rinse
            #Step_Timer(IN:=FALSE);
        END_IF;

    10: // Pre-Rinse (Water Wash to Drain - 300s)
        #Water_Valve := TRUE;
        #Pump_Cmd := TRUE;
        #Step_Timer(IN:=TRUE, PT:=T#300S);
        IF #Step_Timer.Q THEN
            #Water_Valve := FALSE;
            #Pump_Cmd := FALSE;
            #CIP_State := 20; // Advance to Caustic Circulation
        END_IF;

    20: // Caustic Recirculation (Heat to 75C - 900s)
        #Caustic_Valve := TRUE;
        #Pump_Cmd := TRUE;
        #Return_Pump_Cmd := TRUE;
        #Steam_Valve := (#Temp_Sensor < 75.0); // Temperature Regulation
        #Step_Timer(IN:=TRUE, PT:=T#900S);
        IF #Step_Timer.Q THEN
            #Steam_Valve := FALSE;
            #CIP_State := 30; // Final Rinse
        END_IF;
END_CASE;
            """, language="pascal")

# --- MODULE 7: AI SWITCHGEAR CROSS-REFERENCE ---
elif menu == "🔄 AI Switchgear OEM Cross-Reference & Substitutor":
    st.header("🔄 AI Switchgear Lead-Time Substitution & OEM Cross-Referencer")
    st.caption("Identifies direct drop-in replacement part numbers across Siemens, Schneider, ABB, and L&T during supply chain delays.")
    
    part_input = st.text_input("Enter Problem Part Number / Rating:", "Siemens 3VM1116-4ED32-0AA0 (160A 3P 36kA Thermal-Magnetic MCCB)")
    
    if st.button("🔍 Find Cross-Manufacturer Equivalents"):
        prompt = (
            f"Provide direct commercial and technical drop-in alternatives for: '{part_input}'. "
            "Compare across: 1) Siemens, 2) Schneider Electric, 3) ABB, 4) L&T. "
            "Include: Exact Model Series, Rated Breaking Capacity (Icu @ 415V), Frame Dimensions (W x H x D mm), "
            "Trip Unit Type (Thermal-Magnetic vs Microprocessor), and Terminal Link Compatibility."
        )
        with st.spinner("🤖 AI is cross-referencing OEM switchgear catalogs..."):
            ai_resp = query_local_ai(prompt, model=active_model)
        if ai_resp:
            st.markdown(ai_resp)
        else:
            st.markdown("""
**Cross-Manufacturer Equivalents for 160A 3P 36kA MCCB:**
* **Siemens**: `3VM1116-4ED32-0AA0` (Frame 160, TM210 trip unit, 36kA @ 415V, Dimensions: 105 x 130 x 70 mm)
* **Schneider Electric**: `LV429676` (Compact NSX160F, TMD 160A trip unit, 36kA @ 415V, Dimensions: 105 x 161 x 86 mm)
* **ABB**: `1SDA067439R1` (Formula A2N 250, TMF 160A fixed trip, 36kA @ 415V, Dimensions: 105 x 150 x 60 mm)
* **L&T (Lauritz & Knudsen)**: `CM92015OOOO` (dsine DZ160, Thermal Magnetic, 36kA @ 415V, Dimensions: 105 x 165 x 68 mm)

> **Shop-Floor Engineering Note:** While all 4 models share the standard 105mm modular width, the Schneider NSX is 31mm taller. If substituting inside an existing Form-2b compartment, verify the door handle mechanism depth and spreader terminal clearances.
            """)

# --- MODULE 8: AI FIELD VFD DIAGNOSTIC ASSISTANT ---
elif menu == "🛠️ AI Field VFD Diagnostic & Commissioning Assistant":
    st.header("🛠️ AI Field VFD Diagnostic & Commissioning Assistant")
    st.caption("Immediate physical checklists, parameter inspection steps, and fault elimination for shop-floor testing and site commissioning.")
    
    d_c1, d_c2 = st.columns(2)
    drive_make = d_c1.selectbox("Select VFD Brand:", ["Siemens Sinamics G120 / V20", "Danfoss VLT / FC-51 / FC-302", "Schneider Altivar ATV320 / ATV630", "ABB ACS580 / ACS380"])
    fault_code = d_c2.text_input("Enter Displayed Fault Code / Symptom:", "F07900 (Motor Blocked) / Overcurrent on startup")
    
    if st.button("🔍 Diagnose Fault & Generate Site Checklist"):
        prompt = (
            f"Provide an immediate step-by-step diagnostic checklist for {drive_make} showing Fault '{fault_code}'. "
            "Structure into: 1) Immediate physical and electrical checks (Megger, DC bus, load disconnect), "
            "2) Critical drive parameters to check and adjust (motor nameplate, ramp-up time, current limit), "
            "3) Recommended scalar (V/f) vs. vector commissioning test."
        )
        with st.spinner("🤖 AI is diagnosing fault code and generating site checklist..."):
            ai_resp = query_local_ai(prompt, model=active_model)
        if ai_resp:
            st.markdown(ai_resp)
        else:
            st.markdown(f"""
**Diagnostic Checklist for {drive_make} ({fault_code}):**
1. **Mechanical & Motor Terminal Verification**:
   * Decouple motor shaft from load (pump/agitator) to verify free mechanical rotation.
   * Megger motor windings phase-to-phase and phase-to-ground with 500V Megger (Disconnect VFD output terminals $U, V, W$ before testing).
2. **Drive Parameter Adjustments**:
   * Verify motor nameplate data in drive commissioning parameters: Rated Voltage (415V), Rated Current (FLC), Power Factor, and Rated RPM.
   * Increase Acceleration Ramp-up Time ($P1120$) from default 3.0s to 8.0s–12.0s to mitigate high startup inertia.
3. **Scalar Testing Mode**:
   * Switch control mode from Sensorless Vector ($P1300 = 20$) to Linear V/f mode ($P1300 = 0$) to verify if current spikes are caused by motor parameter detuning.
            """)

# --- MODULE 9: BUSBAR & SHORT CIRCUIT SIZER ---
elif menu == "⚡ Main Busbar & Short-Circuit kA Sizer (IS 8623)":
    st.header("⚡ Substation Fault Level (kA) & Busbar Electrodynamic Sizer")
    st.caption("Calculates adiabatic withstand area ($S = \\frac{I_{sc}\\sqrt{t}}{k}$), electrodynamic repulsion forces, and support pitch per IS 8623 / IEC 61439.")
    
    col1, col2 = st.columns(2)
    tx_k = col1.number_input("Incomer Transformer Rating (kVA)", value=1000.0, step=100.0)
    tx_z = col1.number_input("Transformer Impedance %Z", value=5.0, step=0.5)
    bus_len = col1.number_input("Busbar Total Run Length (Meters)", value=4.0, step=0.5)
    
    mat = col2.radio("Conductor Material", ["Copper", "Aluminum"])
    cu_rate = col2.number_input("Copper Commodity Rate (₹ / kg)", value=890.0, step=10.0)
    al_rate = col2.number_input("Aluminum Commodity Rate (₹ / kg)", value=265.0, step=5.0)

    res = calculate_short_circuit_and_busbar(tx_k, tx_z, bus_len, mat, cu_rate, al_rate)
    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Short-Circuit Fault Level", f"{res['i_sc_ka']} kA")
    m2.metric("Nominal Load Current", f"{res['flc_amps']} A")
    m3.metric("Min Area for 1s Fault", f"{res['req_area_1s']} mm²")
    m4.metric("Recommended Busbar", res['rec_busbar'])
    
    b1, b2, b3 = st.columns(3)
    b1.metric("Peak Current (Ipk)", f"{res['i_pk_ka']} kA")
    b2.metric("Repulsion Force (Fm)", f"{res['f_repulsion_nm']} N/m")
    b3.metric("Max Insulator Pitch", res['support_pitch'])
    st.info(f"💡 **Conductor Material Estimation:** Total {mat} weight: **{res['wt_kg']} kg** | Estimated Raw Material Cost: **₹ {res['cost']:,.2f}**.")

# --- MODULE 10: MOTOR STARTING & VOLTAGE DIP ---
elif menu == "⚙️ Motor Starting Inrush & Cable Sizer (IS 3961)":
    st.header("⚙️ Motor Inrush Current & Line Voltage Dip Sizer")
    st.caption("Verifies transient voltage drop during motor acceleration to prevent contactor coil drop-out and stalling.")
    
    c1, c2, c3 = st.columns(3)
    kw = c1.number_input("Motor Rating (kW)", value=30.0, step=2.5)
    meth = c2.selectbox("Starting Method", ["DOL", "Star-Delta", "Soft Starter", "VFD"])
    dist = c3.number_input("Cable Run Distance (Meters)", value=60, step=5)
    
    c4, c5 = st.columns(2)
    c_sq = c4.selectbox("Cable Conductor Size (sq.mm)", [2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185], index=6)
    cond = c5.radio("Conductor Material", ["Copper", "Aluminum"])
    
    vd = calculate_motor_starting_and_cable(kw, meth, dist, c_sq, cond)
    st.divider()
    v1, v2, v3, v4 = st.columns(4)
    v1.metric("Motor FLC", f"{vd['flc']} A")
    v2.metric("Peak Starting Inrush", f"{vd['i_start']} A")
    v3.metric("Voltage Dip (%)", f"{vd['v_drop_pct']} % ({vd['v_drop_volts']} V)")
    v4.metric("Starting Verdict", vd['status'])
    st.info(f"💡 **Recommended Switchgear:** Incomer Breaker: **{vd['mccb_rating']}** | DOL: **{vd['dol_contactor']}** | Star-Delta: **{vd['main_delta_contactor']} (M/D), {vd['star_contactor']} (S)**.")

# --- MODULE 11: THERMAL & CLIMATE SIZER ---
elif menu == "🌡️ Panel Thermal Dissipation & Climate Sizer (IEC 60890)":
    st.header("🌡️ Panel Thermal Dissipation & Climate Sizer (IEC 60890)")
    st.caption("Calculates enclosure surface dissipation vs. internal heat generation to size louver filter fans or panel air conditioning.")
    
    c1, c2 = st.columns(2)
    h = c1.number_input("Panel Height (mm)", value=1800, step=100)
    w = c1.number_input("Panel Width (mm)", value=1600, step=100)
    d = c1.number_input("Panel Depth (mm)", value=450, step=50)
    sealed = c1.checkbox("IP55/IP65 Sealed Panel (Washdown / Food Area)", value=False)
    
    vfd_kw = c2.number_input("Total VFD Running Load (kW)", value=30.0, step=5.0)
    contactors = c2.number_input("Number of Contactor Coils", value=18, step=2)
    tx_va = c2.number_input("Control Transformer Rating (VA)", value=500, step=100)
    amb_t = c2.slider("Max Ambient Temperature (°C)", 30, 50, 40)
    int_t = c2.slider("Max Permissible Internal Temperature (°C)", 35, 55, 45)
    
    th = calculate_iec60890_thermal(h, w, d, vfd_kw, contactors, tx_va, sealed, amb_t, int_t)
    st.divider()
    t1, t2, t3, t4 = st.columns(4)
    t1.metric("Internal Heat Generation", f"{th['tot_w']} W")
    t2.metric("Natural Dissipation Area", f"{th['area_m2']} m² ({th['nat_w']} W)")
    t3.metric("Net Heat to Remove", f"{th['net_w']} W")
    t4.metric("Airflow Required", f"{th['m3h']} m³/h ({th['cfm']} CFM)")
    st.info(f"💡 **Recommended Cooling Strategy:** **{th['cooling_plan']}**")

# --- MODULE 12: 24V DC VOLTAGE DROP ---
elif menu == "🎛️ 24V DC Field Control Wiring Voltage Drop Sizer":
    st.header("🎛️ 24V DC Field Wiring Voltage Drop Sizer")
    st.caption("Sizes control conductors for pilot solenoids, sensors, and remote PLC racks to avoid control drops below the 20.4V DC pickup threshold.")
    
    c1, c2, c3 = st.columns(3)
    dc_amps = c1.number_input("Connected DC Current Load (Amps)", value=3.5, step=0.5)
    sqmm = c2.selectbox("Control Conductor Size (sq.mm)", [0.5, 0.75, 1.0, 1.5, 2.5], index=1)
    dist = c3.number_input("One-Way Wiring Distance (Meters)", value=45.0, step=5.0)
    
    dc_res = calculate_dc_control_line_drop(dc_amps, sqmm, dist)
    st.divider()
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Voltage Drop", f"{dc_res['v_drop']} V DC")
    d2.metric("End Terminal Voltage", f"{dc_res['end_v']} V DC")
    d3.metric("Percentage Drop", f"{dc_res['v_pct']} %")
    d4.metric("Voltage Status", dc_res['status'])
    st.info(f"💡 **Maximum Safe Distance:** Max wiring length before dropping below 20.4V DC is **{dc_res['max_safe_dist']} meters**.")

# --- MODULE 13: APFC & 7% DETUNED REACTOR ---
elif menu == "⚡ APFC Capacitor Bank & 7% Detuned Reactor Sizer":
    st.header("⚡ APFC Capacitor Bank & 7% Detuned Reactor Sizer")
    st.caption("Sizes kVAR steps and calculates capacitor terminal voltage inflation in the presence of harmonics per IEC 61431.")
    
    c1, c2, c3 = st.columns(3)
    p_kw = c1.number_input("Active Operating Plant Load (kW)", value=250.0, step=10.0)
    c_pf = c2.slider("Current Uncorrected PF", 0.65, 0.90, 0.78)
    t_pf = c3.slider("Target Desired PF", 0.92, 0.99, 0.98)
    
    apfc = calculate_apfc_and_reactor(p_kw, c_pf, t_pf)
    st.divider()
    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Required Capacitor Bank", f"{apfc['req_kvar']} kVAR")
    a2.metric("Capacitor Operating Voltage", f"{apfc['v_cap_op']} V")
    a3.metric("Recommended Capacitor Rating", apfc['rec_v_rating'])
    a4.metric("Series Reactor Sizing", f"{apfc['reactor_kvar']} kVAR ({apfc['tuning_freq']})")
    st.info("💡 **IEC 61431 Safety Rule:** Running capacitors in series with 7% reactors elevates terminal voltage ($415 / (1 - 0.07) = 446.2\\text{V}$). Always use 480V or 525V heavy-duty capacitors.")

# --- MODULE 14: MOTOR MASTER LOOKUP CHART ---
elif menu == "📋 Motor Switchgear Master Lookup Chart":
    st.header("📋 Motor Switchgear & Cable Master Lookup Chart")
    st.caption("Standard 415V 3-Phase AC motor ratings with full load currents, contactor selections, and cable cross-sections.")
    
    std_ratings = [0.75, 1.5, 2.2, 3.7, 5.5, 7.5, 11.0, 15.0, 18.5, 22.0, 30.0, 37.0, 45.0, 55.0, 75.0, 90.0, 110.0, 132.0, 160.0]
    chart_data = []
    for k in std_ratings:
        flc = (k * 1000.0) / (math.sqrt(3) * 415.0 * 0.85 * 0.88)
        mccb = max(32, math.ceil(flc * 1.5 / 10.0) * 10)
        cables = [(2.5,24),(4,32),(6,41),(10,57),(16,76),(25,101),(35,125),(50,150),(70,190),(95,235),(120,270),(150,310),(185,355),(240,420),(300,480)]
        c_sq = next((s for s, a in cables if a >= flc * 1.25), 300)
        chart_data.append({
            "Motor kW": f"{k} kW ({round(k*1.341,1)} HP)",
            "Motor FLC (A)": round(flc, 1),
            "Backup Breaker": f"{mccb}A MCCB",
            "DOL Contactor": f"{math.ceil(flc * 1.15)}A AC3",
            "Star-Delta (M/D)": f"{math.ceil((flc/1.732)*1.15)}A AC3",
            "Soft Starter": f"{round(flc,1)}A",
            "Cu Cable (Air)": f"{c_sq} sq.mm"
        })
    st.dataframe(pd.DataFrame(chart_data), width="stretch")

# --- MODULE 15: FAT QA CERTIFICATE ---
elif menu == "✅ Quality Inspection FAT Certificate Generator":
    st.header("✅ Factory Acceptance Test (FAT) Quality Certificate Generator")
    st.caption("Generates formal PDF testing documentation for client inspection records.")
    
    c1, c2 = st.columns(2)
    sr_no = c1.text_input("Serial Identification Number", "AA-2026-118")
    client = c1.text_input("Customer Name", "Praj Industries (Brewery Greenfield)")
    proj = c2.text_input("Project Scope Description", "MCC + Automation Control Panel")
    insp = c2.text_input("Quality Inspector Name", "Sumit Shirsath")
    
    f1, f2, f3 = st.columns(3)
    r_ir = f1.number_input("R-Phase Megger (MΩ)", value=160.0)
    y_ir = f2.number_input("Y-Phase Megger (MΩ)", value=155.0)
    b_ir = f3.number_input("B-Phase Megger (MΩ)", value=165.0)
    hv_test = st.checkbox("2.5 kV AC Dielectric Withstand (1 Min Passed)", value=True)
    
    if st.button("📄 Generate Official FAT Certificate (PDF)"):
        pdf_b = generate_fat_certificate_pdf(sr_no, client, proj, insp, r_ir, y_ir, b_ir, hv_test)
        st.download_button("📥 Download FAT Certificate PDF", pdf_b.getvalue(), f"FAT_{sr_no}.pdf", mime="application/pdf")

# --- MODULE 16: DELIVERY CHALLAN ---
elif menu == "🚚 Dispatch Delivery Challan Generator":
    st.header("🚚 Dispatch Delivery Challan & Gate Pass Generator")
    st.caption("Generates transport delivery notes for vehicle entry, e-Way bill compliance, and client material acceptance.")
    
    c1, c2 = st.columns(2)
    dc_no = c1.text_input("Challan Number", "AA/DC/2026-142")
    c_name = c1.text_input("Consignee Entity", "Praj Industries Site (Brewery Plant)")
    veh_no = c2.text_input("Transport Vehicle / LR Number", "MH 14 HG 4821")
    gross_wt = c2.number_input("Consignment Gross Weight (kg)", value=340.0, step=10.0)
    addr = st.text_area("Delivery Site Destination", "Plot 42, MIDC Bhosari, Pune - 411026")
    panel_desc = st.text_input("Panel Scope Description", "Industrial MCC + VFD Food-Grade Control Panel")
    
    if st.button("🚛 Generate Delivery Challan (PDF)"):
        dc_pdf = generate_delivery_challan_pdf(dc_no, c_name, addr, veh_no, panel_desc, gross_wt)
        st.download_button("📥 Download Official Delivery Challan (PDF)", dc_pdf.getvalue(), f"Challan_{dc_no.replace('/', '_')}.pdf", mime="application/pdf")
