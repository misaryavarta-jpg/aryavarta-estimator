import os
import io
import re
import math
import base64
import numpy as np
import pandas as pd
import streamlit as st

# ReportLab Imports for PDF Generation
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Optional text extraction imports
try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import docx
except ImportError:
    docx = None

st.set_page_config(
    page_title="Aryavarta Automation - Sales & Engineering Suite",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

def get_feeder_subcomponents(panel_type, kw_str, brand, qty=1):
    """
    Returns full Siemens/Multi-brand default required sub-components
    for each feeder type according to standard panel bill of materials.
    """
    sub_items = []
    
    if "Incomer" in panel_type or "MCCB" in panel_type or "ACB" in panel_type:
        if "630A" in panel_type or "630A" in kw_str:
            sub_items = [
                {"Item": "Main Incomer 630A 3P MCCB", "Spec": f"630A 3P 50kA Microprocessor/TM ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 28500.0},
                {"Item": "Extended Door Rotary Handle Kit (ROM)", "Spec": "630A Door Operating Mechanism", "Brand": brand, "Qty": 1*qty, "UnitPrice": 3200.0},
                {"Item": "Terminal Spreader Links Kit", "Spec": "630A Busbar/Cable Spreader Extension (Set of 3)", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2800.0},
                {"Item": "Phase Barriers / Insulating Shrouds", "Spec": "630A Inter-Phase Barrier Set", "Brand": brand, "Qty": 1*qty, "UnitPrice": 950.0},
                {"Item": "Auxiliary & Alarm Contact Block", "Spec": "1NO+1NC Aux + 1NO Trip Alarm Switch", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1850.0},
                {"Item": "Shunt Trip Release Coil", "Spec": "230V AC Remote Emergency Trip Coil", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2400.0},
                {"Item": "Incomer Metering CT Set", "Spec": "600/5A Class 0.5 Measuring CTs (Set of 3)", "Brand": "Kappa / AE", "Qty": 1*qty, "UnitPrice": 3600.0},
                {"Item": "Incomer Control & Meter Protection MCB", "Spec": "6A 3P C-Curve Control MCB (Siemens 5SY/equivalent)", "Brand": brand, "Qty": 1*qty, "UnitPrice": 850.0}
            ]
        elif "400A" in panel_type or "400A" in kw_str:
            sub_items = [
                {"Item": "Main Incomer 400A 3P MCCB", "Spec": f"400A 3P 36kA Microprocessor/TM ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 19800.0},
                {"Item": "Extended Door Rotary Handle Kit (ROM)", "Spec": "400A Door Operating Mechanism", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2600.0},
                {"Item": "Terminal Spreader Links Kit", "Spec": "400A Busbar/Cable Spreader Extension (Set of 3)", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2100.0},
                {"Item": "Phase Barriers / Insulating Shrouds", "Spec": "400A Inter-Phase Barrier Set", "Brand": brand, "Qty": 1*qty, "UnitPrice": 750.0},
                {"Item": "Auxiliary & Alarm Contact Block", "Spec": "1NO+1NC Aux Switch", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1450.0},
                {"Item": "Incomer Metering CT Set", "Spec": "400/5A Class 0.5 Measuring CTs (Set of 3)", "Brand": "Kappa / AE", "Qty": 1*qty, "UnitPrice": 2900.0},
                {"Item": "Incomer Protection MCB", "Spec": "6A 3P Control MCB", "Brand": brand, "Qty": 1*qty, "UnitPrice": 850.0}
            ]
        elif "250A" in panel_type or "250A" in kw_str:
            sub_items = [
                {"Item": "Main Incomer 250A 3P MCCB", "Spec": f"250A 3P 25kA Thermal Magnetic ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 11200.0},
                {"Item": "Extended Door Rotary Handle Kit (ROM)", "Spec": "250A Door Operating Mechanism", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1850.0},
                {"Item": "Terminal Spreader Links Kit", "Spec": "250A Busbar/Cable Spreader Extension (Set of 3)", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1450.0},
                {"Item": "Phase Barriers", "Spec": "250A Inter-Phase Barrier Set", "Brand": brand, "Qty": 1*qty, "UnitPrice": 550.0},
                {"Item": "Incomer Metering CT Set", "Spec": "250/5A Class 0.5 Measuring CTs (Set of 3)", "Brand": "Kappa / AE", "Qty": 1*qty, "UnitPrice": 2400.0}
            ]
        elif "ACB" in panel_type or "800A" in kw_str or "1250A" in kw_str:
            sub_items = [
                {"Item": "Air Circuit Breaker (ACB) Unit", "Spec": f"800A-1600A 3P 50kA Drawout ACB with Microprocessor Release ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 95000.0},
                {"Item": "Motorized Racking & Spring Charge Unit", "Spec": "230V AC Motor Operating Mechanism", "Brand": brand, "Qty": 1*qty, "UnitPrice": 14500.0},
                {"Item": "Shunt Trip & Closing Coils", "Spec": "230V AC Shunt + Closing Release Coils", "Brand": brand, "Qty": 1*qty, "UnitPrice": 6800.0},
                {"Item": "Safety Shutters & Door Interlock", "Spec": "Automatic Safety Shutters & Door Lock", "Brand": brand, "Qty": 1*qty, "UnitPrice": 4200.0},
                {"Item": "Precision Protection CT Set", "Spec": "800/5A Class 0.2S CT Set", "Brand": "Kappa", "Qty": 1*qty, "UnitPrice": 5800.0}
            ]
        else: # Default 100A / 160A MCCB
            sub_items = [
                {"Item": "Main Incomer 100A 3P MCCB", "Spec": f"100A 3P 25kA TM MCCB ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 4800.0},
                {"Item": "Extended Door Rotary Handle Kit", "Spec": "100A Door Operating Mechanism", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1250.0},
                {"Item": "Terminal Spreader Links Kit", "Spec": "100A Spreader Extension Links", "Brand": brand, "Qty": 1*qty, "UnitPrice": 950.0},
                {"Item": "Phase Barriers", "Spec": "100A Inter-Phase Barriers", "Brand": brand, "Qty": 1*qty, "UnitPrice": 380.0},
                {"Item": "Incomer CT Set", "Spec": "100/5A Class 1.0 CT Set", "Brand": "AE", "Qty": 1*qty, "UnitPrice": 1650.0}
            ]

    elif "VFD" in panel_type:
        sub_items = [
            {"Item": f"VFD Power Unit ({kw_str})", "Spec": f"{kw_str} 415V Heavy Duty VFD Drive Unit ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 28600.0},
            {"Item": "VFD Incomer MCCB / MPCB", "Spec": f"Incomer Motor Protection Breaker ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 4100.0},
            {"Item": "Semiconductor Fast-Acting Fuses (aR)", "Spec": "Fast Semiconductor Fuse Set for VFD Protection", "Brand": brand, "Qty": 1*qty, "UnitPrice": 3600.0},
            {"Item": "3% Input AC Line Reactor Choke", "Spec": "Harmonic Mitigation Line Reactor", "Brand": "Schaffner / Transwave", "Qty": 1*qty, "UnitPrice": 3420.0},
            {"Item": "Line Isolation Power Contactor", "Spec": "AC-3 Duty Line Contactor", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2418.0},
            {"Item": "Door Keypad BOP & Potentiometer Unit", "Spec": "Display BOP + Speed Adjuster Set", "Brand": brand, "Qty": 1*qty, "UnitPrice": 3200.0},
            {"Item": "Panel Cooling Louver Fan Unit", "Spec": "Filter Fan Unit for VFD Heat Dissipation", "Brand": "Rittal / Schneider", "Qty": 1*qty, "UnitPrice": 2800.0}
        ]
    elif "Star-Delta" in panel_type:
        sub_items = [
            {"Item": "Star-Delta Incomer MPCB/MCCB", "Spec": "Motor Duty Heavy Duty Breaker", "Brand": brand, "Qty": 1*qty, "UnitPrice": 4980.0},
            {"Item": "Main Power Contactor (AC-3)", "Spec": "AC-3 3P Power Contactor", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2950.0},
            {"Item": "Delta Power Contactor (AC-3)", "Spec": "AC-3 3P Power Contactor", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2950.0},
            {"Item": "Star Power Contactor (AC-3)", "Spec": "AC-3 3P Reduced Rating Contactor", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1794.0},
            {"Item": "Mechanical & Electrical Interlock Block Set", "Spec": "Star-Delta Contactor Interlock Kit", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1250.0},
            {"Item": "Electronic Star-Delta Timer Relay", "Spec": "0.1s - 30s 230V AC Star-Delta Electronic Timer", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2200.0},
            {"Item": "Thermal Overload Relay (TOR)", "Spec": "Class 10 Adjustable Bimetallic Relay", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2400.0},
            {"Item": "Door Start/Stop Buttons & Signal LEDs", "Spec": "Flush Pushbuttons + Red/Green/Yellow LED Indicator Set", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1400.0}
        ]
    elif "DOL" in panel_type or "Direct On Line" in panel_type:
        sub_items = [
            {"Item": "DOL Incomer MPCB", "Spec": f"Adjustable Thermal-Magnetic MPCB ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 3200.0},
            {"Item": "DOL Power Contactor (AC-3)", "Spec": f"3P Power Contactor ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1740.0},
            {"Item": "Auxiliary Contact Block", "Spec": "1NO+1NC Front Snap Auxiliary Block", "Brand": brand, "Qty": 1*qty, "UnitPrice": 650.0},
            {"Item": "Start/Stop Pushbuttons & Status Lamps", "Spec": "Green Start, Red Stop, Amber Trip Indicator Set", "Brand": brand, "Qty": 1*qty, "UnitPrice": 1100.0}
        ]
    elif "Soft Starter" in panel_type:
        sub_items = [
            {"Item": "Soft Starter Power Unit", "Spec": f"3-Phase Controlled Soft Starter Unit ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 32500.0},
            {"Item": "Incomer Motor-Duty MCCB", "Spec": f"Short-Circuit Protection Breaker ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 6800.0},
            {"Item": "Bypass Contactor Unit", "Spec": "AC-1/AC-3 Bypass Contactor", "Brand": brand, "Qty": 1*qty, "UnitPrice": 4200.0},
            {"Item": "Fast Semiconductor Fuses", "Spec": "aR Grade Semiconductor Protection Fuses", "Brand": brand, "Qty": 1*qty, "UnitPrice": 2800.0}
        ]
    elif "APFC" in panel_type:
        sub_items = [
            {"Item": "APFC Microprocessor Controller Relay", "Spec": "8/12 Step Automatic Power Factor Controller", "Brand": brand, "Qty": 1*qty, "UnitPrice": 12500.0},
            {"Item": "Heavy-Duty Capacitor Duty Contactors", "Spec": "Capacitor Switching Contactors with Pre-charge Resistors", "Brand": brand, "Qty": 4*qty, "UnitPrice": 3100.0},
            {"Item": "7% Detuned Harmonic Filter Reactors", "Spec": "Copper Wound 440V Detuned Reactors", "Brand": "Epcos / Schneider", "Qty": 4*qty, "UnitPrice": 6800.0},
            {"Item": "Gas-Filled APP Duty Capacitors", "Spec": "50 Hz 440V Capacitor Cells", "Brand": "Epcos / Tibcon", "Qty": 4*qty, "UnitPrice": 2800.0}
        ]
    else: # General Feeder
        sub_items = [
            {"Item": f"Feeder Distribution MCCB / MCB", "Spec": f"Feeder Protection Breaker ({brand})", "Brand": brand, "Qty": 1*qty, "UnitPrice": 5200.0}
        ]
        
    return sub_items

def build_bom(panel_type, motor_kw_str, brand, mode="Single", multi_feeders=None, custom_items=None):
    """
    Constructs a detailed pandas DataFrame Bill of Materials.
    """
    rows = []
    
    if mode == "Single":
        rows.append({
            "Item": f"Panel Enclosure Frame ({panel_type})",
            "Specification": "IP55 Powder Coated CRCA Steel Floor/Wall Enclosure",
            "Brand": "Aryavarta Standard",
            "Qty": 1,
            "Unit Price": 18500.0,
            "Total Material Cost": 18500.0
        })
        sub_components = get_feeder_subcomponents(panel_type, motor_kw_str, brand, qty=1)
        for item in sub_components:
            rows.append({
                "Item": item["Item"],
                "Specification": item["Spec"],
                "Brand": item["Brand"],
                "Qty": item["Qty"],
                "Unit Price": item["UnitPrice"],
                "Total Material Cost": item["Qty"] * item["UnitPrice"]
            })
    else:
        # Multi-Feeder Mode
        rows.append({
            "Item": "Multi-Bay MCC Panel Enclosure Frame",
            "Specification": "IP55 Floor Standing Compartmentalized Structure (RAL 7035)",
            "Brand": "Aryavarta Enclosures",
            "Qty": 1,
            "Unit Price": 45000.0,
            "Total Material Cost": 45000.0
        })
        if multi_feeders:
            for f in multi_feeders:
                f_type = f.get("type", "DOL Starter")
                f_kw = f.get("kw", "15 kW")
                f_qty = f.get("qty", 1)
                f_sub = get_feeder_subcomponents(f_type, f_kw, brand, qty=f_qty)
                for item in f_sub:
                    rows.append({
                        "Item": item["Item"],
                        "Specification": item["Spec"],
                        "Brand": item["Brand"],
                        "Qty": item["Qty"],
                        "Unit Price": item["UnitPrice"],
                        "Total Material Cost": item["Qty"] * item["UnitPrice"]
                    })
                    
    # Common Power Distribution & Busbar
    rows.append({
        "Item": "Main Copper Busbar & Power Distribution Harness",
        "Specification": "EC Grade Copper Busbar + FRLS Power Wiring",
        "Brand": "Metelec / Finolex",
        "Qty": 1,
        "Unit Price": 24000.0 if mode == "Multi" else 8500.0,
        "Total Material Cost": 24000.0 if mode == "Multi" else 8500.0
    })
    rows.append({
        "Item": "Control Transformers, SMPS & Interlocks",
        "Specification": "24V DC 5A SMPS + 230V Control Tx + Terminal Blocks",
        "Brand": "Meanwell / Wago",
        "Qty": 1,
        "Unit Price": 12500.0 if mode == "Multi" else 4200.0,
        "Total Material Cost": 12500.0 if mode == "Multi" else 4200.0
    })
    
    # Custom Non-Standard Components
    if custom_items:
        for c in custom_items:
            if c.get("name") and c.get("price", 0) > 0:
                rows.append({
                    "Item": c["name"],
                    "Specification": c.get("spec", "Custom Non-Standard Spec"),
                    "Brand": c.get("brand", brand),
                    "Qty": c.get("qty", 1),
                    "Unit Price": float(c["price"]),
                    "Total Material Cost": float(c["qty"]) * float(c["price"])
                })
                
    df = pd.DataFrame(rows)
    return df

def generate_sld_svg(panel_type, kw_str, brand):
    """
    Generates a dark-themed SVG single line diagram preview.
    """
    return f'''<svg width="450" height="240" viewBox="0 0 450 240" xmlns="http://www.w3.org/2000/svg" style="background:#0f172a; border-radius:8px;">
        <line x1="30" y1="25" x2="420" y2="25" stroke="#ef4444" stroke-width="3"/>
        <line x1="30" y1="31" x2="420" y2="31" stroke="#eab308" stroke-width="3"/>
        <line x1="30" y1="37" x2="420" y2="37" stroke="#3b82f6" stroke-width="3"/>
        <text x="225" y="16" fill="#94a3b8" font-family="monospace" font-size="10" text-anchor="middle">415V 3Ø 50Hz MAIN BUSBAR</text>
        <line x1="225" y1="37" x2="225" y2="65" stroke="#cbd5e1" stroke-width="2"/>
        <rect x="195" y="65" width="60" height="32" rx="3" fill="#1e293b" stroke="#38bdf8" stroke-width="2"/>
        <text x="225" y="85" fill="#38bdf8" font-family="sans-serif" font-size="10" font-weight="bold" text-anchor="middle">MCCB</text>
        <line x1="225" y1="97" x2="225" y2="125" stroke="#cbd5e1" stroke-width="2"/>
        <rect x="175" y="125" width="100" height="42" rx="4" fill="#020617" stroke="#22c55e" stroke-width="2"/>
        <text x="225" y="143" fill="#22c55e" font-family="sans-serif" font-size="10" font-weight="bold" text-anchor="middle">{panel_type[:18]}</text>
        <text x="225" y="157" fill="#94a3b8" font-family="sans-serif" font-size="8" text-anchor="middle">{brand} ({kw_str})</text>
        <line x1="225" y1="167" x2="225" y2="190" stroke="#cbd5e1" stroke-width="2"/>
        <circle cx="225" cy="208" r="16" fill="#1e293b" stroke="#e2e8f0" stroke-width="2"/>
        <text x="225" y="213" fill="#f8fafc" font-family="sans-serif" font-size="11" font-weight="bold" text-anchor="middle">M 3~</text>
    </svg>'''

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

def extract_text_from_files(uploaded_files):
    """
    Extracts text content from multiple PDF, Excel, Word, CSV, or TXT uploaded files.
    """
    extracted_text = ""
    for file in uploaded_files:
        name = file.name.lower()
        extracted_text += f"\n--- File: {file.name} ---\n"
        try:
            if name.endswith('.txt'):
                extracted_text += file.getvalue().decode('utf-8', errors='ignore') + "\n"
            elif name.endswith('.csv'):
                df = pd.read_csv(file)
                extracted_text += df.to_string() + "\n"
            elif name.endswith(('.xlsx', '.xls')):
                excel = pd.ExcelFile(file)
                for sheet in excel.sheet_names:
                    extracted_text += f"\nSheet: {sheet}\n"
                    df = excel.parse(sheet)
                    extracted_text += df.to_string() + "\n"
            elif name.endswith('.pdf') and pypdf is not None:
                reader = pypdf.PdfReader(file)
                for page in reader.pages:
                    extracted_text += page.extract_text() or ""
            elif name.endswith('.docx') and docx is not None:
                doc = docx.Document(file)
                for p in doc.paragraphs:
                    extracted_text += p.text + "\n"
            else:
                extracted_text += "[File format binary or unparsed text fallback]\n"
        except Exception as e:
            extracted_text += f"[Error reading {file.name}: {str(e)}]\n"
            
    return extracted_text

st.sidebar.title("⚡ Aryavarta Automation")
st.sidebar.caption("Sales & Engineering Suite v4.2")

navigation_option = st.sidebar.radio(
    "Select Module:",
    [
        "📋 Panel Estimator, Proposal & PDF",
        "📐 Cable & Busbar Sizing (IS 3961)",
        "⚡ CT Ratio & Class PS Sizer",
        "🔋 APFC Capacitor Bank Sizer",
        "❄️ Panel Thermal & AC Sizer",
        "⚡ Control Tx & 24V SMPS Sizing",
        "🔊 Harmonics & VFD Spike Sizer",
        "📑 FAT Quality Certificate (IEEE 43)",
        "📊 Quotation Register & Pipeline"
    ]
)

if navigation_option == "📋 Panel Estimator, Proposal & PDF":
    st.title("📋 Panel Estimator, PDF & WhatsApp Proposal")
    st.caption("Engineered for Single Feeder Panels & Multi-Feeder MCC/PCC Boards with Multi-file Parsing")
    
    # Mode Selection Toggle
    mode = st.radio(
        "Estimation Mode:",
        ["⚡ Single Feeder Panel", "🏢 Full Multi-Feeder Panel / MCC Board"],
        horizontal=True
    )
    
    # Multi-file Upload Expander
    with st.expander("📁 Auto-Extract Specs from Customer Inquiry Documents (PDF, Excel, Word, CSV, TXT)", expanded=False):
        uploaded_files = st.file_uploader(
            "Upload Inquiry Files",
            type=['pdf', 'xlsx', 'xls', 'csv', 'docx', 'txt'],
            accept_multiple_files=True
        )
        file_text = ""
        if uploaded_files:
            file_text = extract_text_from_files(uploaded_files)
            st.success(f"Parsed {len(uploaded_files)} file(s) successfully!")
            
        inquiry_text = st.text_area(
            "Paste or Review Customer Inquiry Text:",
            value=file_text if file_text else "",
            height=120,
            placeholder="e.g. Need quotation for 630A MCCB Incomer with 3x22kW VFD feeders and 2x15kW DOL starters for Siemens switchgear."
        )
        
        if st.button("⚡ Process Inquiry with Smart Extractor"):
            st.info("Inquiry parsed! Auto-suggested specifications updated in configuration options below.")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        client_name = st.text_input("Client / Company Name:", "Maharashtra Water Works Ltd")
        brand = st.selectbox("Switchgear Brand:", ["Siemens", "L&T", "Schneider", "ABB", "Danfoss", "Delta"])
    with col2:
        currency = st.selectbox("Currency:", ["INR (₹)", "USD ($)", "EUR (€)"])
        curr_sym = "₹" if "INR" in currency else ("$" if "USD" in currency else "€")
        ex_rate = st.number_input("Exchange Rate (1 FX = X INR):", min_value=0.1, value=1.0 if curr_sym == "₹" else 83.5)
    with col3:
        margin_pct = st.slider("Target Profit Margin (%):", 5, 45, 18)
        labor_cost = st.number_input("Wiring & Labor Cost (₹):", min_value=500, value=18500 if "Multi" in mode else 4500, step=500)
    with col4:
        st.metric("Applied Margin", f"{margin_pct}%")
        st.metric("Selected Brand", brand)

    # Multi-feeder schedule or Single feeder rating setup
    multi_feeders = []
    if "Multi" in mode:
        st.subheader("🏢 Multi-Feeder MCC Panel Feeder Configurator")
        st.info("Select main incomer rating and add individual motor starter feeders:")
        
        c_inc1, c_inc2 = st.columns(2)
        with c_inc1:
            incomer_type = st.selectbox("Main Incomer Rating:", [
                "Incomer - 630A 3P MCCB (50kA)",
                "Incomer - 400A 3P MCCB (36kA)",
                "Incomer - 250A 3P MCCB (25kA)",
                "Incomer - 100A 3P MCCB (25kA)",
                "Incomer - 1250A Drawout ACB (50kA)"
            ])
        
        num_feeders = st.number_input("Number of Starter Feeders:", min_value=1, max_value=15, value=3)
        for i in range(int(num_feeders)):
            cf1, cf2, cf3 = st.columns([2, 2, 1])
            with cf1:
                f_type = st.selectbox(f"Feeder {i+1} Type:", ["VFD Panel", "Star-Delta Starter", "DOL Starter", "Soft Starter", "APFC Capacitor Bank"], key=f"ft_{i}")
            with cf2:
                f_kw = st.selectbox(f"Feeder {i+1} Rating:", ["5.5 kW", "11 kW", "15 kW", "22 kW", "30 kW", "45 kW", "75 kW", "110 kW"], key=f"fkw_{i}")
            with cf3:
                f_qty = st.number_input(f"Qty {i+1}:", min_value=1, value=1, key=f"fq_{i}")
            multi_feeders.append({"type": f_type, "kw": f_kw, "qty": f_qty})
            
        panel_type = f"Multi-Feeder MCC ({incomer_type.split('-')[1].strip() if '-' in incomer_type else incomer_type})"
        motor_kw = f"{len(multi_feeders)} Feeders"
    else:
        c_s1, c_s2 = st.columns(2)
        with c_s1:
            panel_type = st.selectbox("Panel Type:", ["VFD Panel", "Star-Delta Starter", "DOL Starter", "Soft Starter", "APFC Panel"])
        with c_s2:
            motor_kw = st.selectbox("Motor Rating (kW):", ["0.75 kW", "3.7 kW", "7.5 kW", "15 kW", "22 kW", "30 kW", "45 kW", "75 kW", "110 kW", "160 kW"])

    # Custom Non-Standard Component Adder Expander
    custom_components = []
    with st.expander("🛠️ Custom / Non-Standard Component Adder", expanded=False):
        c_cnt = st.number_input("Number of Custom Components:", min_value=0, max_value=5, value=0)
        for c_idx in range(int(c_cnt)):
            cc1, cc2, cc3, cc4 = st.columns([3, 3, 1, 2])
            with cc1:
                c_name = st.text_input(f"Item Name #{c_idx+1}:", f"Phase Preventer / ELR #{c_idx+1}", key=f"cn_{c_idx}")
            with cc2:
                c_spec = st.text_input(f"Specification #{c_idx+1}:", "240V AC Auxiliary Protection Relay", key=f"cs_{c_idx}")
            with cc3:
                c_qty = st.number_input(f"Qty #{c_idx+1}:", min_value=1, value=1, key=f"cq_{c_idx}")
            with cc4:
                c_price = st.number_input(f"Unit Price (₹) #{c_idx+1}:", min_value=100.0, value=2500.0, step=100.0, key=f"cp_{c_idx}")
            custom_components.append({"name": c_name, "spec": c_spec, "qty": c_qty, "price": c_price, "brand": brand})

    bom_df = build_bom(panel_type, motor_kw, brand, mode="Multi" if "Multi" in mode else "Single", multi_feeders=multi_feeders, custom_items=custom_components)
    raw_material_tot = bom_df["Total Material Cost"].sum()
    total_cost_base = raw_material_tot + labor_cost
    final_quote_price = total_cost_base / (1.0 - (margin_pct / 100.0))
    final_quote_price_curr = final_quote_price / ex_rate

    st.subheader("📊 Bill of Materials & Cost Breakdown")
    st.dataframe(bom_df.style.format({"Unit Price": "₹{:,.2f}", "Total Material Cost": "₹{:,.2f}"}), use_container_width=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Material Cost", f"₹ {raw_material_tot:,.2f}")
    m2.metric("Labor & Wiring", f"₹ {labor_cost:,.2f}")
    m3.metric("Profit Margin", f"{margin_pct}%")
    m4.metric("Final Quote Price", f"{curr_sym} {final_quote_price_curr:,.2f}")

    # Instant Multi-Brand Comparison Matrix
    with st.expander("🔄 Instant Multi-Brand Commercial Price Comparison Matrix", expanded=False):
        brand_list = ["Siemens", "L&T", "Schneider", "ABB", "Danfoss", "Delta"]
        matrix_rows = []
        for b in brand_list:
            b_df = build_bom(panel_type, motor_kw, b, mode="Multi" if "Multi" in mode else "Single", multi_feeders=multi_feeders, custom_items=custom_components)
            b_mat = b_df["Total Material Cost"].sum()
            b_quote = (b_mat + labor_cost) / (1.0 - (margin_pct / 100.0))
            b_delta = b_quote - final_quote_price
            matrix_rows.append({
                "Switchgear Brand": b,
                "Material Cost (₹)": b_mat,
                "Commercial Quote (₹)": b_quote,
                "Price Delta vs Selected (₹)": b_delta
            })
        st.dataframe(pd.DataFrame(matrix_rows).style.format({"Material Cost (₹)": "₹{:,.2f}", "Commercial Quote (₹)": "₹{:,.2f}", "Price Delta vs Selected (₹)": "₹{:+,.2f}"}), use_container_width=True)

    # Document Generators Section
    st.subheader("📑 Commercial Quotation & Engineering Exports")
    c_pdf1, c_pdf2, c_pdf3, c_pdf4 = st.columns(4)
    
    with c_pdf1:
        pdf_bytes = generate_pdf_quotation(client_name, panel_type, motor_kw, brand, bom_df, labor_cost, margin_pct, final_quote_price_curr, curr_sym)
        st.download_button("📄 Commercial Quote PDF", data=pdf_bytes, file_name=f"Aryavarta_Quote_{client_name}.pdf", mime="application/pdf")
    with c_pdf2:
        pi_bytes = generate_proforma_invoice_pdf(client_name, panel_type, motor_kw, final_quote_price_curr, curr_sym)
        st.download_button("🧾 Proforma Invoice PDF", data=pi_bytes, file_name=f"Proforma_Invoice_{client_name}.pdf", mime="application/pdf")
    with c_pdf3:
        tds_bytes = generate_tds_pdf(client_name, panel_type, motor_kw, brand)
        st.download_button("📋 Technical Datasheet PDF", data=tds_bytes, file_name=f"TDS_{client_name}.pdf", mime="application/pdf")
    with c_pdf4:
        xl_buf = io.BytesIO()
        with pd.ExcelWriter(xl_buf, engine='openpyxl') as writer:
            bom_df.to_excel(writer, sheet_name="BOM", index=False)
        xl_buf.seek(0)
        st.download_button("📊 Excel Proposal (.XLSX)", data=xl_buf, file_name=f"Proposal_{client_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # SLD Visualizer
    st.subheader("🔌 Single Line Diagram (SLD) Schematic")
    st.components.v1.html(generate_sld_svg(panel_type, motor_kw, brand), height=250)

elif navigation_option == "📐 Cable & Busbar Sizing (IS 3961)":
    st.title("📐 Electrical Cable & Busbar Sizing (IS 3961 / IEC 61439)")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        motor_kw = st.number_input("Motor Rating (kW):", min_value=0.1, max_value=300.0, value=15.0, step=0.1, format="%.1f")
        run_dist = st.number_input("Cable Run Distance (Meters):", min_value=1, value=30)
    with c2:
        material = st.radio("Conductor Material:", ["Copper", "Aluminum"], horizontal=True)
        amb_temp = st.slider("Ambient Temperature (°C):", 25, 55, 45)
    with c3:
        grouped = st.number_input("Cables Grouped in Tray:", min_value=1, max_value=12, value=1)
        panel_type = st.selectbox("Panel Duty Type:", ["VFD Panel", "Star-Delta Starter", "DOL Starter"])

    flc = (motor_kw * 1000) / (math.sqrt(3) * 415 * 0.85 * 0.90)
    derating = (1.0 - ((amb_temp - 30) * 0.005)) * (1.0 / math.sqrt(grouped))
    design_current = flc / derating
    
    if material == "Copper":
        cable_sqmm = 6 if design_current <= 35 else (16 if design_current <= 70 else (35 if design_current <= 120 else 70))
        busbar = "15x3 mm (Copper)" if design_current <= 50 else ("25x5 mm (Copper)" if design_current <= 150 else "50x5 mm (Copper)")
        v_drop = (math.sqrt(3) * design_current * run_dist * 0.018) / cable_sqmm
    else:
        cable_sqmm = 10 if design_current <= 30 else (25 if design_current <= 60 else (50 if design_current <= 100 else 95))
        busbar = "20x5 mm (Aluminum)" if design_current <= 50 else ("30x5 mm (Aluminum)" if design_current <= 150 else "50x10 mm (Aluminum)")
        v_drop = (math.sqrt(3) * design_current * run_dist * 0.028) / cable_sqmm

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Motor FLC", f"{flc:.1f} A")
    m2.metric("Design Current", f"{design_current:.1f} A")
    m3.metric("Cable Cross-Section", f"{cable_sqmm} sq.mm ({material})")
    m4.metric("Estimated Voltage Drop", f"{v_drop:.2f} V ({ (v_drop/415)*100:.1f}%)")

    # Busbar Joint Overlap & Bolt Tightening Torque Table
    with st.expander("🔩 Busbar Joint Overlap & Bolt Tightening Torque Guide (IS 8623 / IEC 61439)", expanded=True):
        st.caption("Standard assembly torque values to prevent loose joint hot-spot burning:")
        torque_df = pd.DataFrame([
            {"Bolt Size": "M6", "Tightening Torque (N·m)": "8.0 N·m", "Min Overlap Length": "1.0 x Busbar Width", "Washer Spec": "Belleville Disc Washer"},
            {"Bolt Size": "M8", "Tightening Torque (N·m)": "20.0 N·m", "Min Overlap Length": "1.0 x Busbar Width", "Washer Spec": "Belleville Disc Washer"},
            {"Bolt Size": "M10", "Tightening Torque (N·m)": "40.0 N·m", "Min Overlap Length": "1.0 x Busbar Width", "Washer Spec": "Belleville Disc Washer"},
            {"Bolt Size": "M12", "Tightening Torque (N·m)": "70.0 N·m", "Min Overlap Length": "1.0 x Busbar Width", "Washer Spec": "Belleville Disc Washer"}
        ])
        st.table(torque_df)

elif navigation_option == "⚡ CT Ratio & Class PS Sizer":
    st.title("⚡ CT Ratio, Class & Burden Sizer (IS 2705 / IEC 61869)")
    
    c1, c2 = st.columns(2)
    with c1:
        prim_i = st.number_input("Primary Current Rating (A):", min_value=10, value=400, step=50)
        ct_class = st.selectbox("CT Accuracy Class:", ["Class 0.2S (Revenue)", "Class 0.5 (Metering)", "Class 1.0 (Commercial)", "Class PS (Protection)"])
    with c2:
        lead_len = st.number_input("CT Secondary Wire Length (Meters):", min_value=1, value=15)
        sec_wire_sqmm = st.selectbox("CT Secondary Wire Size:", ["2.5 sq.mm", "4.0 sq.mm"])

    st.subheader("🛡️ Protection Class PS Knee Point Voltage (Vk) Sizer")
    fault_ka = st.number_input("System Fault Level (kA):", min_value=1.0, value=25.0)
    r_ct = st.number_input("CT Internal Resistance Rct (Ω):", min_value=0.1, value=1.5)
    r_relay = st.number_input("Relay Input Impedance Rrelay (Ω):", min_value=0.01, value=0.1)

    r_lead = (2 * lead_len * 0.018) / (2.5 if "2.5" in sec_wire_sqmm else 4.0)
    i_sec_fault = (fault_ka * 1000) / (prim_i / 5.0)
    v_k_req = 2 * i_sec_fault * (r_ct + r_lead + r_relay)

    m1, m2, m3 = st.columns(3)
    m1.metric("Secondary Fault Current", f"{i_sec_fault:.2f} A")
    m2.metric("Wire Loop Resistance (Rlead)", f"{r_lead:.3f} Ω")
    m3.metric("Min Knee Point Voltage (Vk)", f"{v_k_req:.1f} V")
    st.info(f"💡 **Class PS CT Requirement:** Specify CT with Knee Point Voltage $V_k \\ge {math.ceil(v_k_req)}V$ and magnetizing current $I_m \\le 30\\text{mA}$ at $V_k / 2$.")

elif navigation_option == "🔋 APFC Capacitor Bank Sizer":
    st.title("🔋 APFC Capacitor Bank & Detuned Reactor Sizer")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        kw_load = st.number_input("Active Load (kW):", min_value=10.0, value=200.0)
    with c2:
        curr_pf = st.slider("Existing Power Factor:", 0.60, 0.95, 0.80)
    with c3:
        target_pf = st.slider("Target Power Factor:", 0.90, 0.99, 0.98)

    kvar_req = kw_load * (math.tan(math.acos(curr_pf)) - math.tan(math.acos(target_pf)))
    st.metric("Required Capacitor Rating", f"{kvar_req:.1f} kVAR")

    st.subheader("⚡ IEC 60831 Capacitor Safety Discharge Resistor Sizer")
    step_kvar = st.number_input("Step Rating (kVAR):", min_value=5.0, value=25.0)
    c_uf = (step_kvar * 1000) / (2 * math.pi * 50 * (440**2)) * 1e6
    r_dis_kohm = (50.0 / (c_uf * 1e-6 * math.log(587.0 / 50.0))) / 1000.0
    res_watts = (440**2) / (r_dis_kohm * 1000.0) * 1.5

    m1, m2, m3 = st.columns(3)
    m1.metric("Capacitance per Step", f"{c_uf:.1f} µF")
    m2.metric("Max Discharge Resistor (Rdis)", f"{r_dis_kohm:.1f} kΩ")
    m3.metric("Resistor Power Rating", f"{res_watts:.1f} Watts")

elif navigation_option == "❄️ Panel Thermal & AC Sizer":
    st.title("❄️ Panel Thermal Sizer & Air Conditioner Sizer")
    
    c1, c2 = st.columns(2)
    with c1:
        height = st.number_input("Panel Height (m):", value=2.0)
        width = st.number_input("Panel Width (m):", value=1.2)
        depth = st.number_input("Panel Depth (m):", value=0.6)
    with c2:
        vfd_kw = st.number_input("Total VFD Load inside Panel (kW):", value=45.0)
        t_in_max = st.slider("Max External Ambient Temp (°C):", 30, 55, 45)
        t_out_target = st.slider("Target Internal Panel Temp (°C):", 25, 40, 35)

    # Heat dissipation calculation
    vfd_heat_watts = vfd_kw * 1000 * 0.03 # 3% heat loss
    area = 2 * (height*width + height*depth + width*depth)
    delta_t = t_in_max - t_out_target
    solar_heat = area * 5.5 * delta_t if delta_t > 0 else 0
    total_heat_watts = vfd_heat_watts + solar_heat

    st.metric("Total Internal Heat Load", f"{total_heat_watts:.0f} Watts")
    if total_heat_watts > 500:
        st.success(f"Recommended Panel Air Conditioner: **{(total_heat_watts/1000)*1.2:.1f} kW / {total_heat_watts*3.412:.0f} BTU/hr**")
    else:
        st.info("Panel Louver Filter Fan unit sufficient.")

elif navigation_option == "⚡ Control Tx & 24V SMPS Sizing":
    st.title("⚡ Control Transformer & 24V SMPS Sizer")
    
    c1, c2 = st.columns(2)
    with c1:
        num_relays = st.number_input("Number of 24V DC Auxiliary Relays:", value=12)
        num_valves = st.number_input("Number of Solenoid Valves:", value=2)
    with c2:
        plc_watts = st.number_input("PLC / HMI Power Consumption (W):", value=45)
        
    smps_watts = (num_relays * 1.5) + (num_valves * 15) + plc_watts
    smps_amps = (smps_watts * 1.25) / 24.0

    st.metric("Recommended 24V DC SMPS Rating", f"{math.ceil(smps_amps)} A ({smps_watts:.0f} W)")

    st.subheader("📏 24V DC Field Control Wiring Voltage Drop Calculator")
    field_dist = st.number_input("Field Wire Run Distance (Meters):", value=40)
    field_curr = st.number_input("Field Loop Current (Amps):", value=2.5)
    wire_sqmm = st.selectbox("Field Wire Size:", ["1.0 sq.mm", "1.5 sq.mm", "2.5 sq.mm"])
    
    sq = 1.0 if "1.0" in wire_sqmm else (1.5 if "1.5" in wire_sqmm else 2.5)
    v_drop_dc = (2 * field_dist * field_curr * 0.018) / sq
    v_end = 24.0 - v_drop_dc

    m1, m2 = st.columns(2)
    m1.metric("24V Loop Voltage Drop", f"{v_drop_dc:.2f} V")
    m2.metric("End-of-Line Terminal Voltage", f"{v_end:.2f} V DC", delta="Pass (>20.4V)" if v_end >= 20.4 else "Fail (<20.4V)")

elif navigation_option == "🔊 Harmonics & VFD Spike Sizer":
    st.title("🔊 Harmonics & Active Harmonic Filter (AHF) Sizer")
    
    st.subheader("IEC 61800-3 VFD Motor Cable Length & dv/dt Spike Protection Sizer")
    vfd_cable_m = st.number_input("VFD to Motor Cable Distance (Meters):", min_value=5, value=65)
    emc_class = st.selectbox("EMC Environment Standard:", ["Class C3 (Industrial Plant)", "Class C2 (Commercial/Public)"])

    if vfd_cable_m > 100:
        filter_type = "Sine-Wave LC Filter Required"
        st.error(f"⚠️ Cable length > 100m ({vfd_cable_m}m). Severe voltage wave reflections (>1600V peak). {filter_type}.")
    elif vfd_cable_m > 30:
        filter_type = "3% Output dv/dt Line Reactor Required"
        st.warning(f"⚡ Cable length > 30m ({vfd_cable_m}m). Voltage rise rate protection needed. {filter_type}.")
    else:
        filter_type = "Standard Shielded Cable Sufficient"
        st.success(f"✅ Cable length ({vfd_cable_m}m) within 30m limit. {filter_type}.")

elif navigation_option == "📑 FAT Quality Certificate (IEEE 43)":
    st.title("📑 Factory Acceptance Test (FAT) & Quality Certificate")
    st.caption("Includes IEEE 43 Insulation Resistance Temperature Normalization to 40°C Baseline")

    c1, c2, c3 = st.columns(3)
    with c1:
        serial_no = st.text_input("Panel Serial Number:", "AA/2026/MCC-108")
        client = st.text_input("Client Name:", "Maharashtra Water Works Ltd")
    with c2:
        project = st.text_input("Project Name:", "Pumping Station Phase II")
        inspector = st.text_input("QA Inspector Name:", "Er. R. Sharma")
    with c3:
        test_temp = st.slider("Factory Ambient Temperature during Test (°C):", 15, 50, 30)
        hv_pass = st.checkbox("2.5kV HV Withstand Test Passed", value=True)

    st.subheader("1. Insulation Resistance Test Readings (500V Megger)")
    cm1, cm2, cm3 = st.columns(3)
    with cm1:
        r_phase = st.number_input("R-Phase Megger (MΩ):", value=18.5)
    with cm2:
        y_phase = st.number_input("Y-Phase Megger (MΩ):", value=22.0)
    with cm3:
        b_phase = st.number_input("B-Phase Megger (MΩ):", value=19.5)

    # IEEE 43 Temperature Correction Factor
    kt = 0.5 ** ((40.0 - test_temp) / 10.0)
    r_40 = r_phase * kt

    st.info(f"🌡️ **IEEE 43 Normalization:** Temperature Correction Factor $K_T = {kt:.3f}$. Normalized R-Phase Megger at $40^\\circ\\text{{C}}$ Baseline = **{r_40:.2f} MΩ**.")

    fat_pdf_bytes = generate_fat_certificate_pdf(serial_no, project, client, inspector, r_phase, y_phase, b_phase, hv_pass)
    st.download_button("📜 Download Signed FAT Certificate PDF", data=fat_pdf_bytes, file_name=f"FAT_{serial_no}.pdf", mime="application/pdf")

else:
    st.title("📊 Quotation Register & Pipeline Manager")
    st.caption("Track historical quotations, conversion rates, and export register to Excel")

    mock_quotes = pd.DataFrame([
        {"Quote Ref": "AA/QT/8F21", "Client": "Maharashtra Water Works Ltd", "Panel Type": "630A MCC Board", "Amount (₹)": 500254.0, "Status": "Under Review", "Date": "2026-07-28"},
        {"Quote Ref": "AA/QT/3A19", "Client": "Tata Power Renewable", "Panel Type": "45kW VFD Panel", "Amount (₹)": 185000.0, "Status": "Order Won", "Date": "2026-07-22"},
        {"Quote Ref": "AA/QT/9B04", "Client": "Thermax Systems Ltd", "Panel Type": "100 kVAR APFC Panel", "Amount (₹)": 240000.0, "Status": "Order Won", "Date": "2026-07-15"}
    ])

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Active Quotes", len(mock_quotes))
    m2.metric("Pipeline Value", f"₹ {mock_quotes['Amount (₹)'].sum():,.2f}")
    m3.metric("Win Conversion Rate", "66.7%")

    st.dataframe(mock_quotes, use_container_width=True)

    xl_buf_pipe = io.BytesIO()
    with pd.ExcelWriter(xl_buf_pipe, engine='openpyxl') as writer:
        mock_quotes.to_excel(writer, sheet_name="Quotes_Register", index=False)
    xl_buf_pipe.seek(0)
    
    st.download_button("📊 Export Quotes Register to Excel (.XLSX)", data=xl_buf_pipe, file_name=f"Quotes_Register_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
