from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io
import copy

app = Flask(__name__)
CORS(app)

GREEN = "FF00B050"
BLACK = "FF000000"

def green_fill():
    return PatternFill("solid", fgColor=GREEN)

def bold_font(size=10):
    return Font(bold=True, size=size)

def normal_font(size=10):
    return Font(bold=False, size=size)

def style_green_row(ws, row_num, max_col=6):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = green_fill()
        cell.font = bold_font()

def style_bold_row(ws, row_num, max_col=6):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.font = bold_font()

def fmt_currency(val):
    return val if val else 0

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/generate-buildsheet', methods=['POST', 'OPTIONS'])
def generate_buildsheet():
    if request.method == 'OPTIONS':
        return '', 204
    
    d = request.json
    qty = int(d.get('quantity', 1) or 1)
    vins = [d.get(f'vin_{i}', ' ') or ' ' for i in range(1, qty + 1)]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Column widths matching original
    col_widths = {'A': 40, 'B': 15, 'C': 14, 'D': 12, 'E': 8, 'F': 14, 'G': 4, 'H': 14}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    row = 1

    def add(values, bold=False, green=False, size=10):
        nonlocal row
        while len(values) < 8:
            values.append(None)
        for col_idx, val in enumerate(values[:8], 1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.font = Font(bold=bold or green, size=size)
            if green:
                cell.fill = green_fill()
        row += 1

    def blank():
        add([None]*8)

    def section(title):
        add([title, None, None, None, None, None, None, None], bold=True, green=True)

    def item(desc, code='', note='', price=None, qty2=0, amount=None):
        if amount is None:
            amount = (qty2 or 0) * (price or 0)
        add([desc, code or None, note or None, price, qty2, amount, None, None])

    def total_row(label, val):
        nonlocal row
        ws.cell(row=row, column=4, value=label).font = bold_font()
        ws.cell(row=row, column=6, value=val).font = bold_font()
        row += 1

    # Title
    add(["Primetime X2C/Promaster"], bold=True, size=14)

    # Header info
    add(["Customer", d.get('organizationName', ' '), "Quantity", qty])
    add(["Address", d.get('address', ' '), "Chassis Release", d.get('chassisRelease', ' ')])
    add(["City/State", d.get('cityState', ' '), "Fin Code", d.get('finCode', ' ')])
    add(["Contact", d.get('contact', ' '), "Seat Material", d.get('seatMaterial', ' ')])
    add(["Phone", d.get('phone', ' '), "Date", d.get('date', ' ')])
    add(["Email", d.get('email', ' '), "Adaptive Mobility?", d.get('adaptiveMobility', ' ')])
    blank()
    for i, vin in enumerate(vins):
        label = "Vin #" if len(vins) == 1 else f"Vin #{i+1}"
        add([label, vin])
    blank()
    blank()

    # Column headers
    add(["ITEM DESCRIPTION", None, None, "Price", "Qty.", "Amount"], bold=True)
    blank()

    # CHASSIS
    section("CHASSIS")
    chassis_list = [
        ("2026 Low Roof Chassis  - Builders Prep", 54489),
        ("2026 Low Roof Chassis  - 12 Passenger", 55318),
        ("2026 Low Roof Chassis  - 15 Passenger", 56608),
        ("2026 Mid Roof Chassis  - 12 Passenger", 56030),
        ("2026 Mid Roof Chassis  - Builders Prep", 55201),
        ("2025 Promaster - Cargo", 49283),
    ]
    selected_chassis = d.get('chassis', '')
    for name, price in chassis_list:
        q = 1 if selected_chassis and (selected_chassis == name or selected_chassis.replace('  -', ' -') == name.replace('  -', ' -')) else 0
        item(name, price=price, qty2=q)
    item("Full Body Paint - OEM Only", code="Ford X2C", price=329, qty2=1 if d.get('fullBodyPaintOEM') == 'Yes' else 0)
    item("Full Body Paint - Non OEM Only", price=7995, qty2=1 if d.get('fullBodyPaintNonOEM') == 'Yes' else 0)
    blank()

    # INTERIOR
    section("INTERIOR")
    upfit_map = {
        "Interior Upfit - Ford Transit": 3995,
        "Interior Upfit - Ford Transit - Side Rear Lift": 4995,
        "Interior Upfit - Promaster Window": 7995,
        "Interior Upfit - Promaster LF": 11995,
    }
    selected_upfit = d.get('interiorUpfit', '').split(' (+')[0] if d.get('interiorUpfit') else ''
    for name, price in [
        ("Interior Upfit - Ford Transit", 3995),
        ("Interior Uppfit - Ford Transit - Side Rear Lift", 4995),
        ("Interior Upfit - Promaster Window", 7995),
        ("Interior Upfit - Promaster LF", 11995),
        ("Rear Storage Barrier", None),
        ("Storage Walker Mount", 395),
        ("Pa System with Internal/External Speaker", 475),
    ]:
        key = name.replace("Uppfit", "Upfit")
        q = 0
        if key == "Storage Walker Mount" and d.get('storageWalkerMount') == 'Yes': q = 1
        elif key == "Pa System with Internal/External Speaker" and d.get('paSystem') == 'Yes': q = 1
        elif selected_upfit and selected_upfit == key: q = 1
        item(name, price=price, qty2=q)
    blank()

    # RUNNING BOARD
    section("RUNNING BOARD")
    item("Passenger Running Board", price=390, qty2=1 if d.get('passengerRunningBoard') == 'Yes' else 0)
    item("Driver Running Board", price=5, qty2=1 if d.get('driverRunningBoard') == 'Yes' else 0)
    item("Rear Mud Flaps", price=75, qty2=1 if d.get('rearMudFlaps') == 'Yes' else 0)
    blank()

    # A/C
    section("A/C - Heat")
    ac = d.get('acHeat', '')
    item("Twin Air A/C- Heat - Dodge - Tie In", price=1895, qty2=1 if 'Twin' in ac else 0)
    item("Dual A/C Compressor", price=5595, qty2=1 if 'Dual' in ac else 0)
    item("OEM A/C - Heat", price=0, qty2=1 if not ac or 'OEM' in ac else 0)
    blank()

    # FLOORING
    section("FLOORING")
    fk = d.get('flooring', '').split(' (+')[0] if d.get('flooring') else ''
    for name, price, match in [
        ("Plwood Subfloor with Wood Grain Flooring", 1395, "Plywood Subfloor with Wood Grain Flooring"),
        ("Plwood Subfloor with Altro Flooring", 0, "Plywood Subfloor with Altro Flooring"),
        ("Modify Flooring - OEM Seat Package", 1295, "Modify Flooring - OEM Seat Package"),
        ("Pareto Floor -  Ford", 4995, "Pareto Floor - Ford"),
        ("Pareto Floor -  Dodge", 5995, "Pareto Floor - Dodge"),
    ]:
        item(name, price=price, qty2=1 if fk == match else 0)
    blank()

    # SEATING
    section("SEATING")
    item("Freedman SIngle GO Seat", price=827, qty2=int(d.get('seatSingleGO') or 0))
    item("Freedman Double GO Seat  ", price=1695, qty2=int(d.get('seatDoubleGO') or 0))
    item("Freedman Double GO Integrated Child Seat (1)", qty2=0)
    item("Freedman Double GO Integrated Child Seat (2)", qty2=0)
    item("Freedman Double Go Seat - Foldaway", price=2075, qty2=int(d.get('seatDoubleFoldaway') or 0))
    item("Freedman SIngle Go Seat - Foldaway", price=1195, qty2=int(d.get('seatSingleFoldaway') or 0))
    item("Pareto Seat Base", price=200, qty2=int(d.get('seatPareto') or 0))
    item("Child Seat Mounting Clips", qty2=0)
    item("Seat Belt Extensions", price=30, qty2=int(d.get('seatBeltExtQty') or 0))
    item("Freedman ARAC Preimeter Seating", qty2=int(d.get('seatARACPerimeter') or 0))
    item("Seating Arm Reast", price=60, qty2=int(d.get('seatArmRests') or 0))
    blank()

    # WHEELCHAIR DOOR
    section("WHEEL CHAIR DOOR")
    item('48x62 Manual Side Doors', price=7995, qty2=1 if d.get('wcDoor') == 'Yes' else 0)
    blank()

    # WHEELCHAIR LIFT
    section("WHEEL CHAIR LIFT")
    lift_map = [
        ("Braun Century 34 x 51 #800", 5819, "Braun Century 34x51 #800"),
        ("Braun Century 34 x 51 #1000", 5995, "Braun Century 34x51 #1000"),
        ("Braun Century 37x 54 #1000", 7195, "Braun Century 37x54 #1000"),
        ("Braun Century Rear Side Door 34 x 51 #1000", 6995, "Braun Century Rear Side Door 34x51 #1000"),
        ("Braun Millenium 34 x 51 #800", 5995, "Braun Millenium 34x51 #800"),
        ("Braun Millenium 34 x 51 #1000", 6995, "Braun Millenium 34x51 #1000"),
        ("Braun Shift N Step Lift (Plus Lift Cost)", 8699, "Braun Shift N Step Lift"),
    ]
    selected_lift = d.get('wcLift', '').split(' (+')[0] if d.get('wcLift') else ''
    for name, price, key in lift_map:
        item(name, price=price, qty2=1 if selected_lift == key else 0)
    item("ADA Interlock", price=695, qty2=1 if d.get('adaInterlock') == 'Yes' else 0)
    item("Passenger Call Bell System w Touch Pads", price=495, qty2=1 if d.get('passengerCallBell') == 'Yes' else 0)
    blank()

    # WC RESTRAINTS
    section("WHEEL CHAIR RESTRAINTS")
    item("L-Track - Per Track", price=150, qty2=int(d.get('lTrackQty') or 0))
    item("Shoulder Anchor Point - Per Position", price=0, qty2=0)
    item("Shoulder Anchor Point - Per Position - DRW", price=249, qty2=int(d.get('shoulderAnchor') or 0))
    item("Q Straint - L Track", price=595, qty2=int(d.get('qStraintLTrack') or 0))
    item("Slide N Click  Floor Mounts", price=52, qty2=int(d.get('slideNClick') or 0))
    item("Q-Straint - Slide N Click", price=723, qty2=int(d.get('qStraintSlide') or 0))
    blank()

    # DESTINATION SIGNAGE
    section("DESTINATION SIGNAGE")
    item("Front Destination Sign (TransSIgn)", price=3495, qty2=1 if d.get('frontDestSign') == 'Yes' else 0)
    item("Side Destination Sign (TransSign)", price=1494, qty2=1 if d.get('sideDestSign') == 'Yes' else 0)
    blank()

    # STANTIONS
    section("STANTIONS/POLES")
    item("Entrance Grab Bar", price=199, qty2=1 if d.get('entranceGrabBar') == 'Standard (+$199)' else 0)
    item("Entrance Grab Bar - Yellow", price=189, qty2=1 if d.get('entranceGrabBar') == 'Yellow (+$189)' else 0)
    item("Parallel Grab Bars - Entry Door", price=195, qty2=1 if d.get('parallelGrabBars') == 'Yes' else 0)
    item("Stantions  ", price=495, qty2=1 if d.get('stantions') == 'Yes' else 0)
    blank()

    # ENTRANCE DOOR
    section("ENTRANCE DOOR")
    ed = d.get('entranceDoor', '')
    item("Entrance Door", price=5295, qty2=1 if 'Standard' in ed else 0)
    item("Entrance Door - L.F.", price=6995, qty2=1 if 'L.F.' in ed else 0)
    item("Keyed Remote Entry", price=95, qty2=1 if d.get('keyedRemoteEntry') == 'Yes' else 0)
    item("Remote Entry", price=75, qty2=1 if d.get('remoteEntry') == 'Yes' else 0)
    blank()

    # SAFETY
    section("SAFETY ITEMS")
    item("Safety Kit - Fire Bottle, Triangle Kit, Backalarm", price=395, qty2=1 if d.get('safetyKit') == 'Yes' else 0)
    item("Transpec Roof Hatch", price=695, qty2=1 if d.get('roofHatch') == 'Yes' else 0)
    strobe = d.get('strobeLight', '')
    item("Strobe Light", code="Color", price=395, qty2=1 if 'Color' in strobe else 0)
    item("Vehicle Height Decal", price=20, qty2=1 if d.get('heightDecal') == 'Yes' else 0)
    item("Watch Your Step Decal", price=20, qty2=1 if d.get('watchStepDecal') == 'Yes' else 0)
    item("Strobe Light", note="Clear", price=395, qty2=1 if 'Clear' in strobe else 0)
    blank()

    # AUDIO
    section("AUDIO")
    item("PA System", price=395, qty2=1 if d.get('paSystemAudio') == 'Yes' else 0)
    item("External Speaker", price=50, qty2=1 if d.get('externalSpeaker') == 'Yes' else 0)
    item("Lockable Storage Box - Remove Copliot Seat - Wood", price=395, qty2=1 if d.get('lockableStorageWood') == 'Yes' else 0)
    item("Lockable Storage Box  - Remove Copliot Seat - Steel", price=495, qty2=1 if d.get('lockableStorageSteel') == 'Yes' else 0)
    blank()

    # ELECTRICAL
    section("ELECTRICAL")
    item("Upgraded Dome Light Package (6)", price=100, qty2=1 if d.get('upgradedDomeLights') == 'Yes' else 0)
    item("Passenger Call Bell System", price=495, qty2=0)
    item("Heated Step Well", price=347.5, qty2=1 if d.get('heatedStepWell') == 'Yes' else 0)
    item("USB Ports - Each", price=50, qty2=int(d.get('usbPorts') or 0))
    blank()

    # MISCELLANEOUS
    section("MISCELLANEOUS")
    item("Recertifications", price=395, qty2=0)
    item("Drop Ship Credit", price=-500, qty2=1)
    item("Freight", price=475, qty2=1)
    blank()

    # SPECIAL BUILDS
    section("SPECIAL BUILDS")
    sb = d.get('specialBuild', '')
    item("Modify for OEM Seating", note="10 Passenger", price=1295, qty2=1 if '10 Passenger' in sb else 0)
    item("Modify for OEM Seating", note="15 Passenger", price=3000, qty2=1 if '15 Passenger' in sb else 0)
    item("Seat Package B", price=5495, qty2=1 if 'Seat Package B' in sb else 0)
    item("Lockabable Stroage Box", price=395, qty2=1 if d.get('lockableStorageBox') == 'Yes' else 0)
    item("Fairbox Prewire", price=30, qty2=1 if d.get('fairboxPrewire') == 'Yes' else 0)
    blank()

    # SPECIAL NOTES
    section("SPECIAL NOTES")
    notes = d.get('specialNotes', '')
    notes_price = float(d.get('specialNotesPrice') or 0)
    item(notes or ' ', price=notes_price if notes else None, qty2=1 if notes else 0)
    blank()

    # Calculate totals
    pt = calculate_pt(d)
    bsi = calculate_bsi(d)
    grand = pt + bsi

    total_row("Primetime Total", pt)
    blank()

    # BSI SUPPLIED
    section("BSI SUPPLIED ITEMS")
    item("Basic Graphics  Package", price=325, qty2=1 if d.get('basicGraphics') == 'Yes' else 0)
    item("Graphic Package", qty2=0)
    item("BSI Add Ons - See Attached", price=350, qty2=1 if d.get('bsiAddOns') == 'Yes' else 0)
    item("Angeltrax 2 Camera System", price=2000, qty2=1 if d.get('angeltrax') == 'Yes' else 0)
    item("Undercoat", price=400, qty2=1 if d.get('undercoat') == 'Yes' else 0)
    item("Custom Graphic Package", price=0, qty2=1 if d.get('customGraphics') == 'Yes' else 0)
    item("OEM Seat Package", price=1200, qty2=1 if d.get('oemSeatPackage') == 'Yes' else 0)
    item("Class 2 Hitch with 4 Pin Connector", price=499, qty2=1 if d.get('classHitch') == 'Yes' else 0)
    item("Video System - 2-4 Camera/4 Channel", qty2=0)
    item("Spare Fuses", price=20, qty2=0)
    item("Rooftop School Transportation Sign", price=325, qty2=1 if d.get('schoolSign') == 'Yes' else 0)
    blank()

    section("INCENTIVE")
    item("Mobility - If Applicable", price=0, qty2=1 if d.get('mobilityIncentive') == 'Yes' else 0)
    total_row("BSI Supplied Total", bsi)
    total_row("Total", grand)
    add(["Qty.", None, None, None, qty, None])
    total_row("Total", grand * qty)

    if d.get('internalNotes') or d.get('arrivalAddOns'):
        blank()
        section("INTERNAL NOTES (BSI ONLY)")
        if d.get('internalNotes'):
            add(["Notes:", d.get('internalNotes')])
        if d.get('arrivalAddOns'):
            add(["Arrival Add-Ons:", d.get('arrivalAddOns')])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    org = (d.get('organizationName') or 'Build').replace(' ', '_')
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'BSI_{org}_BuildSheet.xlsx')


@app.route('/generate-proposal', methods=['POST', 'OPTIONS'])
def generate_proposal():
    if request.method == 'OPTIONS':
        return '', 204

    d = request.json
    qty = int(d.get('quantity', 1) or 1)
    vins = [d.get(f'vin_{i}', ' ') or ' ' for i in range(1, qty + 1)]

    pt = calculate_pt(d)
    bsi = calculate_bsi(d)
    grand = pt + bsi

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PROPOSAL"

    col_widths = {'A': 43, 'B': 32, 'C': 37, 'D': 36, 'E': 0.4, 'F': 24.7,
                  'G': 5.4, 'H': 49.7, 'I': 32, 'J': 48.6, 'K': 10.6, 'L': 9.1, 'M': 0.1}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    row = [1]

    def add(vals, bold=False, size=10):
        while len(vals) < 13:
            vals.append(None)
        for col_idx, val in enumerate(vals[:13], 1):
            cell = ws.cell(row=row[0], column=col_idx, value=val)
            if bold or size != 10:
                cell.font = Font(bold=bold, size=size)
        row[0] += 1

    def blank():
        add([None] * 13)

    def section(text, size=24):
        add([text], bold=True, size=size)

    chassis = d.get('chassis', '')
    fk = d.get('flooring', '').split(' (+')[0] if d.get('flooring') else ''
    ac = d.get('acHeat', '')
    strobe = d.get('strobeLight', '')
    ed = d.get('entranceDoor', '')
    selected_lift = d.get('wcLift', '').split(' (+')[0] if d.get('wcLift') else ''

    # Header
    add(["Customer", None, "Delivery Address", None, None, "Contact"], bold=True, size=24)
    add([d.get('organizationName', ''), None, d.get('address', ''), None, None, "contact name", None, d.get('contact', '')])
    add([None, None, d.get('cityState', ''), None, None, "contact phone #", None, d.get('phone', '')])
    add([None, None, None, None, None, "contact email", None, d.get('email', '')])
    blank()
    section("Vehicle / Contact Information", size=26)
    blank()
    add(["Model Year:", 2026, None, "Agreement #:"])
    add(["Vin #:", vins[0], None, "Delivery Time:"])
    add(["Graphic:", "Yes" if d.get('customGraphics') == 'Yes' else '', None, "Representative:", d.get('salesperson', '')])
    add(["Seat Material:", d.get('seatMaterial', 'Vinyl'), None, "Phone/Email:"])
    blank()
    section("Prime-Time SV Mobility Vans", size=28)
    for _ in range(9): blank()

    add(["MASTERS OF VAN CONVERSIONS"], bold=True)
    add(["ADA - FULL COMPLIANCE - American with Disabilities Act"])
    add(["FMVSS - FULL COMPLIANCE - Federal Motor Vehicle Safety Standards", None, None, "FORD QVM - FULL COMPLIANCE - Qualified Vehicle Modifier program"])
    section("STANDARD CHASSIS EQUIPMENT AND FEATURES")
    std = [
        ("3.5L PFDI V6 Gasoline Engine", "Safety Canopy Passenger Curtain Airbag System"),
        ("10 Speed Electronic Automatic Transmission", "Securilok Anti-Theft System"),
        ("Transmission Oil Cooler", "Ford Co-Pilot 360 - Side Wind Stabilization, Post Collision Braking"),
        ("Front and Rear Disc Brakes with Full ABS", "Ford Co-Pilot 360 - Anti-Roll Bar Mitigation"),
        ("235/65R16 All Season Tires", "Three (3) USB Charging Ports in Passenger Area"),
        ("Single 70AH Battery", "AdvanceTrac with RSC"),
        ("Heavy Duty Alternator 250 AMP", "Rear Window Defroster"),
        ("Power Windows and Door Locks", "Cruise Control"),
        ("Electronic Stability Control", "Lane-Keeping System & Forward Collision Warning"),
        ("Auto-High Beam Headlamps", "12 inch Center Display (SYNC4)"),
        ("2-Way Manual Driver Seat", "Rear Recovery Tow Hook"),
        (" ", " "),
    ]
    for a, b in std:
        add([a, None, None, b])

    section("PRIME-TIME SPECIALTY VANS", size=26)
    add([None, None, None, None, None, "Code", None, "Qty", None, "Extend"], bold=True)
    add(["MODEL", None, None, None, None, None, None, 1])
    add(["Ford Transit 136 inch Mid Roof Mobility Passenger Van", None, None, None, None, None, None, 0, "Included"])
    add(["Ford Transit 148 inch Mid Roof Mobility Passenger Van", None, None, None, None, None, None, 1 if "12 Passenger" in chassis else 0, "Included"])
    add(["Ford Transit 148 inch Extended Length High Roof Mobility Passenger Van", None, None, None, None, None, None, 1 if "15 Passenger" in chassis else 0, "Included"])
    notes = d.get('specialNotes', '')
    add(["SPECIAL BUILD ITEMS / INSTRUCTIONS", None, None, None, None, None, None, 1 if notes else 0])
    add([notes or None, None, None, None, None, None, None, 1 if notes else 0, "Included"])

    section("SIDEWALL / REARWALL / CEILING")
    add(["Ford OEM Cloth Interior", None, None, None, "05", "STD", None, 1, "Included"])
    add(["PSV Fully Insulated Walls and Ceiling", None, None, None, None, None, None, 1, "Included"])
    add(["PSV ABS Interior with Insulated Walls and Ceiling", None, None, None, None, None, None, 0, "Included"])

    section("VAN FLOORING")
    add(["Altro Anti-Slip Flooring", None, None, None, "05", 2248, None, 1 if "Altro" in fk else 0, "Included"])
    add(["3/4 inch Exterior Grade Plywood SubFloor", None, None, None, None, None, None, 1, "Included"])
    add(["Gerflor Anti-Slip Flooring", None, None, None, "05", 2824, None, 0, "Included"])
    add(["Altro Wood Safety Anti-Slip Flooring", None, None, None, "05", 2175, None, 1 if "Wood" in fk else 0, "Included"])

    section("EXTERIOR GRAPHICS AND PAINT")
    add(["Full Body Exterior Paint", None, None, None, None, None, None, 1 if d.get('fullBodyPaintOEM') == 'Yes' or d.get('fullBodyPaintNonOEM') == 'Yes' else 0, "Included"])
    add(["Custom Graphic Package", None, None, None, "05", 2235, None, 1 if d.get('customGraphics') == 'Yes' else 0, "Included"])

    section("TRANSIT VAN CHASSIS")
    for item_name in ["Rear Wheel Drive Vehicle", "Safety Canopy Passenger Curtain Airbag System",
                       "Ford Co-Pilot 360 - Anti-Roll and Side Wind Stabilization",
                       "Electronic Stability Control - Lane Keeping System",
                       "Pre-Collision Assist with Automatic Emergency Braking (AEB)",
                       "GPS Navigation", "Two (2) Keyless Remote Entry Fobs"]:
        add([item_name, None, None, None, None, None, None, 1, "Included"])
    add(["Heavy Duty Anti-Slip Running Board on Passenger Side", None, None, None, "05", 2623, None, 1 if d.get('passengerRunningBoard') == 'Yes' else 0, "Included"])
    add(["Heavy Duty Anti-Slip Running Board on Driver Side", None, None, None, "05", 2116, None, 1 if d.get('driverRunningBoard') == 'Yes' else 0, "Included"])

    section("ENVIRONMENTAL CONTROL")
    add(["Ford OEM Front and Rear Air-Conditioning", None, None, None, None, None, None, 1 if not ac or 'OEM' in ac else 0, "Included"])
    add(["Ford Auxiliary Air-Conditioning System Upgrade - 32,000 btu System", None, None, None, None, None, None, 1 if 'Twin' in ac else 0, "Included"])

    section("REAR (AUXILIARY) HEATING")
    add(["Ford OEM Front Mount Floor Heater", None, None, None, "05", 2044, None, 1, "Included"])

    section("DESTINATION SIGNS & WINDOWS")
    add(["Front Destination Sign (TransSign)", None, None, None, "05", 2033, None, 1 if d.get('frontDestSign') == 'Yes' else 0, "Included"])
    add(["Side Destination Sign (TransSign)", None, None, None, None, None, None, 1 if d.get('sideDestSign') == 'Yes' else 0, "Included"])

    section("EXTERIOR LIGHTS")
    add(["Rear Center Mount Brake Light", None, None, None, "05", 2802, None, 1, "Included"])
    add(["Roof Mounted Strobe Light", None, None, None, "05", 2427, None, 1 if strobe and strobe != 'No' else 0, "Included"])

    section("INTERIOR LIGHTS")
    add(["Door Activated Interior Lights", None, None, None, None, None, None, 1, "Included"])
    add(["LED Overhead Interior Strip Lights", None, None, None, "05", 2262, None, 1 if d.get('upgradedDomeLights') == 'Yes' else 0, "Included"])

    section("AUDIO / VISUAL")
    add(["Ford Transit OEM 10.2 inch Display Media Center", None, None, None, "05", 2158, None, 1, "Included"])
    add(["PA System with 2 Speakers (Independent of Radio)", None, None, None, "05", 2388, None, 1 if d.get('paSystem') == 'Yes' or d.get('paSystemAudio') == 'Yes' else 0, "Included"])
    add(["External Speaker with On/Off Switch", None, None, None, "05", 2556, None, 1 if d.get('externalSpeaker') == 'Yes' else 0, "Included"])

    section("DOORS / WINDOWS / ROOF HATCHES")
    add(["OEM Passenger Sliding Entry Door", None, None, None, "05", 2887, None, 1 if not ed or ed == 'None' else 0, "Included"])
    add(["Electric Bi-Fold Bus Passenger Entry Door upgrade", None, None, None, "05", 2056, None, 1 if 'Bi Fold' in ed else 0, "Included"])
    add(["Remote Entry Key Fob for Passenger entry door", None, None, None, "05", 2241, None, 1 if d.get('remoteEntry') == 'Yes' else 0, "Included"])
    add(["A&M Remote Entry Door Keypad", None, None, None, "05", 2876, None, 1 if d.get('keyedRemoteEntry') == 'Yes' else 0, "Included"])
    add(["Roof Hatch - Transpec", None, None, None, "05", 2133, None, 1 if d.get('roofHatch') == 'Yes' else 0, "Included"])

    section("BRAUNABILITY WHEELCHAIR LIFTS")
    lift_list = [
        ("Braun Century Series Wheelchair Lift - 800lb (34x51)", "Braun Century 34x51 #800"),
        ("Braun Century Series Wheelchair Lift - 1000lb (34x51)", "Braun Century 34x51 #1000"),
        ("Braun Century Series Wheelchair Lift - 1000lb (37x54)", "Braun Century 37x54 #1000"),
        ("Braun Century Rear Side Door - 1000lb (34x51)", "Braun Century Rear Side Door 34x51 #1000"),
        ("Braun Millenium Series Wheelchair Lift - 800lb (34x51)", "Braun Millenium 34x51 #800"),
        ("Braun Millenium Series Wheelchair Lift - 1000lb (34x51)", "Braun Millenium 34x51 #1000"),
        ("Braun Shift N Step Lift", "Braun Shift N Step Lift"),
    ]
    for name, key in lift_list:
        add([name, None, None, None, "05", None, None, 1 if selected_lift == key else 0, "Included"])

    add(["WHEELCHAIR LIFT FAST IDLE WITH INTERLOCK"], bold=True)
    add(["ADA Fast Idle with Lift Interlock", None, None, None, "05", 2714, None, 1 if d.get('adaInterlock') == 'Yes' else 0, "Included"])

    add(["Q-STRAINT WHEELCHAIR SECUREMENTS & ACCESSORIES"], bold=True)
    add(["Q-Straint Slide-N-Click Automatic Wheelchair Securements", None, None, None, "05", 2245, None, int(d.get('qStraintSlide') or 0), "Included"])
    add(["Q-Straint Automatic Wheelchair Securements (L-Track)", None, None, None, None, None, None, int(d.get('qStraintLTrack') or 0), "Included"])
    add(["Full-Length L-Track mounted to Sidewall", None, None, None, None, None, None, int(d.get('lTrackQty') or 0), "Included"])

    section("SAFETY OPTIONS")
    safety = d.get('safetyKit') == 'Yes'
    add(["5 lb Fire Extinguisher", None, None, None, "05", 2089, None, 1 if safety else 0, "Included"])
    add(["16 Unit First Aid Kit", None, None, None, "05", 2090, None, 1 if safety else 0, "Included"])
    add(["Emergency Triangle Kit", None, None, None, "05", 2091, None, 1 if safety else 0, "Included"])
    add(["Back-Up Alarm", None, None, None, "05", 2092, None, 1 if safety else 0, "Included"])
    add(["Ford OEM Back-Up Camera System", None, None, None, "05", 2123, None, 1, "Included"])
    add(["Decal Please Watch Your Step", None, None, None, None, None, None, 1 if d.get('watchStepDecal') == 'Yes' else 0, "Included"])
    add(["Decal Vehicle Height Sticker", None, None, None, None, None, None, 1 if d.get('heightDecal') == 'Yes' else 0, "Included"])

    section("GRAB RAILS / STANCHIONS / PANELS")
    add(["Stainless Steel Right Hand Entry Vertical Grab Rail", None, None, None, "05", 2049, None, 1, "Included"])
    add(["Parallel Grab Bars - Entry Door", None, None, None, None, None, None, 1 if d.get('parallelGrabBars') == 'Yes' else 0, "Included"])
    add(["Stanchion and Modesty Panel Behind Driver", None, None, None, "05", 2857, None, 1 if d.get('stantions') == 'Yes' else 0, "Included"])

    section("PASSENGER SEATING")
    add(["High-Back Double Seat w/ Vinyl cover", None, None, None, "05", 2065, None, int(d.get('seatDoubleGO') or 0), "Included"])
    add(["High-Back Single Seat w/ Vinyl cover", None, None, None, "05", 2066, None, int(d.get('seatSingleGO') or 0), "Included"])
    add(["Mid-High Double Foldaway Seat w/ Vinyl cover", None, None, None, "05", 2067, None, int(d.get('seatDoubleFoldaway') or 0), "Included"])
    add(["Mid-High Single Foldaway Seat w/ Vinyl cover", None, None, None, "05", 2068, None, int(d.get('seatSingleFoldaway') or 0), "Included"])
    add(["Black Flip-Up Armrests on Aisle Seats", None, None, None, "05", 2077, None, int(d.get('seatArmRests') or 0), "Included"])

    add(["SEAT BELTS"], bold=True)
    add(["3-Point Lap and Shoulder Harness Seat Belts for All Passenger Seats", None, None, None, "05", 2086, None, 1, "Included"])
    add(["Seat Belt Extension 12 inch", None, None, None, "05", 2087, None, int(d.get('seatBeltExtQty') or 0), "Included"])

    # Terms
    section("TERMS AND CONDITIONS", size=26)
    blank()
    add([None]*7 + ["Vehicle Price:", None, pt])
    blank()
    add(["Terms:  ", "Due Upon Receipt", None, None, None, "Mobility Rebate:", None, "Large Fleet Discount", None, ""])
    blank()
    mobility_rebate = -6451 if d.get('mobilityIncentive') == 'Yes' else 0
    add(["Deposit:  ", "10% Down Payment Due at Signing", None, None, None, "Chassis Rebate:", None, "Accessibility Rebate:", None, mobility_rebate])
    blank()
    add(["Quote Valid For:", "30 Days from Date", None, None, None, None, None, "Subtotal:", None, pt + mobility_rebate])
    blank()
    add(["This Agreement is Confidential and not Subject to Distribution", None, None, None, None, None, None, "Delivery:", None, 475])
    blank()
    blank()
    blank()
    add(["Customer Signature", None, None, None, None, None, None, "Unit Total:", None, grand])
    blank()
    add(["Printed Name", None, None, None, None, None, None, " "])
    add([None]*7 + ["Quantity:", qty])
    add(["Date"])
    add([None]*7 + ["Net Due:", grand * qty])
    blank()
    add(["To Purchase this Vehicle, please sign this intent to purchase agreement. By Signing this agreement the Signer understands they are entering a contract to purchase the product/service presented."])

    for _ in range(80): blank()
    add(["Customer Order Information"], bold=True)
    add(["Please fill in the information below when placing your order."])
    add(["Customer Name to Read on Title:"])
    add(["Customer Address to Read on Title:"])
    add(["Federal Tax ID#:"])
    add(["Are you Tax-Exempted?  Yes/No"])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    org = (d.get('organizationName') or 'Build').replace(' ', '_')
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'BSI_{org}_Proposal.xlsx')


def calculate_pt(d):
    total = 0
    chassis_prices = {
        "2026 Low Roof Chassis  - Builders Prep": 54489,
        "2026 Low Roof Chassis  - 12 Passenger": 55318,
        "2026 Low Roof Chassis  - 15 Passenger": 56608,
        "2026 Mid Roof Chassis  - 12 Passenger": 56030,
        "2026 Mid Roof Chassis  - Builders Prep": 55201,
        "2025 Promaster - Cargo": 49283,
    }
    if d.get('chassis') in chassis_prices:
        total += chassis_prices[d['chassis']]
    if d.get('fullBodyPaintOEM') == 'Yes': total += 329
    if d.get('fullBodyPaintNonOEM') == 'Yes': total += 7995

    upfit_prices = {"Interior Upfit - Ford Transit": 3995, "Interior Upfit - Ford Transit - Side Rear Lift": 4995,
                    "Interior Upfit - Promaster Window": 7995, "Interior Upfit - Promaster LF": 11995}
    uk = d.get('interiorUpfit', '').split(' (+')[0] if d.get('interiorUpfit') else ''
    if uk in upfit_prices: total += upfit_prices[uk]

    if d.get('paSystem') == 'Yes': total += 475
    if d.get('storageWalkerMount') == 'Yes': total += 395
    if d.get('passengerRunningBoard') == 'Yes': total += 390
    if d.get('driverRunningBoard') == 'Yes': total += 5
    if d.get('rearMudFlaps') == 'Yes': total += 75

    ac = d.get('acHeat', '')
    if 'Twin' in ac: total += 1895
    elif 'Dual' in ac: total += 5595

    floor_prices = {"Plywood Subfloor with Wood Grain Flooring": 1395, "Pareto Floor - Ford": 4995,
                    "Pareto Floor - Dodge": 5995, "Modify Flooring - OEM Seat Package": 1295}
    fk = d.get('flooring', '').split(' (+')[0] if d.get('flooring') else ''
    if fk in floor_prices: total += floor_prices[fk]

    total += int(d.get('seatSingleGO') or 0) * 827
    total += int(d.get('seatDoubleGO') or 0) * 1695
    total += int(d.get('seatDoubleFoldaway') or 0) * 2075
    total += int(d.get('seatSingleFoldaway') or 0) * 1195
    total += int(d.get('seatPareto') or 0) * 200
    total += int(d.get('seatArmRests') or 0) * 60
    total += int(d.get('seatBeltExtQty') or 0) * 30

    if d.get('wcDoor') == 'Yes': total += 7995
    lift_prices = {"Braun Century 34x51 #800": 5819, "Braun Century 34x51 #1000": 5995,
                   "Braun Century 37x54 #1000": 7195, "Braun Century Rear Side Door 34x51 #1000": 6995,
                   "Braun Millenium 34x51 #800": 5995, "Braun Millenium 34x51 #1000": 6995, "Braun Shift N Step Lift": 8699}
    lk = d.get('wcLift', '').split(' (+')[0] if d.get('wcLift') else ''
    if lk in lift_prices: total += lift_prices[lk]
    if d.get('adaInterlock') == 'Yes': total += 695
    if d.get('passengerCallBell') == 'Yes': total += 495

    total += int(d.get('lTrackQty') or 0) * 150
    total += int(d.get('shoulderAnchor') or 0) * 249
    total += int(d.get('qStraintLTrack') or 0) * 595
    total += int(d.get('slideNClick') or 0) * 52
    total += int(d.get('qStraintSlide') or 0) * 723

    if d.get('frontDestSign') == 'Yes': total += 3495
    if d.get('sideDestSign') == 'Yes': total += 1494

    ed = d.get('entranceDoor', '')
    if 'Standard' in ed: total += 5295
    elif 'L.F.' in ed: total += 6995
    if d.get('remoteEntry') == 'Yes': total += 75
    if d.get('keyedRemoteEntry') == 'Yes': total += 95
    if d.get('entranceGrabBar') == 'Standard (+$199)': total += 199
    if d.get('entranceGrabBar') == 'Yellow (+$189)': total += 189
    if d.get('parallelGrabBars') == 'Yes': total += 195
    if d.get('stantions') == 'Yes': total += 495

    if d.get('safetyKit') == 'Yes': total += 395
    if d.get('roofHatch') == 'Yes': total += 695
    if d.get('strobeLight') and d.get('strobeLight') != 'No': total += 395
    if d.get('heightDecal') == 'Yes': total += 20
    if d.get('watchStepDecal') == 'Yes': total += 20

    if d.get('upgradedDomeLights') == 'Yes': total += 100
    if d.get('heatedStepWell') == 'Yes': total += 347.5
    total += int(d.get('usbPorts') or 0) * 50
    if d.get('paSystemAudio') == 'Yes': total += 395
    if d.get('externalSpeaker') == 'Yes': total += 50
    if d.get('lockableStorageWood') == 'Yes': total += 395
    if d.get('lockableStorageSteel') == 'Yes': total += 495

    sb = d.get('specialBuild', '')
    if '10 Passenger' in sb: total += 1295
    elif '15 Passenger' in sb: total += 3000
    elif 'Seat Package B' in sb: total += 5495
    if d.get('lockableStorageBox') == 'Yes': total += 395
    if d.get('fairboxPrewire') == 'Yes': total += 30
    total += float(d.get('specialNotesPrice') or 0)

    total += -500 + 475  # drop ship credit + freight
    return total


def calculate_bsi(d):
    total = 0
    if d.get('basicGraphics') == 'Yes': total += 325
    if d.get('bsiAddOns') == 'Yes': total += 350
    if d.get('angeltrax') == 'Yes': total += 2000
    if d.get('undercoat') == 'Yes': total += 400
    if d.get('oemSeatPackage') == 'Yes': total += 1200
    if d.get('classHitch') == 'Yes': total += 499
    if d.get('schoolSign') == 'Yes': total += 325
    return total


if __name__ == '__main__':
    app.run(debug=True, port=5000)
