import openpyxl
import pandas as pd
from itertools import combinations_with_replacement
from pyscript import document, when

# ---------------------------------------------------------
# 1. DYNAMIC EXCEL MAPPING (Loads directly from Excel)
# ---------------------------------------------------------
df_sheet1 = pd.read_excel("BELLOW_CALCULATOR.xlsx", sheet_name="Sheet1")

# Forward fill missing process sizes for grouped rows
df_sheet1["Process Size"] = df_sheet1["Process Size"].ffill()
df_clean = df_sheet1[df_sheet1["Process Size"] != "-"].copy()

OPTIONS_MAP = {}
DESCRIPTION_MAP = {}
DEFLECTION_MAP = {}

# Map load cycle UI values directly to Excel column names
CYCLE_COLUMNS = {
    "10000": "Total Deflection 10000 Load Cycles, EJMA (mm) ",
    "4000": "Total Deflection 4000 Load Cycles, EJMA (mm) ",
    "1700": "Total Deflection 1700 Load Cycles, EJMA (mm) ",
    "1000": "Total Deflection 1000 Load Cycles, EJMA (mm) "
}

for _, row in df_clean.iterrows():
    size = str(row["Process Size"]).strip()
    try:
        corr = int(row["Corrugation (W)"])
    except (ValueError, TypeError):
        continue

    corr_str = str(corr)
    desc = str(row["Description"]).strip()

    if size not in OPTIONS_MAP:
        OPTIONS_MAP[size] = []
    if corr not in OPTIONS_MAP[size]:
        OPTIONS_MAP[size].append(corr)

    DESCRIPTION_MAP[(size, corr_str)] = desc

    for cycle_key, col_name in CYCLE_COLUMNS.items():
        if col_name in row:
            DEFLECTION_MAP[(size, corr_str, cycle_key)] = row[col_name]

# ---------------------------------------------------------
# 2. UI HELPER FUNCTIONS & EVENT LISTENERS
# ---------------------------------------------------------
def update_corrugation_options(selected_size):
    """Update options in the second dropdown based on the dynamic map."""
    corr_select = document.querySelector("#corrugation")
    corr_select.innerHTML = ""
    
    corrugations = OPTIONS_MAP.get(selected_size, [])
    for val in corrugations:
        option_el = document.createElement("option")
        option_el.value = str(val)
        option_el.innerText = str(val)
        corr_select.appendChild(option_el)


@when("change", "#process_size")
def on_process_size_change(event):
    selected_size = event.target.value
    update_corrugation_options(selected_size)


# Populate initial corrugation options on page load
initial_size = document.querySelector("#process_size").value
update_corrugation_options(initial_size)

# ---------------------------------------------------------
# 3. LN2 PROPERTIES & CONTRACTION CALCULATION
# ---------------------------------------------------------
df_ln2 = pd.read_excel(
    "BELLOW_CALCULATOR.xlsx",
    sheet_name="Properties of LN2",
    header=3
)

pressure = df_ln2["Pressure (bar)"].tolist()
temperature = df_ln2["Temperature (K)"].tolist()


def pressure_to_temperature(P):
    """Linear interpolation logic."""
    for i in range(len(pressure) - 1):
        if pressure[i] <= P <= pressure[i + 1]:
            X1, X2 = pressure[i], pressure[i + 1]
            Y1, Y2 = temperature[i], temperature[i + 1]
            return Y1 + ((Y2 - Y1) / (X2 - X1)) * (P - X1)
    return None


def obtain_contraction(T):
    coeff_a = -295.46
    coeff_b = -0.40518
    coeff_c = 0.0094014
    coeff_d = -0.000021098
    coeff_e = 0.00000001878
    
    calc_contraction = abs((coeff_a + coeff_b*T + coeff_c*(T**2) + coeff_d*(T**3) + coeff_e*(T**4)) * 0.00001 * 1000)
    return calc_contraction


def update_excel_constants_sheet(temp_k):
    """Loads the workbook using openpyxl, updates H3, and saves it."""
    wb = openpyxl.load_workbook("BELLOW_CALCULATOR.xlsx")
    wb["Constants"]["H3"] = temp_k
    wb.save("BELLOW_CALCULATOR.xlsx")


def calculate_combination_w(exact_w, available_corrugations):

    available_corrugations = sorted(available_corrugations)

    if not available_corrugations:
        return None

    # --------------------------------------------------
    # FIRST: Always take the largest corrugation once
    # --------------------------------------------------
    largest_w = available_corrugations[-1]

    remaining_w = exact_w - largest_w

    # If largest corrugation alone is enough
    # within the 0.5W margin
    if remaining_w <= 0.5:
        return [largest_w]

    # --------------------------------------------------
    # SECOND: Find the minimum number of additional
    # corrugations required
    # --------------------------------------------------
    min_w = available_corrugations[0]

    max_quantity = int(remaining_w / min_w) + 1

    for quantity in range(1, max_quantity + 1):

        best_combination = None
        best_total = float("inf")

        # Try every possible combination using this quantity
        for combination in combinations_with_replacement(
            available_corrugations, quantity
        ):

            total = sum(combination)

            # --------------------------------------------------
            # 0.5W MARGIN
            # Combination is acceptable if it reaches the
            # required W OR is within 0.5W below it
            # --------------------------------------------------
            if total >= remaining_w - 0.5:

                # Among combinations with the same quantity,
                # choose the one closest to the required W
                difference = abs(total - remaining_w)

                if difference < abs(best_total - remaining_w):
                    best_total = total
                    best_combination = combination

        # The first valid quantity is automatically
        # the minimum quantity required
        if best_combination is not None:

            return [largest_w] + list(best_combination)

    return None

# PART 1 RETURNS MAXIMUM LENGTH
@when("click", "#calculate-button")
def calculate_click(event):
    output_div = document.querySelector("#output")
    process_size = document.querySelector("#process_size").value
    corrugation = document.querySelector("#corrugation").value
    p_input = document.querySelector("#P").value
    load_cycles = document.querySelector("#load_cycles").value

    if not p_input:
        output_div.innerText = "Error: Please enter a pressure value."
        return

    try:
        P = float(p_input)
    except ValueError:
        output_div.innerText = "Error: Invalid number format."
        return

    if P < pressure[0] or P > pressure[-1]:
        output_div.innerText = (
            f"Error: Pressure must be between {pressure[0]} and {pressure[-1]} bar."
        )
        return

    T = pressure_to_temperature(P)
    
    if T is not None:
        # Update the Constants sheet in the virtual filesystem with calculated temperature
        update_excel_constants_sheet(T)

        contraction = obtain_contraction(T)
        description = DESCRIPTION_MAP.get((process_size, corrugation), f"{process_size} ({corrugation}W)")
        raw_deflection = DEFLECTION_MAP.get((process_size, corrugation, load_cycles), None)

        if raw_deflection is not None and not pd.isna(raw_deflection):
            total_deflection = float(raw_deflection)
            final_length = (total_deflection / contraction) * 1000  # in mm
            
            output_div.innerHTML = (
                f"<strong>Description:</strong> {description}<br>"
                f"<strong>Temperature:</strong> {T:.5f} K<br>"
                f"<strong>Contraction per metre:</strong> {contraction:.5f} mm<br>"
                f"<strong>Total Deflection per metre:</strong> {total_deflection:.2f} mm<br><br>"
                f"<div style='font-size: 1.4em; color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px;'>"
                f"<strong>Final Length: {final_length:.3f} mm</strong>"
                f"</div>"
            )
        else:
            output_div.innerHTML = (
                f"<strong>Description:</strong> {description}<br>"
                f"<strong>Temperature:</strong> {T:.5f} K<br>"
                f"<strong>Contraction:</strong> {contraction:.5f} mm<br>"
                f"<strong>Total Deflection:</strong> N/A<br>"
            )
    else:
        output_div.innerText = "Error: Unable to interpolate temperature for the given pressure."


# Correction factor map for load cycles
K_DELTA_N = {
    "10000": 1.0,
    "4000": 1.2,
    "1700": 1.4,
    "1000": 1.6
}

# Base max deflection per corrugation at 10000 cycles for each Process Size
# (Extracted directly from column: 'Maximum Deflection Allowed per Corrugation')
PER_CORRUGATION_DEFLECTION_10000 = {
    "R05T": 0.34,
    "R10T": 0.60,
    "R15T": 0.63,
    "R15P": 0.67,
    "R20P": 0.90,
    "R25P": 0.85,
    "R30P": 0.90,
    "R40P": 0.90
}


# PART 2 RETURNS CORRUGATION
@when("click", "#calculate-button2")
def calculate_click2(event):
    output2_div = document.querySelector("#output2")
    process_size2 = document.querySelector("#process_size2").value
    load_cycles2 = str(document.querySelector("#load_cycles2").value)
    length_input2 = document.querySelector("#length2").value
    p_input2 = document.querySelector("#P2").value

    # 1. Input Validation
    if not length_input2 or not p_input2:
        output2_div.innerText = "Error: Please enter both desired length and pressure."
        return

    try:
        desired_length = float(length_input2)
        P2 = float(p_input2)
    except ValueError:
        output2_div.innerText = "Error: Invalid number format."
        return

    if desired_length <= 0:
        output2_div.innerText = "Error: Desired length must be greater than 0."
        return

    if P2 < pressure[0] or P2 > pressure[-1]:
        output2_div.innerText = f"Error: Pressure must be between {pressure[0]} and {pressure[-1]} bar."
        return

    # 2. Temperature & Thermal Contraction
    T2 = pressure_to_temperature(P2)
    if T2 is None:
        output2_div.innerText = "Error: Temperature interpolation failed."
        return

    update_excel_constants_sheet(T2)
    contraction_per_m = obtain_contraction(T2)

    required_deflection = (desired_length / 1000.0) * contraction_per_m

    if required_deflection < 1:
        output2_div.innerHTML = (
            f"<div style='font-size: 1.4em; color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px;'>"
            f"Bellow not required, required deflection is {required_deflection:.2f} mm, under 1mm. Use engineering methods to combat contraction."
            f"</div>"
        )
        return

    # 3. Calculate Corrugation (W)
    base_deflection = PER_CORRUGATION_DEFLECTION_10000.get(process_size2, None)
    factor = K_DELTA_N.get(load_cycles2, 1.0)

    if base_deflection is None:
        output2_div.innerText = "Error: Base deflection per corrugation not found."
        return

    # Deflection allowed per corrugation for chosen load cycles
    deflection_per_corrugation = base_deflection * factor

    # Calculated exact corrugations needed
    exact_w = required_deflection / deflection_per_corrugation

    # 4. Select Corrugation Combination from Excel List (OPTIONS_MAP)
    raw_options = OPTIONS_MAP.get(process_size2, [])
    available_corrugations = sorted([int(x) for x in raw_options if str(x).isdigit()])

    if not available_corrugations:
        output2_div.innerText = f"Error: No available corrugations found for {process_size2}."
        return

    # Find the minimum quantity combination,
    # always taking the largest corrugation first
    combination = calculate_combination_w(
        exact_w,
        available_corrugations
    )

    if combination is None:
        output2_div.innerHTML = (
            f"<strong>Process Size:</strong> {process_size2}<br>"
            f"<strong>Operating Temperature:</strong> {T2:.5f} K<br>"
            f"<strong>Required Thermal Contraction:</strong> {required_deflection:.2f} mm<br>"
            f"<strong>Deflection per Corrugation ({load_cycles2} cycles):</strong> {deflection_per_corrugation:.2f} mm<br>"
            f"<strong>Exact Calculated Corrugations:</strong> {exact_w:.3f}<br><br>"
            f"<div style='font-size: 1.2em; color: #721c24; background-color: #f8d7da; padding: 10px; border-radius: 5px; border: 1px solid #f5c6cb;'>"
            f"<strong>Error: Unable to find a suitable corrugation combination.</strong>"
            f"</div>"
        )
        return

    combination_text = " + ".join(
        f"{w}W" for w in combination
    )

    combination_total = sum(combination)
    combination_excess = combination_total - exact_w

    # 5. Display Results
    output2_div.innerHTML = (
        f"<strong>Process Size:</strong> {process_size2}<br>"
        f"<strong>Operating Temperature:</strong> {T2:.5f} K<br>"
        f"<strong>Required Thermal Contraction:</strong> {required_deflection:.2f} mm<br>"
        f"<strong>Deflection per Corrugation ({load_cycles2} cycles):</strong> {deflection_per_corrugation:.2f} mm<br>"
        f"<strong>Exact Calculated Corrugations:</strong> {exact_w:.3f}<br><br>"
        f"<div style='font-size: 1.4em; color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px;'>"
        f"<strong>Recommended Combination: {combination_text}</strong><br>"
        f"<strong>Total W: {combination_total}W</strong><br>"
        f"<strong>Excess: {combination_excess:.3f}W</strong>"
        f"</div>"
    )


creds_div = document.querySelector("#creds")
creds_div.innerHTML = (
    f"<div style='font-size: 0.9em; color: #052569; margin-top: 10px;'>"
   # f"Note: <br>The recommended corrugation is based on the closest available option from the Excel data and Witzenmann Bellow Catalogue."
   # f"<br>"
    #f"Temperature is calculated using linear interpolation (small scale data) from the LN2 properties table. <br>"
   # f"Source: Eric W. Lemmon, Ian H. Bell, Marcia L. Huber, and Mark O. McLinden, \"Thermophysical Properties of Fluid Systems\" in NIST Chemistry WebBook, NIST Standard Reference Database Number 69, Eds. P.J. Linstrom and W.G. Mallard, National Institute of Standards and Technology, Gaithersburg MD, 20899, https://doi.org/10.18434/T4D303, (retrieved August 6, 2026).)"
   # f"<br>"
    f"Programmed by TJQ | CSM"
    f"<br>"
    f"</div>"
)