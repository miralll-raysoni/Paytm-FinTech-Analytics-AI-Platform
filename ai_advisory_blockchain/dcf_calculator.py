import importlib.util
import math
from pathlib import Path

# Load seed modules dynamically
PART_C_DIR = Path(
    "/Users/miralraysoni/Downloads/BitSoM - Fintech and AI/Final Assessment/Part C"
)


def load_module_from_path(module_name: str, file_path: Path):
    if not file_path.exists():
        raise FileNotFoundError(f"Required seed file not found: {file_path}")
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


stock_mod = load_module_from_path("stock_universe", PART_C_DIR / "stock_universe.py")
STOCK_UNIVERSE = stock_mod.STOCK_UNIVERSE
R_F = stock_mod.RISK_FREE_RATE  # 0.07 (7.0%)
E_RM = stock_mod.MARKET_RETURN  # 0.13 (13.0%)

# 1. Base FCFF Parameters (INR Crores)
EBIT = 500.0
TAX_RATE = 0.25
DEP_AMORT = 50.0
CAPEX = 80.0
DELTA_NWC = 20.0

# 2. Capital Structure & Cost of Capital (PAYFIN beta = 1.35)
BETA_PAYFIN = STOCK_UNIVERSE["PAYFIN"]["beta"]  # 1.35
COST_OF_EQUITY = R_F + BETA_PAYFIN * (E_RM - R_F)  # 15.10%
PRE_TAX_COST_DEBT = 0.08
AFTER_TAX_COST_DEBT = PRE_TAX_COST_DEBT * (1 - TAX_RATE)  # 6.00%
WEIGHT_EQUITY = 0.70
WEIGHT_DEBT = 0.30

# WACC = 0.70 * 15.1% + 0.30 * 6.0% = 12.37%
BASE_WACC = (WEIGHT_EQUITY * COST_OF_EQUITY) + (WEIGHT_DEBT * AFTER_TAX_COST_DEBT)

# 3. Growth Rates
INITIAL_GROWTH = 0.12  # 12.0%
TERMINAL_GROWTH = 0.04  # 4.0%

# 4. Peer EBITDA Multiple
ILLUSTRATIVE_EBITDA = EBIT + DEP_AMORT  # ₹550.0 Cr
EV_EBITDA_MULTIPLE = 12.0


def compute_fcff_base() -> float:
    return (EBIT * (1 - TAX_RATE)) + DEP_AMORT - CAPEX - DELTA_NWC


def calculate_dcf_ev(
    wacc: float, g_term: float, base_fcff: float, g_init: float, n_years: int = 5
) -> dict:
    fcff_projections = []
    pv_fcff_list = []

    current_fcff = base_fcff
    for t in range(1, n_years + 1):
        g_t = g_init - ((g_init - g_term) * (t - 1) / (n_years - 1))
        current_fcff *= 1 + g_t
        fcff_projections.append(current_fcff)
        pv_fcff_list.append(current_fcff / ((1 + wacc) ** t))

    # Year 6 Cash Flow for Terminal Value
    fcff_yr6 = fcff_projections[-1] * (1 + g_term)
    terminal_value = fcff_yr6 / (wacc - g_term)
    pv_terminal_value = terminal_value / ((1 + wacc) ** n_years)

    sum_pv_fcff = sum(pv_fcff_list)
    enterprise_value = sum_pv_fcff + pv_terminal_value

    return {
        "projections": fcff_projections,
        "pv_fcff": sum_pv_fcff,
        "terminal_value": terminal_value,
        "pv_terminal_value": pv_terminal_value,
        "enterprise_value": enterprise_value,
    }


def generate_sensitivity_grid(base_wacc: float, base_g: float, base_fcff: float):
    wacc_steps = [base_wacc - 0.01, base_wacc, base_wacc + 0.01]
    g_steps = [base_g - 0.01, base_g, base_g + 0.01]

    worst_case_spread = (base_wacc - 0.01) - (base_g + 0.01)
    if worst_case_spread < 0.01:
        raise ValueError("Self-check failed: Spread < 1.0%")

    grid = {}
    for w in wacc_steps:
        grid[w] = {}
        for g in g_steps:
            res = calculate_dcf_ev(
                wacc=w, g_term=g, base_fcff=base_fcff, g_init=INITIAL_GROWTH
            )
            grid[w][g] = res["enterprise_value"]

    return grid, worst_case_spread


if __name__ == "__main__":
    base_fcff = compute_fcff_base()
    dcf_base = calculate_dcf_ev(
        wacc=BASE_WACC,
        g_term=TERMINAL_GROWTH,
        base_fcff=base_fcff,
        g_init=INITIAL_GROWTH,
    )
    sens_grid, spread_check = generate_sensitivity_grid(
        BASE_WACC, TERMINAL_GROWTH, base_fcff
    )

    ebitda = EBIT + DEP_AMORT
    ev_ebitda_val = ebitda * EV_EBITDA_MULTIPLE
    premium_pct = ((ev_ebitda_val - dcf_base["enterprise_value"]) / dcf_base["enterprise_value"]) * 100

    print("=========================================================================================")
    print("                    DISCOUNTED CASH FLOW (DCF) VALUATION ENGINE                         ")
    print("=========================================================================================")
    print(f"Base FCFF (Year 0)        : ₹{base_fcff:.2f} Cr")
    print(f"Cost of Equity (PAYFIN β) : {COST_OF_EQUITY*100:.2f}% (R_f=7.0%, E(R_m)=13.0%, β=1.35)")
    print(f"After-Tax Cost of Debt    : {AFTER_TAX_COST_DEBT*100:.2f}% (Pre-Tax=8.0%, Tax=25.0%)")
    print(f"Computed WACC (70/30 D/E) : {BASE_WACC*100:.2f}%")
    print(f"Terminal Growth Rate (g)  : {TERMINAL_GROWTH*100:.2f}%")
    print("-" * 89)
    print("Explicit 5-Year FCFF Projections (INR Cr):")
    for yr, fcf in enumerate(dcf_base["projections"], 1):
        print(f"  Year {yr}: ₹{fcf:.2f} Cr")
    print(f"Sum of PV Explicit FCFFs  : ₹{dcf_base['pv_fcff']:.2f} Cr")
    print(f"PV of Terminal Value (TV) : ₹{dcf_base['pv_terminal_value']:.2f} Cr")
    print(f"BASE ENTERPRISE VALUE (EV): ₹{dcf_base['enterprise_value']:.2f} Cr")
    print("-" * 89)

    print("\n3x3 ENTERPRISE VALUE SENSITIVITY GRID (INR Crores):")
    label = "WACC / g"
    header_g = [f"g = {g*100:.1f}%" for g in [TERMINAL_GROWTH-0.01, TERMINAL_GROWTH, TERMINAL_GROWTH+0.01]]
    print(f"{label:<15} | {header_g[0]:<15} | {header_g[1]:<15} | {header_g[2]:<15}")
    print("-" * 68)

    for w, g_map in sens_grid.items():
        row_str = f"WACC = {w*100:.2f}%  |"
        for g, ev_val in g_map.items():
            row_str += f" ₹{ev_val:<13.2f} |"
        print(row_str)

    print("-" * 68)
    print(f"Required Self-Check Verification: Worst-Case Spread (Min WACC - Max g) = {spread_check*100:.2f}% (>= 1.00% Pass)")
    print("-" * 89)

    print("\n=========================================================================================")
    print("                    EV/EBITDA MULTIPLE CROSS-CHECK & COMPARISON                         ")
    print("=========================================================================================")
    print(f"Base EBITDA (EBIT + D&A)  : ₹{ebitda:.2f} Cr")
    print(f"Selected Peer Multiple    : {EV_EBITDA_MULTIPLE:.1f}x")
    print(f"Multiple-Based EV         : ₹{ev_ebitda_val:.2f} Cr")
    print(f"DCF Base Case EV          : ₹{dcf_base['enterprise_value']:.2f} Cr")
    print("-" * 89)
    print(
        "Comparison Comment:\n"
        f"The DCF base-case valuation of ₹{dcf_base['enterprise_value']:.2f} Cr provides an intrinsic baseline, "
        f"while the 12.0x EV/EBITDA peer multiple yields ₹{ev_ebitda_val:.2f} Cr (a {premium_pct:.1f}% market premium). "
        "This variance reflects public market multiples pricing in immediate scale and broader market momentum, "
        "whereas our DCF conservatively discounts future cash flows at a 12.37% WACC with a fading growth profile."
    )
    print("=========================================================================================")