import importlib.util
import math
import os
from pathlib import Path

# Load seed data dynamically from local directory
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
investor_mod = load_module_from_path(
    "investor_profiles", PART_C_DIR / "investor_profiles.py"
)

STOCK_UNIVERSE = stock_mod.STOCK_UNIVERSE
RISK_FREE_RATE = stock_mod.RISK_FREE_RATE  # 0.07
MARKET_RETURN = stock_mod.MARKET_RETURN  # 0.13
INVESTOR_PROFILES = investor_mod.INVESTOR_PROFILES

RHO = 0.30  # Stated pairwise correlation

PRESCRIBED_ALLOCATIONS = {
    "Conservative": ["PAYBOND", "PAYGOLD", "PAYRETAIL"],
    "Moderate": ["PAYRETAIL", "PAYINFRA", "PAYGOLD"],
    "Aggressive": ["PAYTECH", "PAYFIN", "PAYINFRA"],
}


def get_stock_data(ticker: str) -> dict:
    """Tool function to fetch stock attributes from STOCK_UNIVERSE."""
    if ticker in STOCK_UNIVERSE:
        return STOCK_UNIVERSE[ticker]
    raise ValueError(f"Ticker '{ticker}' not found in STOCK_UNIVERSE.")


def run_advisory_agent(profile: dict) -> dict:
    investor_id = profile["investor_id"]
    risk_tolerance = profile["risk_tolerance"]

    # 1. THINK: Look up exact prescribed 1/3-each allocation
    allocated_tickers = PRESCRIBED_ALLOCATIONS[risk_tolerance]
    weights = [1.0 / 3.0, 1.0 / 3.0, 1.0 / 3.0]

    # 2. ACT (Tool call): Fetch stock metrics
    stock_metrics = [get_stock_data(t) for t in allocated_tickers]

    # 3. OBSERVE -> DECIDE: Math Computations
    # CAPM Expected Return per stock: E(R_i) = R_f + beta_i * (E(R_m) - R_f)
    capm_returns = [
        RISK_FREE_RATE + m["beta"] * (MARKET_RETURN - RISK_FREE_RATE)
        for m in stock_metrics
    ]
    std_devs = [m["std_dev"] for m in stock_metrics]

    portfolio_capm_return = sum(w * r for w, r in zip(weights, capm_returns))

    # Portfolio Variance: Var(R_p) = Sum(w_i^2 * sigma_i^2) + 2 * Sum_{i<j}(w_i * w_j * Cov(R_i, R_j))
    var_direct = sum((w**2) * (sig**2) for w, sig in zip(weights, std_devs))
    var_cov = 0.0
    n = len(weights)
    for i in range(n):
        for j in range(i + 1, n):
            cov_ij = RHO * std_devs[i] * std_devs[j]
            var_cov += 2.0 * weights[i] * weights[j] * cov_ij

    portfolio_std_dev = math.sqrt(var_direct + var_cov)

    # Human-in-the-loop Escalation Check (> 20.0% Std Dev)
    escalated = portfolio_std_dev > 0.20
    status = "ESCALATED_TO_HUMAN_ADVISOR" if escalated else "FINALIZED"

    # 4. FINAL NARRATIVE SENTENCE (Gated by MOCK_LLM)
    mock_llm_flag = os.getenv("MOCK_LLM", "1")

    if mock_llm_flag == "1" or mock_llm_flag.lower() == "true":
        # Graded Baseline Mock Mode (f-string template)
        recommendation_text = (
            f"For {risk_tolerance} investor {investor_id}, we recommend an allocation across "
            f"{allocated_tickers} with an expected portfolio return of {portfolio_capm_return:.1%} "
            f"and volatility of {portfolio_std_dev:.1%}."
        )
    else:
        # Optional MOCK_LLM=0 Extension (Groq API Call)
        try:
            from groq import Groq

            client = Groq(api_key=os.getenv("GROQ_API_KEY"))
            prompt = (
                f"Phrase a professional portfolio advisory sentence for investor {investor_id} "
                f"({risk_tolerance} tier). Recommended tickers: {allocated_tickers}. "
                f"CAPM Expected Return: {portfolio_capm_return:.1%}, Volatility: {portfolio_std_dev:.1%}. "
                f"Keep it under 30 words."
            )
            completion = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
            )
            recommendation_text = completion.choices[0].message.content.strip()
        except Exception as e:
            recommendation_text = (
                f"For {risk_tolerance} investor {investor_id}, we recommend an allocation across "
                f"{allocated_tickers} with an expected portfolio return of {portfolio_capm_return:.1%} "
                f"and volatility of {portfolio_std_dev:.1%}."
            )

    return {
        "investor_id": investor_id,
        "risk_tolerance": risk_tolerance,
        "allocated_tickers": allocated_tickers,
        "portfolio_capm_return": portfolio_capm_return,
        "portfolio_std_dev": portfolio_std_dev,
        "status": status,
        "recommendation_text": recommendation_text,
    }


if __name__ == "__main__":
    for profile in INVESTOR_PROFILES:
        res = run_advisory_agent(profile)
        print(f"[{res['investor_id']}] Risk Tier: {res['risk_tolerance']}")
        print(f"Allocated Tickers: {res['allocated_tickers']}")
        print(f"CAPM Return E(R) : {res['portfolio_capm_return']*100:.2f}%")
        print(f"Portfolio Risk σ : {res['portfolio_std_dev']*100:.2f}%")
        print(f"Execution Status : {res['status']}")
        print(f"Narrative Output : {res['recommendation_text']}")
        print("-" * 80)