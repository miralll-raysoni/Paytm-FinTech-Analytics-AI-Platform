import importlib.util
import os
from pathlib import Path

# Load STOCK_UNIVERSE dynamically from local directory
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


def run_multi_agent_debate(ticker: str = "PAYTECH") -> dict:
    if ticker not in STOCK_UNIVERSE:
        raise ValueError(f"Ticker '{ticker}' not found in STOCK_UNIVERSE.")

    metrics = STOCK_UNIVERSE[ticker]
    beta = metrics["beta"]
    analyst_er = metrics["analyst_expected_return"]
    std_dev = metrics["std_dev"]

    mock_llm_flag = os.getenv("MOCK_LLM", "1")

    # =====================================================================
    # GRADED MOCK MODE BASELINE (MOCK_LLM=1)
    # =====================================================================
    if mock_llm_flag == "1" or mock_llm_flag.lower() == "true":
        # Stage 1: Bull Agent Argument
        bull_arg = (
            f"BULL AGENT: With an analyst expected return of {analyst_er:.1%} against a beta "
            f"of {beta:.2f}, {ticker} offers attractive upside potential and momentum to capture "
            f"strong market rallies."
        )

        # Stage 2: Bear Agent Argument
        bear_arg = (
            f"BEAR AGENT: However, {ticker} exhibits a high standard deviation of {std_dev:.1%}, "
            f"signaling significant volatility and downside exposure during broader market corrections."
        )

        # Stage 3: Synthesizer Agent (2-3 sentence balanced summary)
        synthesizer_summary = (
            f"SYNTHESIZER: While {ticker} provides strong return prospects ({analyst_er:.1%}) "
            f"amplified by a beta of {beta:.2f}, its elevated volatility ({std_dev:.1%}) requires strict risk controls. "
            f"We synthesize a HOLD/CAUTIOUS stance, recommending exposure only for aggressive risk profiles."
        )

    # =====================================================================
    # OPTIONAL MOCK_LLM=0 EXTENSION (Groq / External API)
    # =====================================================================
    else:
        try:
            from groq import Groq

            client = Groq(api_key=os.getenv("GROQ_API_KEY"))

            bull_prompt = (
                f"You are a bullish financial analyst. Make a concise 2-sentence case for ticker {ticker} "
                f"using analyst expected return {analyst_er:.1%} and beta {beta:.2f}."
            )
            bull_resp = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": bull_prompt}],
                temperature=0.3,
            )
            bull_arg = f"BULL AGENT: {bull_resp.choices[0].message.content.strip()}"

            bear_prompt = (
                f"You are a bearish risk analyst. Make a concise 2-sentence counter-argument against ticker {ticker} "
                f"focusing on its standard deviation of {std_dev:.1%}."
            )
            bear_resp = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": bear_prompt}],
                temperature=0.3,
            )
            bear_arg = f"BEAR AGENT: {bear_resp.choices[0].message.content.strip()}"

            synth_prompt = (
                f"Synthesize these two views into a balanced 2-sentence summary for {ticker}:\n"
                f"Bull View: {bull_arg}\n"
                f"Bear View: {bear_arg}\n"
                f"Reference beta={beta:.2f}, return={analyst_er:.1%}, and std_dev={std_dev:.1%}."
            )
            synth_resp = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": synth_prompt}],
                temperature=0.3,
            )
            synthesizer_summary = (
                f"SYNTHESIZER: {synth_resp.choices[0].message.content.strip()}"
            )

        except Exception:
            # Fallback to Mock if API fails
            bull_arg = (
                f"BULL AGENT: With an analyst expected return of {analyst_er:.1%} against a beta "
                f"of {beta:.2f}, {ticker} offers attractive upside potential."
            )
            bear_arg = (
                f"BEAR AGENT: However, {ticker} exhibits a high standard deviation of {std_dev:.1%}, "
                f"signaling significant volatility risk."
            )
            synthesizer_summary = (
                f"SYNTHESIZER: While {ticker} provides strong return prospects ({analyst_er:.1%}), "
                f"its elevated volatility ({std_dev:.1%}) requires strict risk controls. "
                f"We synthesize a HOLD/CAUTIOUS stance."
            )

    return {
        "ticker": ticker,
        "metrics": metrics,
        "bull_argument": bull_arg,
        "bear_argument": bear_arg,
        "synthesizer_summary": synthesizer_summary,
    }


if __name__ == "__main__":
    SELECTED_TICKER = "PAYTECH"
    result = run_multi_agent_debate(SELECTED_TICKER)

    print("=========================================================================================")
    print(f"                MULTI-AGENT DEBATE DEMO: TICKER [{SELECTED_TICKER}]                     ")
    print("=========================================================================================")
    print(f"Stock Metrics    : Beta = {result['metrics']['beta']} | Expected Return = {result['metrics']['analyst_expected_return']*100:.1f}% | Volatility = {result['metrics']['std_dev']*100:.1f}%\n")
    print(result["bull_argument"])
    print("-" * 89)
    print(result["bear_argument"])
    print("-" * 89)
    print(result["synthesizer_summary"])
    print("=========================================================================================")