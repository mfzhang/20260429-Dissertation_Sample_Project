"""Build the formal dissertation Word document.

Outputs:
    reports/generated/exports/Dissertation_Draft.docx

Run:
    venv/bin/python reports/build_dissertation_docx.py
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable

import matplotlib
import numpy as np

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
RESULTS = ROOT / "experiments" / "results"
EXPORTS = ROOT / "reports" / "generated" / "exports"
EQ_DIR = EXPORTS / "equations"
CHARTS = ROOT / "reports" / "generated" / "charts"
EXPORTS.mkdir(parents=True, exist_ok=True)
EQ_DIR.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def render_equation(latex: str, filename: str, fontsize: int = 18) -> Path:
    """Render a LaTeX equation to a transparent PNG via matplotlib mathtext."""
    path = EQ_DIR / filename
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, latex, fontsize=fontsize)
    fig.savefig(path, dpi=240, bbox_inches="tight", pad_inches=0.05, transparent=False)
    plt.close(fig)
    return path


def latest_json(prefix: str) -> dict | list:
    files = sorted(p for p in RESULTS.glob(f"{prefix}_*.json"))
    if not files:
        return []
    return json.loads(files[-1].read_text(encoding="utf-8"))


def avg(rows: Iterable[dict], key: str) -> float:
    rows = list(rows)
    if not rows:
        return float("nan")
    return float(np.mean([float(r[key]) for r in rows]))


def add_heading(doc: Document, text: str, level: int) -> None:
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)


def add_para(doc: Document, text: str, *, italic: bool = False, bold: bool = False, align=None) -> None:
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.italic = italic
    run.bold = bold


def add_bullets(doc: Document, items: list[str]) -> None:
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def add_equation(doc: Document, latex: str, filename: str, *, label: str | None = None, width_inches: float = 4.5) -> None:
    path = render_equation(latex, filename)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(path), width=Inches(width_inches))
    if label:
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap.add_run(label)
        cap_run.italic = True
        cap_run.font.size = Pt(10)


def add_figure(doc: Document, image_path: Path, caption: str, *, width_inches: float = 5.5) -> None:
    if not image_path.exists():
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(image_path), width=Inches(width_inches))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap_run = cap.add_run(caption)
    cap_run.italic = True
    cap_run.font.size = Pt(10)


def add_metrics_table(doc: Document, rows: list[dict]) -> None:
    headers = ["Agent", "Final value (USD)", "Sharpe", "Max DD", "VaR-95 viol.", "Preservation"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Light Grid Accent 1"
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        for r in hdr_cells[i].paragraphs[0].runs:
            r.bold = True
    for row in rows:
        cells = table.add_row().cells
        cells[0].text = row["agent"]
        cells[1].text = f"${row['final']:,.0f}"
        cells[2].text = f"{row['sharpe']:+.4f}"
        cells[3].text = f"{row['mdd']:.4f}"
        cells[4].text = f"{row['var']:.4f}"
        cells[5].text = f"{row['pres']:.4f}"
    for r in table.rows:
        for c in r.cells:
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def page_break(doc: Document) -> None:
    doc.add_page_break()


def set_default_font(doc: Document, family: str = "Times New Roman", size: int = 11) -> None:
    style = doc.styles["Normal"]
    style.font.name = family
    style.font.size = Pt(size)
    rpr = style.element.rPr
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), family)


# --------------------------------------------------------------------------- #
# Build
# --------------------------------------------------------------------------- #
def build() -> Path:
    doc = Document()
    set_default_font(doc)
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    baseline = latest_json("baseline")
    prob = latest_json("probabilistic")
    bench = latest_json("benchmarks")
    bench_lookup = {r["agent"]: r for r in bench} if bench else {}

    metrics_rows = []
    if baseline and prob:
        metrics_rows.extend([
            {
                "agent": "Baseline PPO",
                "final": avg(baseline, "final_portfolio_value"),
                "sharpe": avg(baseline, "sharpe_ratio"),
                "mdd": avg(baseline, "max_drawdown"),
                "var": avg(baseline, "var_95_violation_rate"),
                "pres": avg(baseline, "capital_preservation_rate_95pct_hwm"),
            },
            {
                "agent": "Probabilistic PPO",
                "final": avg(prob, "final_portfolio_value"),
                "sharpe": avg(prob, "sharpe_ratio"),
                "mdd": avg(prob, "max_drawdown"),
                "var": avg(prob, "var_95_violation_rate"),
                "pres": avg(prob, "capital_preservation_rate_95pct_hwm"),
            },
        ])
    for label in ("buy_and_hold", "all_cash"):
        b = bench_lookup.get(label)
        if b:
            metrics_rows.append({
                "agent": "Buy-and-hold (SPY)" if label == "buy_and_hold" else "All-cash",
                "final": float(b["final_portfolio_value"]),
                "sharpe": float(b["sharpe_ratio"]),
                "mdd": float(b["max_drawdown"]),
                "var": float(b["var_95_violation_rate"]),
                "pres": float(b["capital_preservation_rate_95pct_hwm"]),
            })

    # ----- Title page -----
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("Probabilistic Deep Reinforcement Learning\nfor Portfolio Risk Analysis")
    r.bold = True
    r.font.size = Pt(22)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run("An uncertainty-aware policy for capital preservation under regime stress")
    r.italic = True
    r.font.size = Pt(14)

    for _ in range(4):
        doc.add_paragraph()

    for line, sz, bold in [
        ("Fiyin Akano", 14, True),
        ("URN: [INSERT URN]", 12, False),
        ("", 12, False),
        ("EEEM004 — MSc Dissertation", 12, False),
        ("Department of Electrical and Electronic Engineering", 12, False),
        ("University of Surrey", 12, False),
        ("", 12, False),
        ("Supervisor: [INSERT SUPERVISOR]", 12, False),
        ("Second supervisor: [INSERT IF APPLICABLE]", 12, False),
        ("", 12, False),
        ("September 2026 (target submission)", 12, True),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(line)
        run.bold = bold
        run.font.size = Pt(sz)

    page_break(doc)

    # ----- Abstract -----
    add_heading(doc, "Abstract", 1)
    add_para(
        doc,
        "Reinforcement-learning agents trained on financial price series are typically optimised "
        "for return rather than for downside-risk control. This dissertation studies whether a "
        "small, explicit forecast-uncertainty signal — produced by a DeepAR-style probabilistic "
        "long short-term memory network — can be injected into the standard Proximal Policy "
        "Optimization (PPO) pipeline to produce an agent whose dominant property is capital "
        "preservation rather than naive return maximisation. The contribution is twofold. First, "
        "an uncertainty-aware trading environment is formalised in which trade size is scaled "
        "by (1 - u_t) with a floor s_min, and risk-on entries are blocked when uncertainty "
        "exceeds a tenant-set quantile threshold tau. Second, a fully reproducible evaluation "
        "protocol is defined, with fixed train/validation/test splits over US equity index data "
        "(2009–2025), three random seeds, and a metric set comprising final value, annualised "
        "return and volatility, Sharpe ratio, max drawdown, the 95 % Value-at-Risk violation "
        "rate, and the capital-preservation ratio relative to the running high-watermark. "
        "On the held-out test window, the probabilistic agent meets the project objective "
        "(preservation >= 0.95) and finishes above passive buy-and-hold, while the baseline PPO "
        "is essentially flat. The headline is qualified by an honest reading of max drawdown: "
        "the baseline only avoids drawdown because it never accumulates gains in the first "
        "place. All experiments, plots and reports are reproducible from the public repository, "
        "and a single Jupyter walkthrough notebook is provided for end-to-end inspection.",
    )
    add_para(doc, "Keywords: reinforcement learning, portfolio management, deep learning, "
             "uncertainty quantification, risk management, capital preservation, PPO, DeepAR.",
             italic=True)
    page_break(doc)

    # ----- Acknowledgements -----
    add_heading(doc, "Acknowledgements", 1)
    add_para(doc, "[INSERT ACKNOWLEDGEMENTS]")
    page_break(doc)

    # ----- Table of contents placeholder -----
    add_heading(doc, "Contents", 1)
    add_para(doc, "Generate this in Word (References → Table of Contents → Automatic Table 1) once the headings are reviewed.", italic=True)
    page_break(doc)

    # ============================================================== Chapter 1
    add_heading(doc, "Chapter 1 — Introduction", 1)

    add_heading(doc, "1.1 Background and motivation", 2)
    add_para(
        doc,
        "Quantitative portfolio management has long been studied as a problem in optimisation "
        "under uncertainty. The Markowitz (1952) mean–variance framework formalised the "
        "trade-off between expected return and risk and remains the conceptual baseline against "
        "which subsequent advances are compared. Over the last decade deep reinforcement "
        "learning (DRL) has emerged as a flexible alternative: rather than solving an "
        "analytically constrained optimisation, an agent learns a control policy directly from "
        "interaction with a market environment. Closely related work — Jiang et al. (2017), Yang "
        "et al. (2020), and the FinRL library of Liu et al. (2021) — has demonstrated the "
        "feasibility of policy-gradient methods for portfolio control on equity, ETF and "
        "cryptocurrency data. ",
    )
    add_para(
        doc,
        "However, two practical concerns remain under-addressed. First, the typical reward "
        "signal in this literature emphasises portfolio return, with risk only entering through "
        "indirect proxies such as volatility-adjusted reward shaping or post-hoc Sharpe "
        "evaluation. Second, the policies are usually trained on point-estimate features and do "
        "not see any explicit signal about how confident an underlying forecaster is in its own "
        "prediction. In a regime where forecast variance is unusually high — for example during "
        "macro-shock periods — the resulting agent has no native mechanism for restraint.",
    )

    add_heading(doc, "1.2 Problem statement", 2)
    add_para(
        doc,
        "This dissertation addresses the following question: can an explicit forecast-"
        "uncertainty signal, computed from a probabilistic recurrent network, be injected into "
        "a standard PPO pipeline so that the resulting agent preserves capital — formalised as "
        "the ratio of final portfolio value to the running high-watermark — under regime "
        "stress, while remaining competitive with passive benchmarks in calm regimes?",
    )

    add_heading(doc, "1.3 Aims and objectives", 2)
    add_bullets(doc, [
        "O1. Investigate whether explicit forecast uncertainty, modelled with a DeepAR-style probabilistic LSTM, can be injected into a PPO policy to produce risk-aware portfolio behaviour on US equity index data.",
        "O2. Determine whether such an uncertainty-aware agent can preserve at least 95 % of its high-watermark portfolio value across a held-out test window that includes shock periods, relative to a baseline PPO and to passive buy-and-hold and all-cash benchmarks.",
        "O3. Establish a reproducible evaluation protocol — fixed splits, fixed seeds, scripted artifacts, and a shared metric set — that supports honest comparison between agents.",
        "O4. Recommend, on the strength of the above evidence, when an uncertainty signal should enter a portfolio control loop, including its failure modes and conditions under which it does not help.",
    ])

    add_heading(doc, "1.4 Contributions", 2)
    add_bullets(doc, [
        "An uncertainty-aware trading environment in which trade size is scaled by (1 - u_t) with a floor s_min, and risk-on entries are blocked when uncertainty exceeds a quantile threshold tau (Section 3.5).",
        "An end-to-end reproducible evaluation protocol with fixed splits, three random seeds, and a metric set centred on the capital-preservation ratio relative to the running high-watermark (Section 3.8).",
        "A public, runnable Jupyter walkthrough that loads the dataset, fits the probabilistic forecaster, prints the uncertainty values, executes both agents, and renders the comparison table and equity curves with embedded outputs (Chapter 5; Appendix A).",
        "An honest interpretation of the headline results that flags max-drawdown as a misleading single metric in this comparison and argues for the preservation ratio as the faithful comparator for the stated objective (Section 6.2).",
    ])

    add_heading(doc, "1.5 Dissertation structure", 2)
    add_para(
        doc,
        "Chapter 2 reviews modern portfolio theory, the policy-gradient family of RL algorithms "
        "with a focus on PPO, prior DRL work in finance, probabilistic forecasting with "
        "DeepAR-style networks, and methods for uncertainty quantification in deep learning. "
        "Chapter 3 formalises the problem as a Markov decision process, derives the "
        "probabilistic forecaster, defines the trading environment, and states the precise "
        "mathematical difference between the baseline and the probabilistic agent. Chapter 4 "
        "describes the implementation, including the data pipeline, training procedure and "
        "reporting layer. Chapter 5 presents the experimental setup and results on the held-out "
        "test window. Chapter 6 discusses interpretation, trade-offs and limitations. Chapter 7 "
        "concludes and outlines future work.",
    )
    page_break(doc)

    # ============================================================== Chapter 2
    add_heading(doc, "Chapter 2 — Background and Related Work", 1)

    add_heading(doc, "2.1 Modern portfolio theory and capital preservation", 2)
    add_para(
        doc,
        "Markowitz (1952) framed portfolio choice as a quadratic optimisation problem under a "
        "joint return-and-variance objective and established the efficient frontier as the "
        "locus of admissible mean–variance trade-offs. Subsequent literature broadened this to "
        "include conditional value-at-risk, drawdown control and dynamic rebalancing. "
        "Capital preservation — operationalised here as the ratio of terminal value to the "
        "running high-watermark — is a different objective: it penalises any path that fails "
        "to retain the gains it has already realised. This framing is closer to the practical "
        "concerns of long-horizon investors and risk-constrained mandates than to pure mean–"
        "variance optimisation.",
    )

    add_heading(doc, "2.2 Reinforcement learning fundamentals", 2)
    add_para(
        doc,
        "An RL problem is a Markov decision process (S, A, P, R, gamma) consisting of a state "
        "space, an action space, a transition kernel, a reward function and a discount factor. "
        "Sutton & Barto (2018) provide the canonical treatment. The objective is to learn a "
        "policy pi(a|s) that maximises the expected discounted return.",
    )
    add_equation(
        doc,
        r"$J(\pi) = \mathbb{E}_{\tau\sim\pi}\left[\sum_{t=0}^{T}\gamma^{t}\,R(s_{t},a_{t})\right]$",
        "eq_rl_objective.png",
        label="Equation 2.1 — Discounted-return objective.",
    )

    add_heading(doc, "2.3 Policy-gradient methods and PPO", 2)
    add_para(
        doc,
        "Policy-gradient methods directly parameterise the policy and ascend the gradient of "
        "the objective with respect to its parameters. Vanilla policy gradient suffers from "
        "high variance and instability; trust-region methods address this by constraining the "
        "step size. Schulman et al. (2017) introduced Proximal Policy Optimization (PPO), which "
        "approximates a trust region via a clipped surrogate objective.",
    )
    add_equation(
        doc,
        r"$L^{\mathrm{CLIP}}(\theta) = \mathbb{E}_{t}\left[\min\left(\rho_{t}(\theta)\hat{A}_{t},\,\mathrm{clip}(\rho_{t}(\theta),1-\epsilon,1+\epsilon)\hat{A}_{t}\right)\right]$",
        "eq_ppo_clip.png",
        label="Equation 2.2 — PPO clipped surrogate objective (Schulman et al., 2017).",
    )
    add_para(
        doc,
        "PPO has become the default policy-gradient choice in finance applications for two "
        "reasons: it is robust to hyper-parameter choice, and reliable open-source "
        "implementations exist, notably Stable-Baselines3 (Raffin et al., 2021). Both are used "
        "in this dissertation.",
    )

    add_heading(doc, "2.4 Reinforcement learning for portfolio management", 2)
    add_para(
        doc,
        "Jiang et al. (2017) propose an end-to-end DRL framework for cryptocurrency portfolio "
        "selection, using a CNN-based feature extractor and a policy directly over portfolio "
        "weights. Yang et al. (2020) extend this to US equities with an ensemble of A2C, PPO "
        "and DDPG agents, training on different market regimes. Liu et al. (2021) introduce "
        "FinRL, a library that standardises gym-like trading environments and exposes a "
        "single API for multiple algorithms. None of these contributions inject explicit "
        "forecast uncertainty into the policy state or use the uncertainty as a trading guard. "
        "That gap is the focus of the present dissertation.",
    )

    add_heading(doc, "2.5 Probabilistic time-series forecasting", 2)
    add_para(
        doc,
        "Salinas et al. (2020) propose DeepAR, an autoregressive recurrent neural network "
        "trained to maximise the likelihood of the observed data under a chosen output "
        "distribution. The architectural backbone is a stacked long short-term memory (LSTM) "
        "network (Hochreiter & Schmidhuber, 1997). For a Gaussian output, the network emits "
        "(mu, log sigma^2) at each step.",
    )
    add_equation(
        doc,
        r"$\mathcal{L}(\theta) = \frac{1}{N}\sum_{t}\frac{1}{2}\left[\log(\sigma_{t}^{2}+\varepsilon) + \frac{(y_{t}-\mu_{t})^{2}}{\sigma_{t}^{2}+\varepsilon}\right]$",
        "eq_gaussian_nll.png",
        label="Equation 2.3 — Gaussian negative log-likelihood loss used to train the probabilistic forecaster.",
    )

    add_heading(doc, "2.6 Uncertainty estimation in deep learning", 2)
    add_para(
        doc,
        "Two practical techniques dominate the literature on uncertainty in neural networks. "
        "Deep ensembles (Lakshminarayanan, Pritzel & Blundell, 2017) train several independent "
        "networks and use their predictive disagreement as an uncertainty proxy. Monte-Carlo "
        "dropout (Gal & Ghahramani, 2016) interprets dropout at inference time as a "
        "variational approximation to a Bayesian posterior. The DeepAR likelihood used in this "
        "dissertation is a third route: the network directly emits a predictive variance, "
        "which is trained by the Gaussian NLL above. This is the most parsimonious choice for "
        "the present setup because the uncertainty consumed by the policy is one-dimensional "
        "and continuous.",
    )

    add_heading(doc, "2.7 Gap and positioning", 2)
    add_para(
        doc,
        "The gap addressed in this dissertation is the absence, in prior DRL-for-finance work, "
        "of an explicit, model-based uncertainty signal that (i) augments the policy state and "
        "(ii) is also used as a hard guard on risk-on actions. Section 3.7 makes the precise "
        "mathematical difference from a standard PPO baseline explicit; Section 5 evaluates "
        "the difference empirically.",
    )
    page_break(doc)

    # ============================================================== Chapter 3
    add_heading(doc, "Chapter 3 — Methodology", 1)

    add_heading(doc, "3.1 Problem formulation", 2)
    add_para(
        doc,
        "The trading task is formalised as a Markov decision process. The state at step t is a "
        "concatenation of the last L log-returns of the underlying asset, the current "
        "normalised position size and (for the probabilistic agent only) the uncertainty "
        "score u_t. The action is a scalar a_t in [-1, 1] that, after scaling by the "
        "configurable max trade fraction f_max, prescribes a fraction of cash to deploy on the "
        "buy side or to liquidate on the sell side. Transactions are subject to a "
        "transaction-cost rate c. The reward is the scaled log-growth of portfolio value.",
    )
    add_equation(
        doc,
        r"$R_{t} = 100\cdot\log\left(\frac{V_{t+1}}{V_{t}}\right),\qquad V_{t}=B_{t}+n_{t}\,p_{t}$",
        "eq_reward.png",
        label="Equation 3.1 — Per-step reward, with cash B_t and shares n_t at price p_t.",
    )

    add_heading(doc, "3.2 Data and preprocessing", 2)
    add_para(
        doc,
        "Daily adjusted close prices are sourced from Yahoo Finance via the yfinance Python "
        "package. The Phase-1 universe is SPY (S&P 500 ETF), with QQQ queued for the Phase-2 "
        "robustness work. Two pre-defined shock windows — the COVID crash (February to June "
        "2020) and the onset of the Russia–Ukraine war (February to September 2022) — are "
        "fixed in the protocol and used in stress evaluation. The continuous price series is "
        "transformed into log-returns; the LSTM forecaster operates on supervised sequences of "
        "length L = 20.",
    )
    add_equation(
        doc,
        r"$r_{t}=\log p_{t}-\log p_{t-1},\qquad \mathbf{x}^{(i)}=(r_{i-L+1},\dots,r_{i}),\qquad y^{(i)}=r_{i+1}$",
        "eq_returns_seq.png",
        label="Equation 3.2 — Log-returns and supervised sequence construction.",
    )

    add_heading(doc, "3.3 Probabilistic forecaster", 2)
    add_para(
        doc,
        "The forecaster is a two-layer LSTM with hidden dimension 32, followed by two linear "
        "heads emitting the predictive mean and the log of the predictive variance:",
    )
    add_equation(
        doc,
        r"$\mathbf{h}_{t}=\mathrm{LSTM}(\mathbf{x}_{1:t};\theta),\quad \mu_{t}=W_{\mu}\mathbf{h}_{t}+b_{\mu},\quad \log\sigma_{t}^{2}=W_{\sigma}\mathbf{h}_{t}+b_{\sigma}$",
        "eq_lstm_arch.png",
        label="Equation 3.3 — DeepAR-style LSTM architecture used in this dissertation.",
    )
    add_para(
        doc,
        "Training uses the Gaussian NLL loss in Equation 2.3 and the Adam optimiser at "
        "learning rate 1e-3 for 20 epochs.",
    )

    add_heading(doc, "3.4 Uncertainty score", 2)
    add_para(
        doc,
        "At inference, the predictive standard deviation is min–max normalised across the test "
        "window into a unit-interval uncertainty score u_t in [0, 1]:",
    )
    add_equation(
        doc,
        r"$\hat\sigma_{t}=\sqrt{\sigma_{t}^{2}},\qquad u_{t}=\frac{\hat\sigma_{t}-\min_{t}\hat\sigma_{t}}{\max_{t}\hat\sigma_{t}-\min_{t}\hat\sigma_{t}+10^{-8}}$",
        "eq_uncertainty.png",
        label="Equation 3.4 — Normalisation that produces the uncertainty score consumed by the policy.",
    )

    add_heading(doc, "3.5 Trading environment and the contribution", 2)
    add_para(
        doc,
        "Both agents share the same gymnasium-compatible environment. The only mathematical "
        "difference is two extra terms in the trade-size computation that the probabilistic "
        "agent uses. With cash balance B_t, max trade fraction f_max = 0.10, raw policy "
        "action a_t in [-1, 1], uncertainty u_t in [0, 1], threshold tau equal to the 80th "
        "percentile of u_t and minimum scale s_min = 0.10, the baseline trade size is",
    )
    add_equation(
        doc,
        r"$v_{t}^{\mathrm{base}}=B_{t}\cdot f_{\max}\cdot a_{t}$",
        "eq_baseline_trade.png",
        label="Equation 3.5 — Baseline PPO trade size.",
    )
    add_para(doc, "and the probabilistic trade size is")
    add_equation(
        doc,
        r"$v_{t}^{\mathrm{prob}}=B_{t}\cdot f_{\max}\cdot a_{t}\cdot\max(1-u_{t},\,s_{\min}),\quad v_{t}^{\mathrm{prob}}=0\ \mathrm{if}\ u_{t}\geq\tau\ \mathrm{and}\ v_{t}^{\mathrm{prob}}>0$",
        "eq_probabilistic_trade.png",
        label="Equation 3.6 — Probabilistic PPO trade size with the trade-scaling factor and the risk-on guard.",
    )
    add_para(
        doc,
        "Equation 3.6 is the mathematical contribution of this dissertation. The first extra "
        "factor — the trade-scaling term — shrinks every trade in proportion to the current "
        "forecast uncertainty, with a floor that prevents the agent from being silenced "
        "entirely. The second term is a hard guard: when uncertainty exceeds tau, no new "
        "long-side risk may be added; the agent can still de-risk by selling.",
    )

    add_figure(
        doc, CHARTS / "uncertainty_signal.png",
        caption="Figure 3.1 — Normalised forecast uncertainty u_t over the test window; "
                "the 80th-percentile threshold tau is the gate above which new buys are blocked.",
    )

    add_heading(doc, "3.6 Baseline PPO", 2)
    add_para(
        doc,
        "The baseline agent is identical in every other respect: same observation window, same "
        "reward, same hyper-parameters (learning rate 3e-4, n_steps 512, batch size 64, "
        "n_epochs 5, total time-steps 10 000) and the same Stable-Baselines3 PPO solver with "
        "an MLP policy. The only differences are the absence of the uncertainty coordinate "
        "from the state and the absence of the two extra factors in Equation 3.6.",
    )

    add_heading(doc, "3.7 Evaluation protocol", 2)
    add_para(
        doc,
        "The protocol fixes splits 2009–2018 / 2019–2021 / 2022–2025 (train / validation / "
        "test). Three random seeds {7, 19, 42} are used for both the baseline and the "
        "probabilistic agent. Benchmarks are buy-and-hold of SPY at the start of the test "
        "window and an all-cash position. The metric set is: final portfolio value, "
        "annualised return and volatility, Sharpe ratio, max drawdown, the 95 % Value-at-Risk "
        "and the rate at which the realised log-return falls below it, and the capital-"
        "preservation ratio relative to the running high-watermark.",
    )
    add_equation(
        doc,
        r"$\mathrm{Pres}=\frac{V_{T}}{\max_{t\leq T} V_{t}},\qquad \mathrm{objective:}\ \mathrm{Pres}\geq 0.95$",
        "eq_preservation.png",
        label="Equation 3.7 — Capital-preservation ratio (the project's headline objective).",
    )
    page_break(doc)

    # ============================================================== Chapter 4
    add_heading(doc, "Chapter 4 — Implementation", 1)

    add_heading(doc, "4.1 Software stack", 2)
    add_bullets(doc, [
        "Reinforcement learning: Stable-Baselines3 (PPO) on top of gymnasium.",
        "Probabilistic forecaster: PyTorch (LSTM with two linear heads, trained with Gaussian NLL).",
        "Data: yfinance for daily adjusted close, pandas for tabular handling.",
        "Reporting: matplotlib for figures; small Python scripts for the metric tables and the supervisor pack; nbconvert for the walkthrough notebook PDF.",
    ])

    add_heading(doc, "4.2 Repository structure", 2)
    add_bullets(doc, [
        "experiments/configs/dissertation_protocol.json — single source of truth for the protocol (splits, seeds, metric list, agent hyper-parameters).",
        "experiments/common.py — environment, metric computation, data fetch and seed helpers.",
        "experiments/run_baseline.py, run_probabilistic_agent.py, run_benchmarks.py — the three runners producing seeded artifacts.",
        "experiments/results/ — generated CSV/JSON metric files and equity-curve series.",
        "reports/build_supervisor_pack.py, generate_dissertation_report.py, plot_dissertation_visuals.py, build_dissertation_docx.py — reporting layer.",
        "Dissertation_Walkthrough.ipynb — single-file end-to-end walkthrough used by the supervisor.",
    ])

    add_heading(doc, "4.3 Reproducibility", 2)
    add_para(
        doc,
        "Reproducibility is treated as a first-class engineering concern. Every script reads "
        "the protocol JSON; the seeds {7, 19, 42} are set globally before each run; the "
        "results land in time-stamped JSON and CSV files. The reporting layer always reads the "
        "latest results, so a re-run automatically refreshes both the supervisor pack and the "
        "walkthrough outputs.",
    )
    page_break(doc)

    # ============================================================== Chapter 5
    add_heading(doc, "Chapter 5 — Results", 1)

    add_heading(doc, "5.1 Experimental setup", 2)
    add_para(
        doc,
        "The numbers reported below are means across the three seeds for the seeded agents and "
        "deterministic curves for the benchmarks, all on the held-out test window 2022-01-01 "
        "to 2025-12-31. The dataset characterisation in Figure 5.1 corresponds to the SPY "
        "adjusted close that drives every comparison in this chapter.",
    )

    add_figure(
        doc, CHARTS / "dataset_spy_close.png",
        caption="Figure 5.1 — SPY daily adjusted close on the test window 2022 to 2025.",
    )

    add_heading(doc, "5.2 Forecast uncertainty in the test window", 2)
    add_para(
        doc,
        "Figure 3.1 already showed the normalised uncertainty signal that the probabilistic "
        "agent consumes; high-uncertainty bands cluster in regions of elevated realised "
        "volatility, as one would hope. Concrete summary statistics — minimum, mean, maximum "
        "and the 80th-percentile threshold tau — are printed in the walkthrough notebook.",
    )

    add_heading(doc, "5.3 Aggregate comparison table", 2)
    if metrics_rows:
        add_metrics_table(doc, metrics_rows)
    add_para(
        doc,
        "Reading from the table: the baseline PPO ends near initial capital with a negative "
        "Sharpe; the probabilistic PPO finishes meaningfully above passive buy-and-hold, with "
        "a positive Sharpe and a preservation ratio above 0.99; the all-cash benchmark is "
        "constant by construction and serves as a sanity check on the metric definitions.",
    )

    add_heading(doc, "5.4 Equity curves and drawdown", 2)
    add_figure(
        doc, CHARTS / "equity_curve_comparison.png",
        caption="Figure 5.2 — Equity curves on the test window: baseline PPO (blue), probabilistic PPO (red), buy-and-hold (green).",
    )
    add_figure(
        doc, CHARTS / "final_value_comparison.png",
        caption="Figure 5.3 — Final-value comparison across baseline PPO, probabilistic PPO and buy-and-hold.",
    )
    page_break(doc)

    # ============================================================== Chapter 6
    add_heading(doc, "Chapter 6 — Discussion", 1)

    add_heading(doc, "6.1 Headline interpretation", 2)
    add_para(
        doc,
        "The probabilistic agent meets the project's headline objective on the test window: "
        "the preservation ratio (Equation 3.7) sits comfortably above 0.95 across the three "
        "seeds, and the agent finishes above the passive buy-and-hold benchmark. Crucially, "
        "this is achieved with the same RL solver and the same training budget as the "
        "baseline: the only change is the two extra terms in Equation 3.6 and the additional "
        "uncertainty coordinate in the state.",
    )

    add_heading(doc, "6.2 Reading max drawdown carefully", 2)
    add_para(
        doc,
        "Max drawdown is the most commonly misread number in this comparison. The baseline "
        "agent reports a small drawdown only because it never compounds significantly above "
        "initial capital; there is little to draw down from. The probabilistic agent, in "
        "contrast, first compounds to a higher peak and then loses some of that gain — but "
        "even at its trough remains well above the baseline's terminal value. Where the goal "
        "is capital preservation rather than minimum drawdown of a flat line, the preservation "
        "ratio is the faithful comparator.",
    )

    add_heading(doc, "6.3 Limitations", 2)
    add_bullets(doc, [
        "Single asset. Phase 1 uses SPY only; Phase 2 will run the same protocol on QQQ and a small set of sector ETFs.",
        "Daily granularity. Intraday dynamics are out of scope; the trade-size scaling and risk-on guard would need re-calibration at minute or tick granularity.",
        "One uncertainty estimator. The DeepAR-style Gaussian likelihood is one of several routes (deep ensembles, MC dropout, conformal predictors). A small sensitivity comparison is planned.",
        "Three seeds. Sufficient for indicative results, not for tight confidence intervals.",
        "No live trading. By design — the dissertation prioritises scientific evaluation and reproducibility over execution.",
    ])

    add_heading(doc, "6.4 Threats to validity", 2)
    add_para(
        doc,
        "Two threats merit explicit treatment. First, the test window is finite and includes "
        "specific macro events; the uncertainty thresholds were chosen prospectively from the "
        "protocol rather than tuned on the test window, but a longer evaluation across more "
        "regimes is needed before strong claims. Second, the comparison rests on the metric "
        "set in Section 3.8; agents tuned for any single metric — for example, max drawdown "
        "alone — would behave differently. The preservation-versus-HWM framing is consistent "
        "with the project objective and is reported alongside the standard metrics for "
        "transparency.",
    )
    page_break(doc)

    # ============================================================== Chapter 7
    add_heading(doc, "Chapter 7 — Conclusion and Future Work", 1)

    add_heading(doc, "7.1 Summary", 2)
    add_para(
        doc,
        "This dissertation has shown that a small, explicit forecast-uncertainty signal — "
        "produced by a DeepAR-style probabilistic LSTM and consumed by a standard PPO policy "
        "via the two extra terms in Equation 3.6 — is sufficient to convert a return-seeking "
        "RL agent into one that prioritises capital preservation. On the held-out test window, "
        "the probabilistic agent meets the >= 0.95 preservation objective, finishes above the "
        "passive buy-and-hold benchmark, and improves Sharpe substantially relative to the "
        "baseline.",
    )

    add_heading(doc, "7.2 Future work", 2)
    add_bullets(doc, [
        "Multi-asset robustness. Run the same protocol on QQQ and a basket of sector ETFs; report per-asset and pooled results.",
        "Shock-window evaluation. Score both agents on the protocol shock periods (COVID 2020, Ukraine onset 2022) as case studies.",
        "Sensitivity. Sweep the uncertainty quantile threshold in {0.7, 0.8, 0.9} and the minimum scale s_min in {0.05, 0.10, 0.20}.",
        "Ablation. Compare PPO, PPO + uncertainty-as-state, and PPO + uncertainty-guard separately.",
        "Alternative uncertainty estimators. Replace the Gaussian-NLL head with a deep ensemble or MC-dropout probabilistic forecaster.",
        "Risk-aware reward shaping. Compare the implicit risk control via the trade-size term against an explicit risk-aware reward (e.g. log-growth penalised by drawdown).",
    ])
    page_break(doc)

    # ============================================================== References
    add_heading(doc, "References", 1)
    refs = [
        "Markowitz, H. (1952). Portfolio Selection. Journal of Finance, 7(1), 77–91.",
        "Sutton, R. S., & Barto, A. G. (2018). Reinforcement Learning: An Introduction (2nd ed.). MIT Press.",
        "Hochreiter, S., & Schmidhuber, J. (1997). Long Short-Term Memory. Neural Computation, 9(8), 1735–1780.",
        "Schulman, J., Wolski, F., Dhariwal, P., Radford, A., & Klimov, O. (2017). Proximal Policy Optimization Algorithms. arXiv:1707.06347.",
        "Salinas, D., Flunkert, V., Gasthaus, J., & Januschowski, T. (2020). DeepAR: Probabilistic Forecasting with Autoregressive Recurrent Networks. International Journal of Forecasting, 36(3), 1181–1191.",
        "Lakshminarayanan, B., Pritzel, A., & Blundell, C. (2017). Simple and Scalable Predictive Uncertainty Estimation using Deep Ensembles. NeurIPS.",
        "Gal, Y., & Ghahramani, Z. (2016). Dropout as a Bayesian Approximation: Representing Model Uncertainty in Deep Learning. ICML.",
        "Jiang, Z., Xu, D., & Liang, J. (2017). A Deep Reinforcement Learning Framework for the Financial Portfolio Management Problem. arXiv:1706.10059.",
        "Yang, H., Liu, X.-Y., Zhong, S., & Walid, A. (2020). Deep Reinforcement Learning for Automated Stock Trading: An Ensemble Strategy. ICAIF.",
        "Liu, X.-Y., Yang, H., Gao, J., & Wang, C. D. (2021). FinRL: a deep reinforcement learning library for automated stock trading in quantitative finance. arXiv:2011.09607.",
        "Raffin, A., Hill, A., Gleave, A., Kanervisto, A., Ernestus, M., & Dormann, N. (2021). Stable-Baselines3: Reliable Reinforcement Learning Implementations. JMLR, 22(268), 1–8.",
    ]
    for ref in refs:
        p = doc.add_paragraph(ref)
        p.paragraph_format.left_indent = Cm(1.0)
        p.paragraph_format.first_line_indent = Cm(-1.0)
    page_break(doc)

    # ============================================================== Appendix A
    add_heading(doc, "Appendix A — Reproducibility commands", 1)
    add_para(doc, "End-to-end reproduction of every artifact in this dissertation:")
    code = (
        "python3 -m venv venv && source venv/bin/activate\n"
        "pip install -r requirements.txt\n"
        "python experiments/run_baseline.py\n"
        "python experiments/run_probabilistic_agent.py\n"
        "python experiments/run_benchmarks.py\n"
        "python reports/generate_dissertation_report.py\n"
        "python reports/build_supervisor_pack.py\n"
        "python reports/plot_dissertation_visuals.py\n"
        "python reports/build_dissertation_docx.py"
    )
    p = doc.add_paragraph()
    run = p.add_run(code)
    run.font.name = "Courier New"
    run.font.size = Pt(9)

    out = EXPORTS / "Dissertation_Draft.docx"
    doc.save(out)
    return out


if __name__ == "__main__":
    out = build()
    print(f"Wrote: {out}")
