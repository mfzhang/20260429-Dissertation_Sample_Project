# EEEM004 Interim Review — Draft Content

> Use this file to fill the **black** boxes in `InterimReview.docx`.
> The blue supervisor boxes are intentionally left empty.

---

## Cover information

- **Name:** Fiyin Akano
- **URN:** [FILL IN]
- **Supervisor:** [FILL IN]
- **Second supervisor (if applicable):** [FILL IN]
- **Date of meeting:** [FILL IN]

---

## Project title

**Probabilistic Deep Reinforcement Learning for Portfolio Risk Analysis: an
uncertainty-aware policy for capital preservation under regime stress.**

---

## Objectives (latest version)

- **O1.** Investigate whether explicit forecast uncertainty, modelled with a
  DeepAR-style probabilistic LSTM, can be injected into a PPO policy to
  produce risk-aware portfolio behavior on US equity index data.
- **O2.** Determine whether such an uncertainty-aware agent can preserve at
  least 95% of its high-watermark portfolio value across a held-out test
  window that includes shock periods, relative to a baseline PPO and to
  passive buy-and-hold and all-cash benchmarks.
- **O3.** Establish a reproducible evaluation protocol — fixed splits, fixed
  seeds, scripted artifacts, and shared metrics (final value, Sharpe,
  max drawdown, VaR violation rate, preservation ratio) — that supports
  honest comparison between agents.
- **O4.** Recommend, on the strength of the above evidence, when and how an
  uncertainty signal should enter a portfolio control loop, including its
  failure modes and any conditions under which it does not help.

---

## Literature Review (key references)

The list below is the *working* set of key references that have informed the
design of the probabilistic forecasting layer, the RL policy, and the
evaluation protocol. It is deliberately compact (~10 entries) per the form
guidance.

1. **Schulman, Wolski, Dhariwal, Radford, Klimov (2017) — “Proximal Policy
   Optimization Algorithms.”** *arXiv:1707.06347.*
   PPO is the policy-gradient algorithm used for both the baseline and the
   uncertainty-aware agent in this project. The clipped objective and
   stability properties are why PPO was chosen over vanilla policy gradients
   or TRPO for a single-asset trading environment.

2. **Sutton & Barto (2018) — *Reinforcement Learning: an Introduction*,
   2nd ed., MIT Press.**
   Provides the foundational MDP framing, return formulations, and policy
   evaluation/improvement language used throughout the methodology chapter.

3. **Salinas, Flunkert, Gasthaus, Januschowski (2020) — “DeepAR:
   Probabilistic Forecasting with Autoregressive Recurrent Networks.”**
   *International Journal of Forecasting, 36(3), 1181–1191.*
   Source of the probabilistic forecasting design pattern — RNN that emits
   distribution parameters (mean and variance) trained by Gaussian negative
   log-likelihood. The implementation in `experiments/run_probabilistic_agent.py`
   is a DeepAR-style LSTM emitting (μ, log σ²).

4. **Hochreiter & Schmidhuber (1997) — “Long Short-Term Memory.”** *Neural
   Computation, 9(8), 1735–1780.*
   Architectural basis of the probabilistic forecaster.

5. **Lakshminarayanan, Pritzel, Blundell (2017) — “Simple and Scalable
   Predictive Uncertainty Estimation using Deep Ensembles.”** *NeurIPS.*
   Background reading on tractable uncertainty estimation in neural
   networks; informs the choice to use predictive variance as the
   uncertainty signal rather than Bayesian posteriors directly.

6. **Gal & Ghahramani (2016) — “Dropout as a Bayesian Approximation:
   Representing Model Uncertainty in Deep Learning.”** *ICML.*
   Theoretical underpinning for treating learned predictive variance as a
   meaningful epistemic/aleatoric proxy.

7. **Jiang, Xu, Liang (2017) — “A Deep Reinforcement Learning Framework
   for the Financial Portfolio Management Problem.”** *arXiv:1706.10059.*
   Closely related work on RL-based portfolio control. Used to position the
   contribution of this project (uncertainty-aware *risk* policy, not
   purely return maximization).

8. **Yang, Liu, Zhong, Walid (2020) — “Deep Reinforcement Learning for
   Automated Stock Trading: An Ensemble Strategy.”** *ICAIF.*
   Comparative reference for protocol design (train/validation/test splits,
   seed averaging, benchmark inclusion).

9. **Liu, Yang, Gao, Wang (2021) — “FinRL: a deep reinforcement learning
   library for automated stock trading in quantitative finance.”** *NeurIPS
   Deep RL Workshop / arXiv:2011.09607.*
   Reference design of finance-RL environments. The custom `StockEnv` in
   `experiments/common.py` was deliberately written from scratch for
   transparency; FinRL is cited to motivate the environment shape.

10. **Raffin, Hill, Gleave, Kanervisto, Ernestus, Dormann (2021) —
    “Stable-Baselines3: Reliable Reinforcement Learning Implementations.”**
    *Journal of Machine Learning Research, 22(268), 1–8.*
    Implementation source for the PPO solver used in both runners.

11. **Markowitz (1952) — “Portfolio Selection.”** *Journal of Finance,
    7(1), 77–91.*
    Used to frame the risk/return objective and to motivate why
    capital-preservation rather than mean–variance is the chosen objective.

---

## Technical progress

### Summary

A clean, reproducible Phase-0 → Phase-1 pipeline is now in place. The
project compares **baseline PPO** against a **probabilistic-PPO** variant
that consumes a DeepAR-style uncertainty signal, both evaluated against
**buy-and-hold** and **all-cash** benchmarks under a single shared protocol
on `SPY` (with `QQQ` queued as a robustness ticker).

### What has been built

- **Probabilistic forecaster** (`experiments/run_probabilistic_agent.py`):
  An LSTM trained with Gaussian NLL emits (μ, log σ²) for next-step log
  returns; predictive σ is min-max normalised into an uncertainty score in
  `[0, 1]`.
- **Uncertainty-aware trading environment** (`experiments/common.py:StockEnv`):
  - Action ∈ `[-1, 1]` over a `max_trade_fraction` of cash, scaled by
    `(1 - uncertainty_level)` with a floor `min_trade_scale`.
  - High-uncertainty regime (signal above the protocol quantile, default
    `0.80`) **blocks new risk-on buys** while still allowing exits.
  - Reward = log of next-step portfolio-value ratio × 100 (encourages
    compounding while penalising drawdowns).
- **Baseline PPO runner** (`experiments/run_baseline.py`): identical
  environment, no uncertainty signal/guard, used as the controlled
  comparator.
- **Benchmark runner** (`experiments/run_benchmarks.py`): buy-and-hold
  and all-cash on the same test window for sanity checks.
- **Evaluation protocol** (`experiments/configs/dissertation_protocol.json`):
  splits **2009–2018 train / 2019–2021 validation / 2022–2025 test**, seeds
  `[7, 19, 42]`, and a fixed metric set (final value, annualised return /
  vol, Sharpe, max drawdown, VaR-95 + violation rate, capital-preservation
  rate vs high-watermark, ≥ 0.95 goal flag).
- **Reporting layer**: `reports/generate_dissertation_report.py` and
  `reports/build_supervisor_pack.py` materialise the evidence in
  `reports/generated/` (markdown summary + supervisor chart). Richer
  visuals are in `reports/plot_dissertation_visuals.py`.

### Phase-0 → Phase-1 status table

| Step | Status | Notes |
|------|--------|-------|
| 0.1 Environment + dependencies | Done | `requirements.txt`, SB3, PyTorch, gymnasium, yfinance |
| 0.2 PPO baseline on sample data | Done | `phase0_examples/ppo_stock_trading_standalone.py` |
| 0.3 DeepAR-style probabilistic example | Done | `phase0_examples/deepar_style_example.py` |
| 1.1 Shared protocol + metrics | Done | `experiments/configs/dissertation_protocol.json`, `experiments/common.py` |
| 1.2 Reproducible baseline / probabilistic / benchmark runners | Done | three runners, seeded |
| 1.3 Dissertation report + supervisor pack | Done | `reports/generated/` |
| 1.4 Robustness (multi-ticker, ablations, shock windows) | In progress | Phase-2 (see plan) |

### Current results (mean across 3 seeds, test window)

| Agent | Final value (USD) | Sharpe | Max drawdown | VaR-95 violation | Preservation vs HWM |
|---|---:|---:|---:|---:|---:|
| baseline PPO            | 985,463.41 | −0.4285 | 0.0209 | 0.0105 | 0.9811 |
| probabilistic PPO       | 1,618,577.16 | 0.8511 | 0.1833 | 0.0500 | 0.9965 |
| buy-and-hold (SPY)      | 1,520,353.38 | — | 0.2450 | — | — |

Reference figure: `reports/generated/charts/final_value_comparison.png`
(equity curves and uncertainty signal in `equity_curve_comparison.png` and
`uncertainty_signal.png`).

### Honest interpretation

- The probabilistic agent **beats the baseline PPO and the passive
  buy-and-hold** on final value and on the project’s headline objective
  (preservation vs high-watermark, ≥ 0.95).
- The **max-drawdown comparison is misleading** taken alone: the baseline
  PPO barely moves from initial capital, so it cannot draw down by much,
  while the probabilistic agent first compounds to a higher peak and then
  experiences a larger absolute drawdown back to a still-higher final
  value. The preservation ratio is the more faithful measure of the stated
  objective; max drawdown is reported for transparency, not as a
  contradicting result.
- All numbers are **provisional** pending the robustness work in the
  next-steps plan (multi-ticker, shock windows, threshold sensitivity,
  ablation between *uncertainty as state feature* vs *uncertainty as a
  trading guard*).

### Reproducibility

```bash
python experiments/run_baseline.py
python experiments/run_probabilistic_agent.py
python experiments/run_benchmarks.py
python reports/generate_dissertation_report.py
python reports/build_supervisor_pack.py
python reports/plot_dissertation_visuals.py
```

Artifacts land in `experiments/results/` and `reports/generated/`. Source
code is on GitHub at `TheFinix13/Dissertation_Sample_Project`.

---

## Future plan

The plan is sized to the form’s phasing (10–11 weeks across June–July,
4 weeks in August, September for viva). Milestones map to objectives O1–O4.

| Working period | Tasks to undertake | Milestones to meet (with target dates) |
|---|---|---|
| **May 2026 (remaining)** | Threshold sensitivity (`uncertainty_quantile_stop` ∈ {0.7, 0.8, 0.9}); ablation: *uncertainty as state feature only* vs *as trading guard*; rerun with extended timesteps. | M1: ablation results checked into `experiments/results/` (by **end of May**). |
| **June 2026 (4 weeks)** | Multi-ticker robustness (`SPY`, `QQQ`, sector ETFs); event-window analysis on the protocol shock periods (COVID crash; Ukraine-war onset). Begin Chapter 2 (Background) and Chapter 3 (Methodology) drafts. | M2: multi-ticker + shock-window report with stable conclusions; M3: Chapter 2 draft submitted to supervisor. (by **end of June**). |
| **July 2026 (4–6 weeks)** | Sensitivity to environment choices (transaction cost, max trade fraction, lookback). Final experimental sweep with locked seeds. Draft Chapter 4 (Results) and Chapter 1 (Introduction). | M4: locked final results table; M5: Chapters 1–4 first draft (by **mid-late July**). |
| **August 2026 (4 weeks)** | Chapter 5 (Discussion) and Chapter 6 (Conclusion). Polish figures, integrate supervisor feedback, finalise dissertation. Light-touch fixes only on code. | M6: full dissertation draft to supervisor (early August); M7: final submission-ready version (end August). |
| **September 2026** | Viva preparation — slide deck, demo of reproducible pipeline, pre-emptive Q&A using `reports/templates/viva_qa_notes.md`. | M8: viva-ready presentation and demo (by viva date). |

### Risks and mitigations

- **Compute time** — current runners use CPU-friendly settings (10k PPO
  timesteps, 3 seeds). If multi-ticker × shock-window × ablation grows the
  matrix, batch runs overnight; partial-grid results are acceptable for
  the interim.
- **Data API drift** — `yfinance` occasionally changes column shape; the
  `_close_1d` helper in both runners already normalises this.
- **Result fragility** — provisional numbers may move under robustness
  tests. Mitigation: report intervals across seeds and tickers, and lead
  with the preservation objective (the project’s stated goal) rather than
  a single point estimate of final value.

---

## Extenuating circumstances

[FILL IN — e.g., “None to declare,” or describe and indicate that personal
tutor / student support are aware. Do **not** describe medical detail
here.]

---

## Indicative project hours and progress (student self-tick before meeting)

- [ ] The work has exceeded the first 100 hours of time allocated.
- [ ] The work has sufficiently met the first 100 hours.
- [ ] The majority of the first 100 hours have been completed but some time
  has been lost and will be made up.
- [ ] Engagement in the project has been insufficient and progress is of
  concern.

> Recommended self-tick based on the technical evidence above:
> *“The work has sufficiently met the first 100 hours of time allocated to
> the project.”* — supported by the reproducible Phase-0 + Phase-1 pipeline,
> protocol document, baseline + probabilistic agents, benchmarks, and
> generated reports.

---
