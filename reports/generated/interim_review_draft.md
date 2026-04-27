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

I have refined the original objectives during Phase 0 and Phase 1. The current
working set is:

- **O1.** See whether an explicit forecast-uncertainty signal — produced by a
  DeepAR-style probabilistic LSTM — can be plugged into a PPO policy and turn
  it into something that behaves with risk in mind on US equity index data.
- **O2.** Show, on a held-out test window that contains real shock periods,
  whether the resulting agent preserves at least 95% of its high-watermark
  portfolio value relative to a baseline PPO and to passive buy-and-hold and
  all-cash benchmarks.
- **O3.** Pin down a reproducible evaluation protocol: fixed splits, fixed
  seeds, scripted artifacts, and a small shared metric set (final value,
  Sharpe, max drawdown, VaR-95 violation rate, preservation ratio against the
  running high-watermark). The protocol is the thing that lets me make a
  fair, like-for-like claim.
- **O4.** Based on the evidence from O1–O3, take a position on when an
  uncertainty signal helps a portfolio control loop, and — just as
  importantly — when it doesn't.

---

## Literature Review (key references)

The list below is the working set of references I keep coming back to. It is
intentionally compact (around ten entries) per the form guidance; the full
bibliography in the dissertation will be longer.

1. **Schulman, Wolski, Dhariwal, Radford, Klimov (2017) — "Proximal Policy
   Optimization Algorithms."** *arXiv:1707.06347.*
   PPO is the policy-gradient algorithm I use for both the baseline and the
   uncertainty-aware variant. The clipped surrogate is the reason I picked
   PPO over vanilla policy gradient or TRPO: it is forgiving with respect to
   hyper-parameters, which matters on a single-asset environment with limited
   training budget.

2. **Sutton & Barto (2018) — *Reinforcement Learning: an Introduction*,
   2nd ed., MIT Press.**
   The standard reference for the MDP formulation and the policy-gradient
   derivations I rely on in the methodology chapter.

3. **Salinas, Flunkert, Gasthaus, Januschowski (2020) — "DeepAR:
   Probabilistic Forecasting with Autoregressive Recurrent Networks."**
   *International Journal of Forecasting, 36(3), 1181–1191.*
   The blueprint for my forecaster: an RNN that emits the parameters of a
   predictive distribution and is trained by negative log-likelihood. My
   implementation in `experiments/run_probabilistic_agent.py` is a stripped-
   down DeepAR with a Gaussian head emitting (μ, log σ²).

4. **Hochreiter & Schmidhuber (1997) — "Long Short-Term Memory."** *Neural
   Computation, 9(8), 1735–1780.*
   The architectural workhorse inside the forecaster.

5. **Lakshminarayanan, Pritzel, Blundell (2017) — "Simple and Scalable
   Predictive Uncertainty Estimation using Deep Ensembles."** *NeurIPS.*
   I read this to decide whether to go with deep ensembles or a single
   network with a Gaussian head. The single-head approach is cheaper and
   sufficient given that the policy only needs a one-dimensional uncertainty
   summary; ensembles are noted as future work.

6. **Gal & Ghahramani (2016) — "Dropout as a Bayesian Approximation:
   Representing Model Uncertainty in Deep Learning."** *ICML.*
   Background reading on why a learned predictive variance can be treated as
   a meaningful uncertainty proxy at all.

7. **Jiang, Xu, Liang (2017) — "A Deep Reinforcement Learning Framework
   for the Financial Portfolio Management Problem."** *arXiv:1706.10059.*
   The closest prior work to mine in spirit. They use RL for portfolio
   selection but optimise for return; my contribution is on the *risk* side,
   so I cite them to position what is and isn't new in my approach.

8. **Yang, Liu, Zhong, Walid (2020) — "Deep Reinforcement Learning for
   Automated Stock Trading: An Ensemble Strategy."** *ICAIF.*
   I borrowed the train / validation / test split style and the seed-
   averaging convention from this paper.

9. **Liu, Yang, Gao, Wang (2021) — "FinRL: a deep reinforcement learning
   library for automated stock trading in quantitative finance."**
   *arXiv:2011.09607.*
   The reference design for finance-RL environments. I deliberately wrote my
   own `StockEnv` from scratch (in `experiments/common.py`) rather than
   subclassing FinRL, because I wanted to be able to explain every line in
   the viva. FinRL is cited to motivate the shape of the environment.

10. **Raffin, Hill, Gleave, Kanervisto, Ernestus, Dormann (2021) —
    "Stable-Baselines3: Reliable Reinforcement Learning Implementations."**
    *Journal of Machine Learning Research, 22(268), 1–8.*
    The implementation source for the PPO solver in both runners.

11. **Markowitz (1952) — "Portfolio Selection."** *Journal of Finance,
    7(1), 77–91.*
    Cited to frame why I picked capital preservation rather than mean–
    variance as the objective. Mean–variance is the conceptual baseline
    every risk-aware portfolio paper compares against, and it is worth
    being explicit about why I am not using it directly.

---

## Technical progress

### Summary

Phase 0 and Phase 1 are in place and reproducible end-to-end. The dissertation
compares **baseline PPO** against a **probabilistic-PPO** variant that
consumes a DeepAR-style uncertainty signal, with **buy-and-hold** and
**all-cash** as benchmarks, all on `SPY` for now (`QQQ` is queued for the
multi-ticker robustness study in Phase 2). Everything runs from a single
config file and a small set of scripts.

### What has been built

- A **probabilistic forecaster** (`experiments/run_probabilistic_agent.py`):
  an LSTM trained with Gaussian NLL that emits the mean and log variance of
  the next-step log return. The predictive standard deviation is min-max
  normalised across the test window into a unit-interval uncertainty score.
- An **uncertainty-aware trading environment**
  (`experiments/common.py:StockEnv`):
  - Action space `[-1, 1]` over a configurable `max_trade_fraction` of cash,
    with the trade size shrunk by `(1 - uncertainty_level)` and floored at
    `min_trade_scale` so the agent is never silenced entirely.
  - When the uncertainty signal sits above the protocol quantile (default
    `0.80`) the environment **blocks new long-side trades** but still allows
    exits.
  - Reward is the per-step log of the portfolio-value ratio, multiplied by
    100 for numerical scale. This rewards compounding and penalises
    drawdowns automatically, without an extra term.
- A **baseline PPO runner** (`experiments/run_baseline.py`) that uses the
  same environment without the uncertainty coordinate or the trade-size
  shrinkage, so the comparison is genuinely controlled.
- A **benchmarks runner** (`experiments/run_benchmarks.py`) that evaluates
  buy-and-hold and all-cash on the same test window. These act as sanity
  checks on the metric definitions as much as competitors to beat.
- A single **evaluation protocol**
  (`experiments/configs/dissertation_protocol.json`) that fixes the splits
  (2009–2018 train / 2019–2021 validation / 2022–2025 test), the seeds
  (`[7, 19, 42]`) and the metric set, and is read by every script. This is
  the bit that actually makes the comparisons fair.
- A **reporting layer**: `reports/generate_dissertation_report.py` produces
  the markdown summary, `reports/build_supervisor_pack.py` produces the
  one-page chart, and `reports/plot_dissertation_visuals.py` produces the
  detailed figures. There is also a `Dissertation_Walkthrough.ipynb` that
  re-runs the whole pipeline and renders the embedded outputs for review.

### Phase-0 → Phase-1 status table

| Step | Status | Notes |
|------|--------|-------|
| 0.1 Environment + dependencies | Done | `requirements.txt`, SB3, PyTorch, gymnasium, yfinance |
| 0.2 PPO baseline on sample data | Done | `phase0_examples/ppo_stock_trading_standalone.py` |
| 0.3 DeepAR-style probabilistic example | Done | `phase0_examples/deepar_style_example.py` |
| 1.1 Shared protocol + metrics | Done | `experiments/configs/dissertation_protocol.json`, `experiments/common.py` |
| 1.2 Reproducible baseline / probabilistic / benchmark runners | Done | three runners, seeded |
| 1.3 Dissertation report + supervisor pack | Done | `reports/generated/` |
| 1.4 Robustness (multi-ticker, ablations, shock windows) | In progress | Phase 2 (see plan) |

### Current results (mean across 3 seeds, test window)

| Agent | Final value (USD) | Sharpe | Max drawdown | VaR-95 violation | Preservation vs HWM |
|---|---:|---:|---:|---:|---:|
| baseline PPO            | 985,463.41 | −0.4285 | 0.0209 | 0.0105 | 0.9811 |
| probabilistic PPO       | 1,618,577.16 | 0.8511 | 0.1833 | 0.0500 | 0.9965 |
| buy-and-hold (SPY)      | 1,520,353.38 | — | 0.2450 | — | — |

Reference figure: `reports/generated/charts/final_value_comparison.png`.
Equity curves and the uncertainty signal are in
`equity_curve_comparison.png` and `uncertainty_signal.png` respectively.

### How to read these numbers

A few things are worth flagging before this table is read in isolation:

- The probabilistic agent meets the headline objective (preservation against
  the high-watermark above 0.95) and finishes above passive buy-and-hold.
  The baseline ends roughly where it started.
- Max drawdown on the baseline looks small only because the baseline barely
  compounds in the first place. There is little to draw down from. The
  probabilistic agent compounds to a higher peak, gives some of it back,
  and still finishes well above the baseline. Preservation against the
  high-watermark is the metric that matches my objective; max drawdown is
  reported for transparency, not as a contradicting result.
- These numbers are provisional. The plan below explicitly tests how
  fragile they are to ticker choice, threshold choice, and which part of
  the design is doing the work (the state feature, the trade-size shrink,
  or the entry guard).

### Reproducibility

```bash
python experiments/run_baseline.py
python experiments/run_probabilistic_agent.py
python experiments/run_benchmarks.py
python reports/generate_dissertation_report.py
python reports/build_supervisor_pack.py
python reports/plot_dissertation_visuals.py
```

Artifacts land in `experiments/results/` and `reports/generated/`. The full
source is on GitHub at `TheFinix13/Dissertation_Sample_Project`, with the
walkthrough notebook (`Dissertation_Walkthrough.ipynb`) as the single entry
point for someone reading the project for the first time.

---

## Future plan

The phasing below maps onto the form's structure (10–11 weeks across June and
July, four weeks in August, September for the viva). Milestones are tied to
objectives O1–O4.

| Working period | Tasks to undertake | Milestones to meet (with target dates) |
|---|---|---|
| **May 2026 (remaining)** | Sweep `uncertainty_quantile_stop` over {0.7, 0.8, 0.9}; ablation of the design — *uncertainty as a state feature only* vs *as a trading guard only* vs *both*; rerun the protocol with longer training to check the Phase-1 numbers are not under-trained. | M1: ablation results checked into `experiments/results/` (by **end of May**). |
| **June 2026 (4 weeks)** | Multi-ticker robustness (`SPY`, `QQQ`, plus a small set of sector ETFs); event-window analysis on the protocol shock periods (COVID crash, Ukraine-war onset). Begin Chapter 2 (Background) and Chapter 3 (Methodology) drafts. | M2: multi-ticker and shock-window report with stable conclusions; M3: Chapter 2 draft to supervisor (by **end of June**). |
| **July 2026 (4–6 weeks)** | Sensitivity to environment choices (transaction cost, max trade fraction, lookback). Final experimental sweep with locked seeds. Draft Chapter 5 (Results) and Chapter 1 (Introduction). | M4: locked final results table; M5: Chapters 1, 2, 3 and 5 first draft (by **mid- to late-July**). |
| **August 2026 (4 weeks)** | Chapter 6 (Discussion) and Chapter 7 (Conclusion). Polish figures, integrate supervisor feedback, finalise the dissertation. Code changes from this point are bug-fix only. | M6: full draft to supervisor early August; M7: submission-ready version end August. |
| **September 2026** | Viva preparation. Slide deck, demo of the reproducible pipeline, pre-emptive Q&A using `reports/templates/viva_qa_notes.md`. | M8: viva-ready presentation and demo by viva date. |

### Risks and mitigations

- **Compute time.** All current runs are CPU-friendly (10k PPO timesteps,
  three seeds). If the multi-ticker × shock-window × ablation grid grows,
  I will batch runs overnight and accept partial-grid results for the
  interim. None of the experiments need a GPU at this scale.
- **Data-API drift.** `yfinance` occasionally changes its column shape. The
  `_close_1d` helper used by every runner already normalises this, and the
  protocol pins explicit dates so a re-pull stays comparable.
- **Result fragility.** The Phase-1 numbers may move under the multi-ticker
  and ablation work. To guard against over-claiming, I will report ranges
  across seeds and tickers in the dissertation, lead with the preservation
  objective rather than a single point estimate of final value, and call
  out any case where the probabilistic variant fails to beat the baseline.

---

## Extenuating circumstances

[FILL IN — for example "None to declare", or describe and indicate that the
personal tutor and student-support services have been informed. Do **not**
include medical detail here.]

---

## Indicative project hours and progress (student self-tick before meeting)

- [ ] The work has exceeded the first 100 hours of time allocated.
- [ ] The work has sufficiently met the first 100 hours.
- [ ] The majority of the first 100 hours have been completed but some time
  has been lost and will be made up.
- [ ] Engagement in the project has been insufficient and progress is of
  concern.

> Recommended self-tick, given the evidence above: *"The work has
> sufficiently met the first 100 hours of time allocated to the project."*
> The reproducible Phase-0 and Phase-1 pipeline, the protocol document, the
> baseline and probabilistic agents, the benchmarks and the generated
> reports together support this self-assessment.

---
