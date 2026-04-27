# Probabilistic DRL for Portfolio Risk Analysis

EEEM004 research project: an uncertainty-aware PPO policy for
capital preservation under regime stress.

## For supervisors — single-file walkthrough

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/TheFinix13/Dissertation_Sample_Project/blob/main/Dissertation_Walkthrough.ipynb)

> The fastest path is the badge above. Click it, then *Runtime → Run all*.
> The notebook's first cell automatically clones this repository into the
> Colab session and installs dependencies; nothing has to be checked out
> manually. The walkthrough then loads the SPY test-window dataset, builds
> the DeepAR-style probabilistic forecaster, prints the uncertainty values,
> states the mathematics that differentiates the probabilistic agent from
> the baseline PPO, and renders the comparison table and equity-curve
> plots from the seeded results.
>
> The notebook ships **with executed outputs** so the document is readable
> without running anything.

For a local run instead of Colab:

```bash
git clone https://github.com/TheFinix13/Dissertation_Sample_Project.git
cd Dissertation_Sample_Project
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
pip install jupyter
jupyter notebook Dissertation_Walkthrough.ipynb
```

The interim review draft (Surrey form) lives at
`reports/generated/interim_review_draft.md`.

## Generated artifacts (under `reports/generated/exports/`)

- `Dissertation_Walkthrough.pdf` — fully-rendered PDF of the supervisor walkthrough notebook (with embedded outputs, tables and plots).
- `Dissertation_Draft.docx` — formal Master’s dissertation draft (title page, abstract, 7 chapters, references, appendix; equations rendered as images, figures embedded).
- `equations/` — individual PNGs for every equation in the docx (re-used if the script regenerates the document).

To regenerate:

```bash
venv/bin/python -m nbconvert --to webpdf --allow-chromium-download \
  --output reports/generated/exports/Dissertation_Walkthrough.pdf \
  Dissertation_Walkthrough.ipynb
venv/bin/python reports/build_dissertation_docx.py
```

## Quick Start

```bash
cd portfolio-risk-drl
source venv/bin/activate

# 1. Run PPO stock trading example (Stable-Baselines3)
python phase0_examples/ppo_stock_trading_standalone.py

# 2. Run DeepAR-style probabilistic forecasting example
python phase0_examples/deepar_style_example.py
```

## Project Structure

```
portfolio-risk-drl/
├── venv/                    # Python virtual environment
├── phase0_examples/
│   ├── ppo_stock_trading_standalone.py   # PPO agent (SB3) - works
│   ├── deepar_style_example.py           # Probabilistic LSTM - works
│   └── finrl_ppo_example.py              # FinRL+PPO (needs extra deps)
├── trained_models/          # Saved PPO models
├── requirements.txt
└── README.md
```

## Phase 0 Status

| Step | Status | Notes |
|------|--------|-------|
| 0.1 Virtual env + packages | ✅ | pandas, numpy, SB3, PyTorch, yfinance |
| 0.2 PPO on sample data | ✅ | Use `ppo_stock_trading_standalone.py` |
| 0.3 DeepAR-style example | ✅ | Use `deepar_style_example.py` |

### GluonTS / FinRL Notes

- **GluonTS**: Requires `scipy<1.16` which has no pre-built wheel for Python 3.14. Use Python 3.10–3.11 and `pip install gluonts[torch]` for full GluonTS DeepAR.
- **FinRL**: Has many optional dependencies (alpaca, wrds, elegantrl). The standalone PPO example achieves the same learning objective without FinRL.

## Next Steps (Phase 1)

1. Integrate DeepAR-style outputs (mean, variance) into PPO state
2. Add risk thresholds based on predicted uncertainty
3. Benchmark on S&P 500 / NASDAQ 100 (2009–2025)

## Phase 1 Experiment Pipeline

```bash
source venv/bin/activate

# 1) Baseline PPO with deterministic seeds
python experiments/run_baseline.py

# 2) Probabilistic DeepAR-style uncertainty + PPO
python experiments/run_probabilistic_agent.py

# 3) Benchmarks for sanity checks
python experiments/run_benchmarks.py

# 4) Generate dissertation-ready summary report
python reports/generate_dissertation_report.py

# 5) Build supervisor pack (progress report + chart)
python reports/build_supervisor_pack.py

# 6) Build richer visualizations (equity curves, uncertainty, dataset, intraday)
python reports/plot_dissertation_visuals.py
```

- Protocol config: `experiments/configs/dissertation_protocol.json`
- Artifacts: `experiments/results/`
- Report output: `reports/generated/dissertation_results.md`
