# 🧠 AI Chip DCF Valuation Model

![Python](https://img.shields.io/badge/python-3.9%2B-blue) ![License](https://img.shields.io/badge/license-MIT-green) ![Status](https://img.shields.io/badge/build-passing-brightgreen)

A fully dynamic **Discounted Cash Flow (DCF)** model tailored for the rapidly evolving **AI chip industry**.  
Supports three macro‑demand scenarios, generates a formatted Excel workbook, and provides two‑way sensitivity tables for key drivers.

---

## 📊 Model Architecture

```mermaid
graph TD
    A[Assumptions] --> B[Revenue Forecast]
    B --> C[Income Statement]
    C --> D[Free Cash Flow]
    D --> E[DCF Valuation]
    E --> F[Enterprise / Equity Value]
    D --> G[Terminal Value (Gordon)]
    E --> H[Sensitivity: WACC vs g]
The model projects 5 years of financials (2026‑2030) and discounts the resulting free cash flows to firm (FCFF). Terminal value is calculated using the Gordon Growth Model.

🎭 Three Scenarios
Scenario	Description	Revenue CAGR	Gross Margin	Key Driver
🐻 Bear	Slow AI adoption, competition & margin pressure	10%	55%	Weak hyperscaler demand
🏔️ Base	Balanced growth, steady market share	25%	60%	Enterprise AI expansion
🐂 Bull	Explosive AI demand, premium pricing	40%	65%	AGI breakthroughs & sovereign AI
All scenarios adjust operating expenses, D&A, CapEx intensity, and working capital to reflect the underlying operating environment.

📁 File Structure
text
.
├── dcf_model.py                # Python script to generate the Excel model
├── AI_Chip_DCF_Model.xlsx      # Output Excel (generated after run)
├── requirements.txt            # Python dependencies
└── README.md                   # You are here
🚀 Quick Start
Clone the repository

bash
git clone https://github.com/your-username/ai-chip-dcf-model.git
cd ai-chip-dcf-model
Install dependencies

bash
pip install -r requirements.txt
Run the model

bash
python dcf_model.py
The script will output AI_Chip_DCF_Model.xlsx in the current directory.

📈 Excel Output
Sheet Name	Content
Bear_DCF	Full P&L, cash flow and DCF under Bear scenario
Base_DCF	Same structure for Base case (25% CAGR)
Bull_DCF	Bull scenario valuation
Sensitivity	Two‑way table: WACC (8%‑14%) vs Terminal Growth (1%‑4%)
Each DCF sheet includes:

Revenue → Gross Profit → EBITDA → EBIT → NOPAT

Depreciation add‑back, CapEx, ΔNWC

FCFF for 2025‑2030

Discount factors, PV of cash flows, Terminal Value, Enterprise & Equity Value

The Sensitivity sheet uses Base‑case cash flows and is formatted with a red‑yellow‑green color scale.

⚙️ Customization
You can easily modify the assumptions inside dcf_model.py:

python
scenario_params = {
    'Base': {
        'cagr': 0.25,           # Revenue CAGR
        'gross_margin': 0.60,   # Gross margin
        'opex_pct': 0.18,       # Opex as % of revenue
        'capex_pct': 0.12,      # CapEx intensity
        ...
    }
}
WACC, tax rate, and terminal growth are also at the top of the script.

📦 Dependencies
Python 3.9+

openpyxl (for Excel generation)

See requirements.txt.

📄 License
This project is licensed under the MIT License – feel free to use, modify, and distribute.

🌟 Contributing
Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

Built with ❤️ for AI chip valuation enthusiasts.