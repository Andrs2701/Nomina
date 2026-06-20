# Sistema de Liquidación de Nómina 🇨🇴

<div align="center">

**Parametric payroll calculation engine for Colombian labor law**

Handles the full complexity of Colombian payroll — contracts, overtime, bonuses, social benefits, and deductions — in a single web application.

[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![Flask](https://img.shields.io/badge/Flask-3.x-000000?style=flat-square&logo=flask&logoColor=white)](https://flask.palletsprojects.com)
[![SQLite](https://img.shields.io/badge/SQLite-003B57?style=flat-square&logo=sqlite&logoColor=white)](https://sqlite.org)
[![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20Export-217346?style=flat-square&logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io)
[![Tests](https://img.shields.io/badge/Tests-pytest-0A9EDC?style=flat-square&logo=pytest&logoColor=white)](https://pytest.org)

[![Live Demo](https://img.shields.io/badge/Live%20Demo-Render-46E3B7?style=flat-square&logo=render&logoColor=white)](https://nomina-xoq1.onrender.com)

</div>

---

## 🎯 What Problem Does It Solve?

Calculating payroll in Colombia is genuinely complex. The labor code mandates specific rules for overtime rates, night shifts, Sunday premiums, transportation allowances, social benefits (cesantías, intereses, prima, vacaciones), and parafiscal contributions — all of which change based on salary and contract type.

Most companies either pay expensive ERP licenses (SAP, Siesa, Helisa) or struggle with error-prone Excel sheets that break every time labor rules change.

**This system replaces that** with a parametric engine where every rule is configurable, every calculation is auditable, and the output is a clean professional Excel report ready for payroll signing.

---

## ✨ Key Features

### 🧮 Payroll Calculation Engine (`engine.py`)
- **Full Colombian labor law compliance** — Código Sustantivo del Trabajo
- **All shift types** — Diurno, Nocturno, Dominical Diurno, Dominical Nocturno, Festivo Diurno, Festivo Nocturno
- **Overtime calculation** — Hora Extra Diurna (1.25x), Nocturna (1.75x), Dominical/Festivo Diurna (2.0x), Dominical/Festivo Nocturna (2.5x)
- **Social benefits** — Cesantías, Intereses sobre cesantías, Prima de servicios, Vacaciones
- **Deductions** — Salud (4%), Pensión (4%), Libranza, embargos, otros descuentos
- **Transportation allowance** — Auto-applied based on salary threshold
- **Configurable parameters** — SMMLV, UVT, auxilio de transporte all adjustable per year

### 🌐 Web Interface (Flask)
- Upload employee data via Excel template or manual entry
- Real-time payroll preview before generating reports
- Multi-employee batch processing
- Period selection (monthly, quincenal, or custom)

### 📊 Excel Export (`openpyxl`)
- Professional styled payroll reports with company header
- Individual payment slips (comprobantes de pago) per employee
- Summary report with totals per cost center
- Formatted for printing and digital signature

### ⚙️ Configuration & Persistence
- SQLite database for saving employee profiles and payroll parameters
- Import employee roster from Excel template (`Plantilla_Nomina.xlsx`)
- Export complete payroll to ZIP (individual slips + summary)
- All parameters (SMMLV, recargos, contributions) configurable via admin panel

### 🧪 Test Coverage
- `test_engine.py` — Unit tests for all overtime and benefit calculations
- Edge cases: salary changes mid-period, partial months, disability days

---

## 🛠️ Tech Stack

| Component | Technology |
|-----------|-----------|
| **Backend / API** | Python 3.11+, Flask 3.x |
| **Calculation Engine** | Pure Python (`engine.py`) |
| **Database** | SQLite via `sqlite3` |
| **Excel Processing** | openpyxl (import + styled export) |
| **Frontend** | Jinja2 templates, HTML5, CSS3 |
| **Testing** | pytest |
| **Deployment** | Render (WSGI via `wsgi.py`) |

---

## 📁 Project Structure

```
Nomina/
├── app.py               # Flask application — routes, DB, export logic
├── engine.py            # Payroll calculation engine (pure Python)
├── test_engine.py       # Unit tests for all calculation rules
├── wsgi.py              # WSGI entry point for production deployment
├── requirements.txt     # Python dependencies
├── render.yaml          # Render deployment configuration
│
├── templates/           # Jinja2 HTML templates
│   ├── index.html       # Main dashboard
│   ├── empleados.html   # Employee management
│   ├── liquidar.html    # Payroll calculation form
│   └── resultado.html   # Results preview + export
│
├── static/              # CSS, JS, assets
│   └── style.css
│
├── data/
│   └── nomina.db        # SQLite database (auto-created)
│
└── Plantilla_Nomina.xlsx  # Employee import template
```

---

## 🚀 Getting Started

### Prerequisites
- Python 3.11+
- pip

### Installation

```bash
# Clone the repository
git clone https://github.com/Andrs2701/Nomina.git
cd Nomina

# Create virtual environment
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
# Available at http://localhost:5000
```

### Run tests
```bash
pytest test_engine.py -v
```

---

## 🧮 Calculation Rules Reference

### Recargos (Overtime & Shift Premiums)

| Shift Type | Rate |
|-----------|------|
| Hora Extra Diurna | 125% |
| Hora Extra Nocturna | 175% |
| Recargo Nocturno (no extra) | 135% |
| Hora Extra Dominical / Festiva Diurna | 200% |
| Hora Extra Dominical / Festiva Nocturna | 250% |
| Trabajo Dominical / Festivo (sin extra) | 175% |

### Prestaciones Sociales

| Benefit | Rate | Period |
|---------|------|--------|
| Cesantías | 8.33% of monthly salary | Annual |
| Intereses sobre cesantías | 12% of cesantías | Annual |
| Prima de servicios | 8.33% of monthly salary | Bi-annual |
| Vacaciones | 4.17% of monthly salary | Annual |

### Aportes Parafiscales (Employer)

| Contribution | Rate |
|-------------|------|
| Salud | 8.5% |
| Pensión | 12% |
| ARL | Variable by risk level |
| ICBF | 3% |
| SENA | 2% |
| Caja de Compensación | 4% |

### Deducciones (Employee)

| Deduction | Rate |
|-----------|------|
| Salud | 4% |
| Pensión | 4% |
| Retención en la fuente | Variable (UVT-based) |

---

## 🗺️ Roadmap

- [x] Complete overtime and shift premium calculation
- [x] Social benefits (cesantías, prima, vacaciones, intereses)
- [x] Employee and employer parafiscal contributions
- [x] Excel import (employee roster) and export (payment slips)
- [x] Unit tests for calculation engine
- [x] SQLite persistence for employee profiles
- [x] Production deployment on Render
- [ ] Multi-company support
- [ ] Digital payslip delivery by email
- [ ] Retención en la fuente calculation (full UVT table)
- [ ] Integration with Colombian DIAN e-invoicing
- [ ] REST API for ERP integration
- [ ] Disability and maternity leave handling
- [ ] PostgreSQL migration for multi-tenant production

---

## 📋 Colombian Labor Law Context

This system implements the **Código Sustantivo del Trabajo** (CST) and annual regulatory updates including:
- Decreto de salario mínimo (SMMLV) — updated yearly
- Decreto de auxilio de transporte — updated yearly
- Artículos 159–171 CST — overtime and shift rules
- Artículos 249–259 CST — social benefits (cesantías, prima)
- Ley 21 de 1982 — parafiscal contributions (SENA, ICBF, Cajas)

> ⚠️ **Disclaimer:** This software is provided for educational and automation purposes. Always verify calculations with a certified Colombian accountant (contador público) for official payroll processing.

---

## 👤 Author

**Camilo Andrés Chitiva Castelblanco**
- LinkedIn: [linkedin.com/in/andres-chitiva-204a4b259](https://linkedin.com/in/andres-chitiva-204a4b259)
- GitHub: [@Andrs2701](https://github.com/Andrs2701)
- Email: andrscc2701@gmail.com

*Built as an independent software project for Colombian SMEs seeking an affordable, transparent payroll solution.*

---

<div align="center">

⭐ Found this useful? Give it a star — it helps others find this project.

**[🚀 Try the Live Demo](https://nomina-xoq1.onrender.com)**

</div>
