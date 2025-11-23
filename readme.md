# Engineering Apps Suite

A comprehensive collection of Python-based engineering tools designed to assist civil engineers and construction professionals with concrete mix design, rebar optimization, and structural detailing.

The suite features a unified launcher interface built with PyQt6, providing access to three specialized applications.

## 🚀 Applications Included

### 1. Foundation Cutting List
An automated detailing tool for reinforced concrete footings.
*   **Visual Input:** Define footing geometry (pads, pedestals) and reinforcement parameters (cover, hooks, bar sizes).
*   **Stirrup Designer:** Visual editor for complex stirrup spacing and bundle configurations.
*   **Excel Output:** Generates a professional Excel schedule containing:
    *   Visual bending schedules with shape diagrams.
    *   Optimized purchase lists.
    *   Specific cutting plans to minimize waste.

### 2. Rebar Optimal Purchase
A linear programming optimization tool for solving the 1D Cutting Stock Problem.
*   **Cost Reduction:** Calculates the mathematical optimum to minimize steel waste.
*   **Flexible Inputs:** Handle multiple bar diameters and custom required cut lengths.
*   **Market Length Selection:** Choose from standard market lengths (6m, 7.5m, 9m, etc.) or define custom stocks.
*   **Reporting:** Generates an Excel report detailing exactly which stock bars to buy and how to cut them.

### 3. Concrete Mix Design
A mix proportioning tool based on **ACI 211.1-22** standards.
*   **Precision Calculation:** Calculates weights and absolute volumes for Cement, Water, Coarse/Fine Aggregates, and Air.
*   **Moisture Adjustment:** Automatically adjusts batch weights based on aggregate absorption and field moisture content.
*   **Durability Checks:** Validates w/cm ratios against exposure classes (F, S, W, C).
*   **Strength Estimator:** Includes a "Strength vs. Age" graphing tool using the Dreux-Gorisse and GL2000 maturity models.

## 🛠️ Technology Stack

*   **Language:** Python 3.x
*   **GUI:** PyQt6 (Modern, styled interface with QSS)
*   **Optimization:** PuLP (Linear Programming interface) & CBC Solver
*   **Reporting:** OpenPyXL (Excel generation), Pillow (Image processing)
*   **Plotting:** Matplotlib (Concrete strength graphs)

## 📦 Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/RSv618/engineering_apps.git
    cd engineering_apps
    ```

2.  **Create a virtual environment (optional but recommended):**
    ```bash
    python -m venv .venv
    # Windows
    .venv\Scripts\activate
    # Mac/Linux
    source .venv/bin/activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## ▶️ Usage

To start the main suite launcher:

```bash
python app_launcher.py