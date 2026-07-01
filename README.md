# ============================================================
# structDXF - structural DXF Parser & Eurocode 3 Verifier
# Copyright (C) 2025  Ing. Bruno Zilli
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# ============================================================
# DISCLAIMER:
# This software is intended for EDUCATIONAL AND RESEARCH PURPOSES ONLY.
# It has NOT been certified by any professional engineering body.
# Results produced by this software must NOT be used as the sole basis
# for real structural design, analysis, or decision-making.
# Always consult a qualified structural engineer and verify results
# with official standards and validated commercial software.
#
# The authors assume NO LIABILITY for any damages, losses, or
# consequences arising from the use of this software.
# By using this software you agree that you do so at your own risk.
# ============================================================
# ============================================================
# STRUCTURAL DXF PARSER - STATICS v16.3 FINAL CORRETTA
# EN 1993-1-1 + EN 1990 | Class 1 sections | Beam-Column Method 2
# Fixes: equilibrium sign, beam-column export keys, deformed image saving
# ============================================================

markdown
# StructDXF

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

**StructDXF** is a Python tool that reads 2D structural frames directly from a DXF drawing,
builds a finite element model using [PyNite](https://github.com/JWock82/PyNite),
and performs **complete structural verification** according to **EN 1993‑1‑1 (Eurocode 3)**
and **EN 1990 (Load Combinations)**.

It is designed for **educational and research purposes** – quickly prototype a frame,
understand its behaviour, and check the most important Ultimate Limit State (ULS),
Serviceability Limit State (SLS) and buckling requirements.

---

## ✨ Features

- **DXF parsing** – Reads nodes, beams, supports, point loads, distributed loads and
  concentrated moments directly from AutoCAD / LibreCAD DXF files.
- **Automatic mesh refinement** – Splits each member into multiple finite elements for
  accurate internal force diagrams.
- **FEM analysis** – Powered by PyNite (3D frame solver).
- **Interactive load classification** – Assign permanent (G) or variable (Q) loads
  interactively; partial factors γG, γQ, ψ0 are applied automatically.
- **Four graphical outputs:**
  1. **Structural model** – Loads, supports, reactions.
  2. **Internal force diagrams** – Bending moment, shear and axial force for every member.
  3. **Deformed shape** – Scaled deformed geometry with node displacements and rotations.
  4. **Verification summary** – Console tables and an Excel report.
- **Verifications (EN 1993‑1‑1):**
  - Cross‑section resistance (Class 1 plastic moment)
  - Deflection check (characteristic load combination)
  - Flexural buckling (Euler critical load, χ factor)
  - Beam‑column interaction – **Method 2 (Annex B)** with explicit kyy, kzy factors
- **Excel export** – Nine sheets with load factors, reactions, displacements, ULS,
  SLS, buckling, beam‑column details and final summary.
- **Automatic image saving** – All plots are saved as PNG files; if a display is
  available they are also shown on screen.

---

## 📋 Requirements

- Python 3.8 or later
- Required packages:
  - `ezdxf`
  - `numpy`
  - `matplotlib`
  - `Pynite`
  - `openpyxl`
  - `scipy`

Install them with:

```bash
pip install ezdxf numpy matplotlib Pynite openpyxl scipy
On headless systems (no display), matplotlib will automatically use the Agg backend
and save images without trying to open a window.

🚀 Quick start
Clone the repository

bash
git clone https://github.com/[TuoNome]/StructDXF.git
cd StructDXF
Prepare a DXF file

The DXF must contain the following layers:

STRUCT_NODES – circles + MTEXT with node names

STRUCT_BEAMS – lines + MTEXT with member names

STRUCT_SECTIONS – MTEXT with section names (e.g. IPE200)

STRUCT_MATERIALS– MTEXT with steel grade (e.g. S355)

STRUCT_CONSTRAINTS – triangles (fixed or hinge supports)

STRUCT_LOADS – arrows for distributed loads, MTEXT for point loads and moments

See the example file frame-REV04.dxf included in the repository.

Run the analysis

bash
python3 statics.py frame-REV04.dxf
Answer the prompts

Confirm the section properties.

Classify each load as Permanent (G) or Variable (Q).

Choose the mesh density (default 3, recommended 5‑7 for accurate diagrams).

Read the results

Console output: reactions, internal forces, verification tables.

Saved images: frame-REV04_structure.png, frame-REV04_deformed.png, etc.

Excel report: frame-REV04_results.xlsx.

📁 Example output
text
📐 STRUCTURAL VERIFICATION (EN 1993-1-1 + EN 1990)
   γG=1.35, γQ=1.5, ψ0=0.7, γM0=1.0, γM1=1.0

   1️⃣  SLU - CROSS-SECTION RESISTANCE
   Member       N_Ed     N_Rd     N/Nr     M_Ed    Mpl,Rd     M/Mr     V_Ed     V_Rd Status    
   M2           1.14   271.22    0.004     7.24      7.35    0.985     0.00    93.95 ⚡ ACC     

   2️⃣  SLE - DEFLECTION CHECK
   Member      L [m]   δ_max [mm]   δ_lim [mm]    Ratio        Limit Status      
   M2          6.000      66.2258        24.00    2.759        L/250 ⚠️ OVER     

   3️⃣  BUCKLING & BEAM-COLUMN (Method 2)
   Member         ψ    C_my     k_yy     k_zy  N_Ed/Nb_Rd M_Ed/Mpl,Rd   Inter_yy   Inter_zy      Max Status      
   M4        -1.999   1.000   1.4479   0.8687      0.5599      0.3253     1.0308     0.8424   1.0308 ❌ FAIL!     
⚖️ Disclaimer
This software is intended for EDUCATIONAL AND RESEARCH PURPOSES ONLY.
It has NOT been certified by any professional engineering body.
Results produced by this software must NOT be used as the sole basis
for real structural design, analysis, or decision‑making.
Always consult a qualified structural engineer and verify results with
official standards and validated commercial software.

The authors assume NO LIABILITY for any damages, losses, or
consequences arising from the use of this software.
By using this software you agree that you do so at your own risk.

📜 License
This project is licensed under the GNU General Public License v3.0 – see the
LICENSE file for details.

🤝 Contributing
Contributions are welcome!
Please open an issue or a pull request if you find bugs, want to suggest features,
or improve the documentation.

📧 Contact
Author: Ing. Bruno Zilli
Repository: https://github.com/TuoNome/StructDXF

Made with ❤️ and ☕
