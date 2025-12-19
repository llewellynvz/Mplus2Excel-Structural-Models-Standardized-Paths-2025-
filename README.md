# 📊 Mplus2Excel – Structural Models & Standardized Paths (2025)

### **Automatically generate publication-ready structural model tables from Mplus output files**

---

## 📌 What is this project?

Reporting and comparing **structural models** in **Mplus** often requires manually extracting fit indices, standardized paths, and explained variance from multiple output files. This process is time-consuming, error-prone, and difficult to reproduce.

**Mplus2Excel** is a lightweight utility that scans a folder of **Mplus structural model `.out` files**, automatically:

- Aggregates model fit indices across competing structural models
- Produces clean **structural model comparison tables**
- Extracts **STDYX standardized path coefficients**
- Extracts **R² values** for endogenous variables
- Compiles everything into a single, publication-ready **Excel workbook**

---

## 🎯 What does it do exactly?

When run in a folder containing *only* Mplus **structural model output files**, the tool will:

1. **Scan all `.out` files** in the folder  
2. **Extract key fit indices** (e.g. χ², df, CFI, TLI, RMSEA, SRMR, AIC, BIC)  
3. **Compile all models into a single structural comparison table**  
4. **Extract STDYX standardized structural paths**  
5. **Extract R² values** for all endogenous variables  
6. Export all results to a single Excel workbook  

This eliminates repetitive copy-paste work and substantially reduces reporting errors.

---

## ⚠ Important usage constraint (read this)

> **The folder must contain only Mplus structural model `.out` files.**

- Do **not** include measurement models  
- Do **not** include auxiliary analyses  
- Do **not** mix different projects  

All detected `.out` files are treated as competing structural models and will be included in the same comparison tables.

---

## 🤔 Why is this useful?

I built this tool after repeatedly spending hours manually extracting structural paths, fit indices, and R² values for manuscripts, reviewer responses, and technical reports.

✅ Saves substantial time  
✅ Eliminates transcription errors  
✅ Produces consistent, reproducible tables  
✅ Ideal for SEM, path models, and mediation analyses  
✅ Designed for researchers who rely heavily on Mplus  

---

## 🛠 How to use (Windows – EXE)

1. Download `Mplus2Excel_Structural.exe` from **GitHub Releases**  
2. Place the EXE in the folder containing your Mplus `.out` files  
3. Double-click the EXE  
4. An Excel workbook and a `run.log` file will be created in the same folder  

If something goes wrong, the log file can be shared for troubleshooting.

---

## 🛠 How to use (Python)

```bash
pip install -r requirements.txt
python mplus2excel_structural.py

```

---

## 📂 What the Excel file contains

Instructions tab (variable and path labels)

Table 1: Structural model fit indices

Table 2+: STDYX standardized structural paths

R² table: Explained variance for endogenous variables

Audit tab: Parsed values for verification and transparency

---

## 📌 Output

By default, the workbook is saved as `Structural_Models_Report_APA.xlsx`.  
If the file is open or locked, the script saves an alternative `*_NEW.xlsx` file.


---

## 👨‍💻 **Who Maintains This?**  
This project was created by **Prof. Llewellyn E. van Zyl** .  

- 🌍 **Website:** [www.psynalytics.com](https://www.psynalytics.com)  
- 🔗 **GitHub:** [@llewellynvz](https://github.com/llewellynvz) 
- 

🔥 If you’d like to **contribute**, feel free to fork this repo, submit a pull request, or report bugs!  

---

## 📜 License

This software is proprietary. Use and local modification are permitted for
personal, academic, and internal research purposes only. Redistribution or
commercial use is not permitted.

See the LICENSE file for details.

---

## ⭐ If you found this useful

Star the repository ⭐

Share it with colleagues using Mplus

Contribute improvements or edge-case fixes
