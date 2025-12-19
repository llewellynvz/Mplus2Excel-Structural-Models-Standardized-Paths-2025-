#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
MPLUS APA STRUCTURAL RESULTS TABULATOR AND MODEL COMPARISON (v5)

What it does (double-click friendly):
- Scans the folder containing this script for all Mplus .out files.
- Builds a professional, APA-style Excel workbook:
    0) Instructions (rename factors via editable mapping table)
    1) Table 1. Structural Model Fit (fit indices + info criteria)
    2) Table 2. Standardized Structural Paths (STDYX) + R² for outcomes
    3) Reliability (AVE, alpha, omega for latent measurement factors when indicators exist)
    4) Figures (one diagram per model, embedded)

Important warnings:
- Reliability indices (alpha/omega/AVE) are computed from the model-implied covariance matrix
  reconstructed from STDYX loadings (Λ), factor correlations from standardized WITH (Φ), and θ=1−R².
  This requires (a) STDYX loadings printed, (b) factor correlations printed, (c) R² printed for indicators.
  If any of these are missing, reliability cells are left blank and a note is added.

- For ESEM / bifactor-ESEM models: omega/alpha are for unit-weighted subscale scores using the
  full ΛΦΛ' + Θ matrix as printed (including cross-loadings). This avoids the common inflation problem
  that happens when cross-loadings are ignored in the denominator.

Usage:
- Put this .py file in the folder with your Mplus .out files.
- Double click it (Windows: runs with your Python association), OR run:
    python mplus_structural_v5_professional.py
"""

from __future__ import annotations

import re
import sys
import logging
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set

import numpy as np
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
from openpyxl.utils import get_column_letter

# =============================
# Configuration
# =============================

OUTPUT_FILENAME = "Structural_Models_Report_APA.xlsx"

P_BOLD_CUT = 0.05          # bold significant betas and loadings
CI_BOLD_PCLOSE_CUT = 0.05  # bold RMSEA CI when pclose >= this
CHI_BOLD_P_CUT = 0.05      # bold chi-square when p >= this (non-significant)

# Table 2 meets criteria (measurement loadings table isn't built here, but used for optional checks)
LOADING_MIN_FOR_MEETS = 0.35
THETA_MAX_FOR_MEETS = 0.60

# Fit heuristic (Table 1 meets criteria)
CFI_TLI_STRICT = 0.90
CFI_TLI_LOOSE  = 0.85
RMSEA_MAX = 0.08
SRMR_MAX  = 0.08
PCLOSE_CUT = 0.05
REQUIRE_PCLOSE_FOR_YES = True

# Drawing settings (improves readability a lot)
DRAW_HIDE_NONSIG_LABELS = True  # hide edge labels if p >= P_BOLD_CUT
DRAW_MAX_LABELS = 40            # if more than this, hide all nonsig labels regardless
DRAW_BASE_FONTSIZE = 9
DRAW_FIG_DPI = 220


# =============================
# Logging (enterprise-grade diagnostics)
# =============================

def setup_logging(folder: Path) -> Path:
    """Configure console + file logging. Does not alter analytic functionality."""
    log_path = folder / "run.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    logging.info("Starting Mplus2Excel Structural run")
    return log_path

# =============================
# Regex + parsing helpers
# =============================

NUM_RE = re.compile(r"[-+]?(?:\d+\.\d+|\d+)(?:[EeDd][-+]?\d+)?\*?")

def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return path.read_text(encoding="latin-1", errors="ignore")

def _tofloat(x: Optional[str]) -> Optional[float]:
    if x is None:
        return None
    try:
        s = str(x).strip().replace("D", "E").replace("d", "e").rstrip("*")
        return float(s)
    except Exception:
        return None

def _toint(x: Optional[str]) -> Optional[int]:
    v = _tofloat(x)
    if v is None:
        return None
    try:
        return int(v)
    except Exception:
        return None

def _natural_key(s: str):
    nums = re.findall(r"\d+", s)
    return (int(nums[0]) if nums else 10**9, s.lower())

def _extract_after_heading(text: str, heading: str, window: int = 900) -> Optional[str]:
    idx = text.find(heading)
    if idx == -1:
        idx = text.lower().find(heading.lower())
        if idx == -1:
            return None
    return text[idx: idx + window]

def _extract_std_section(text: str) -> Optional[str]:
    """
    Return the STDYX standardized section if present, else standardized model results section.
    """
    m = re.search(r"\n\s*STDYX\s+Standardization", text, re.I)
    if not m:
        m = re.search(r"\n\s*STANDARDIZED MODEL RESULTS", text, re.I)
    if not m:
        return None
    tail = text[m.start():]
    # stop at next big heading
    stops = [
        r"\nMODEL FIT INFORMATION",
        r"\nMODEL RESULTS",
        r"\nTECHNICAL",
        r"\nQUALITY OF NUMERICAL RESULTS",
        r"\nINPUT READING TERMINATED",
        r"\nEND OF OUTPUT",
    ]
    for sp in stops:
        m2 = re.search(sp, tail, re.I)
        if m2 and m2.start() > 80:
            return tail[:m2.start()]
    return tail

def parse_usevariables(text: str) -> Set[str]:
    m = re.search(r"USEVARIABLES\s+ARE\s+(.*?);", text, re.S | re.I)
    if not m:
        return set()
    block = m.group(1)
    return set(re.findall(r"[A-Za-z0-9_]+", block))

# =============================
# Data structures
# =============================

@dataclass
class FitStats:
    chi2: Optional[float] = None
    df: Optional[int] = None
    p: Optional[float] = None
    cfi: Optional[float] = None
    tli: Optional[float] = None
    rmsea: Optional[float] = None
    rmsea_ci_lo: Optional[float] = None
    rmsea_ci_hi: Optional[float] = None
    pclose: Optional[float] = None
    srmr: Optional[float] = None
    aic: Optional[float] = None
    bic: Optional[float] = None
    abic: Optional[float] = None
    terminated: bool = True

@dataclass
class Loading:
    est: Optional[float] = None
    se: Optional[float] = None
    p: Optional[float] = None

@dataclass
class PathCoef:
    dv: str
    iv: str
    est: Optional[float]
    se: Optional[float]
    p: Optional[float]

    @property
    def sig(self) -> str:
        """Significance stars based on p-value (APA style)."""
        if self.p is None:
            return ""
        try:
            if self.p < 0.001:
                return "***"
            if self.p < 0.01:
                return "**"
            if self.p < 0.05:
                return "*"
        except Exception:
            return ""
        return ""


@dataclass
class ParsedModel:
    out_file: Path
    model_label: str
    model_type: str
    fit: FitStats
    r2: Dict[str, float]                     # for observed vars, and for indicators when present
    paths: List[PathCoef]                    # standardized ON paths
    loadings: Dict[str, Dict[str, Loading]]  # factor -> item -> Loading (observed indicators only)
    phi: Dict[Tuple[str, str], float]        # factor correlations (STDYX WITH)
    usevars: Set[str]

# =============================
# Fit parsing
# =============================

def parse_fit(text: str) -> FitStats:
    fs = FitStats()
    fs.terminated = ("THE MODEL ESTIMATION TERMINATED NORMALLY" in text.upper())

    chi_block = _extract_after_heading(text, "Chi-Square Test of Model Fit", 700)
    if chi_block:
        m = re.search(r"Value\s+(" + NUM_RE.pattern + r")", chi_block, re.I)
        if m: fs.chi2 = _tofloat(m.group(1))
        m = re.search(r"Degrees of Freedom\s+(" + NUM_RE.pattern + r")", chi_block, re.I)
        if m: fs.df = _toint(m.group(1))
        m = re.search(r"P-Value\s+(" + NUM_RE.pattern + r")", chi_block, re.I)
        if m: fs.p = _tofloat(m.group(1))

    rm_block = _extract_after_heading(text, "RMSEA (Root Mean Square Error Of Approximation)", 900)
    if rm_block:
        m = re.search(r"Estimate\s+(" + NUM_RE.pattern + r")", rm_block, re.I)
        if m: fs.rmsea = _tofloat(m.group(1))
        m = re.search(r"90\s*Percent\s*C\.I\.\s+(" + NUM_RE.pattern + r")\s+(" + NUM_RE.pattern + r")", rm_block, re.I)
        if m:
            fs.rmsea_ci_lo = _tofloat(m.group(1))
            fs.rmsea_ci_hi = _tofloat(m.group(2))
        m = re.search(r"Probability\s+RMSEA\s*<=\s*\.?0?5\s+(" + NUM_RE.pattern + r")", rm_block, re.I)
        if m:
            fs.pclose = _tofloat(m.group(1))

    cfi_block = _extract_after_heading(text, "CFI/TLI", 600)
    if cfi_block:
        m = re.search(r"\n\s*CFI\s+(" + NUM_RE.pattern + r")", cfi_block, re.I)
        if m: fs.cfi = _tofloat(m.group(1))
        m = re.search(r"\n\s*TLI\s+(" + NUM_RE.pattern + r")", cfi_block, re.I)
        if m: fs.tli = _tofloat(m.group(1))
    else:
        m = re.search(r"\bCFI\s+(" + NUM_RE.pattern + r")", text, re.I)
        if m: fs.cfi = _tofloat(m.group(1))
        m = re.search(r"\bTLI\s+(" + NUM_RE.pattern + r")", text, re.I)
        if m: fs.tli = _tofloat(m.group(1))

    srmr_block = _extract_after_heading(text, "Standardized Root Mean Square Residual", 600)
    if srmr_block:
        m = re.search(r"Value\s+(" + NUM_RE.pattern + r")", srmr_block, re.I)
        if m: fs.srmr = _tofloat(m.group(1))
    if fs.srmr is None:
        m = re.search(r"\bSRMR\s+(" + NUM_RE.pattern + r")", text, re.I)
        if m: fs.srmr = _tofloat(m.group(1))

    ic_block = _extract_after_heading(text, "Information Criteria", 1200)
    if ic_block:
        m = re.search(r"Akaike\s*\(AIC\)\s+(" + NUM_RE.pattern + r")", ic_block, re.I)
        if m: fs.aic = _tofloat(m.group(1))
        m = re.search(r"Bayesian\s*\(BIC\)\s+(" + NUM_RE.pattern + r")", ic_block, re.I)
        if m: fs.bic = _tofloat(m.group(1))
        m = re.search(r"Sample-Size\s+Adjusted\s+BIC\s+(" + NUM_RE.pattern + r")", ic_block, re.I)
        if m: fs.abic = _tofloat(m.group(1))

    return fs

def classify_meets_criteria(fs: FitStats) -> str:
    cfi, tli, rmsea, srmr = fs.cfi, fs.tli, fs.rmsea, fs.srmr
    hi = fs.rmsea_ci_hi
    if any(v is None for v in [cfi, tli, rmsea, srmr]):
        return "No"
    strict = (cfi >= CFI_TLI_STRICT and tli >= CFI_TLI_STRICT and rmsea <= RMSEA_MAX and srmr <= SRMR_MAX)
    if hi is not None:
        strict = strict and (hi <= RMSEA_MAX)
    pclose_ok = (fs.pclose is None) or (fs.pclose >= PCLOSE_CUT)
    if strict and ((not REQUIRE_PCLOSE_FOR_YES) or pclose_ok):
        return "Yes"
    # marginal logic
    if (cfi >= CFI_TLI_LOOSE and tli >= CFI_TLI_LOOSE):
        return "Marginally"
    return "No"

# =============================
# R-square parsing
# =============================

def parse_rsquare(text: str) -> Dict[str, float]:
    r2: Dict[str, float] = {}
    m = re.search(r"\nR-SQUARE\s*\n", text, re.I)
    if not m:
        return r2
    tail = text[m.end():]
    stop = re.search(r"\n\s*\n\s*\n|\nMODEL FIT INFORMATION|\nMODEL RESULTS|\nSTANDARDIZED MODEL RESULTS|\nTECHNICAL", tail, re.I)
    sec = tail if stop is None else tail[:stop.start()]
    for ln in sec.splitlines():
        parts = ln.strip().split()
        if len(parts) >= 2 and re.fullmatch(r"[A-Za-z0-9_]+", parts[0]) and NUM_RE.fullmatch(parts[1]):
            r2[parts[0].strip().strip(",")] = float(parts[1].rstrip("*").replace("D","E").replace("d","e"))
    return r2

# =============================
# STDYX parsing (paths, loadings, correlations)
# =============================

def _nums_from_parts(parts: List[str]) -> List[str]:
    return [p.rstrip("*") for p in parts if NUM_RE.fullmatch(p)]

def parse_stdyx_on_paths(text: str) -> List[PathCoef]:
    """
    Parse standardized structural coefficients from STDYX section, for DV ON IV blocks.
    """
    out: List[PathCoef] = []
    sec = _extract_std_section(text)
    if not sec:
        return out

    lines = sec.splitlines()
    current_dv: Optional[str] = None
    on_header = re.compile(r"^\s*([A-Za-z][A-Za-z0-9_]*)\s+ON\s*$", re.I)

    for ln in lines:
        m = on_header.match(ln)
        if m:
            current_dv = m.group(1)
            continue

        if current_dv is None:
            continue

        # blank line ends the block
        if ln.strip() == "":
            current_dv = None
            continue

        parts = ln.strip().split()
        if len(parts) < 3:
            continue
        iv = parts[0]
        if not re.fullmatch(r"[A-Za-z0-9_]+", iv):
            continue
        nums = _nums_from_parts(parts[1:])
        # Estimate, S.E., Est/SE, P-value
        if len(nums) >= 2:
            est = _tofloat(nums[0])
            se = _tofloat(nums[1])
            p = _tofloat(nums[3]) if len(nums) >= 4 else None
            out.append(PathCoef(dv=current_dv, iv=iv, est=est, se=se, p=p))

    return out

def parse_stdyx_loadings(text: str, usevars: Set[str]) -> Dict[str, Dict[str, Loading]]:
    """
    Parse standardized factor loadings from STDYX BY blocks.
    Keeps only *observed* indicators that appear in USEVARIABLES (best-effort filter).
    """
    loadings: Dict[str, Dict[str, Loading]] = {}
    sec = _extract_std_section(text)
    if not sec:
        return loadings

    lines = sec.splitlines()
    current_factor: Optional[str] = None
    by_header = re.compile(r"^\s*([A-Za-z][A-Za-z0-9_]*)\s+BY\s*$", re.I)

    for ln in lines:
        m = by_header.match(ln)
        if m:
            current_factor = m.group(1)
            loadings.setdefault(current_factor, {})
            continue

        if current_factor is None:
            continue

        if ln.strip() == "":
            current_factor = None
            continue

        parts = ln.strip().split()
        if len(parts) < 3:
            continue
        ind = parts[0]
        if not re.fullmatch(r"[A-Za-z0-9_]+", ind):
            continue
        # observed indicator filter
        if usevars and (ind not in usevars):
            continue

        nums = _nums_from_parts(parts[1:])
        if len(nums) >= 2:
            est = _tofloat(nums[0])
            se = _tofloat(nums[1])
            p = _tofloat(nums[3]) if len(nums) >= 4 else None
            loadings[current_factor][ind] = Loading(est=est, se=se, p=p)

    # remove empty factors
    return {f: d for f, d in loadings.items() if d}

def parse_stdyx_factor_correlations(text: str) -> Dict[Tuple[str, str], float]:
    """
    Parse factor correlations from standardized WITH blocks.
    """
    phi: Dict[Tuple[str, str], float] = {}
    sec = _extract_std_section(text)
    if not sec:
        return phi

    current: Optional[str] = None
    with_header = re.compile(r"^\s*([A-Za-z][A-Za-z0-9_]*)\s+WITH\s*$", re.I)

    for ln in sec.splitlines():
        m = with_header.match(ln)
        if m:
            current = m.group(1)
            continue
        if current is None:
            continue
        if ln.strip() == "":
            current = None
            continue
        parts = ln.strip().split()
        if len(parts) < 2:
            continue
        other = parts[0]
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", other):
            continue
        nums = _nums_from_parts(parts[1:])
        if len(nums) >= 1:
            est = _tofloat(nums[0])
            if est is None or current == other:
                continue
            phi[(current, other)] = float(est)
            phi[(other, current)] = float(est)
    return phi

# =============================
# Reliability indices (Λ, Φ, Θ reconstruction)
# =============================

def _build_lambda_phi_theta(loadings: Dict[str, Dict[str, Loading]],
                            phi: Dict[Tuple[str, str], float],
                            r2: Dict[str, float]) -> Tuple[List[str], List[str], np.ndarray, np.ndarray, np.ndarray]:
    factors = list(loadings.keys())
    items = sorted({it for d in loadings.values() for it in d.keys()}, key=_natural_key)
    q = len(factors)
    p = len(items)

    idx_f = {f:i for i,f in enumerate(factors)}
    idx_i = {it:i for i,it in enumerate(items)}

    L = np.zeros((p, q), dtype=float)
    for f, d in loadings.items():
        j = idx_f[f]
        for it, ld in d.items():
            i = idx_i[it]
            if ld.est is not None:
                L[i, j] = float(ld.est)

    Phi = np.eye(q, dtype=float)
    for (a, b), v in phi.items():
        if a in idx_f and b in idx_f and a != b and v is not None:
            Phi[idx_f[a], idx_f[b]] = float(v)

    Theta = np.full((p,), np.nan, dtype=float)
    for it in items:
        if it in r2:
            Theta[idx_i[it]] = max(0.0, 1.0 - float(r2[it]))
    return items, factors, L, Phi, Theta

def _model_implied_cov(L: np.ndarray, Phi: np.ndarray, Theta: np.ndarray) -> Optional[np.ndarray]:
    if np.any(np.isnan(Theta)):
        return None
    return (L @ Phi @ L.T) + np.diag(Theta)

def _alpha_from_cov(S: np.ndarray) -> Optional[float]:
    k = S.shape[0]
    if k < 2:
        return None
    tr = float(np.trace(S))
    tot = float(S.sum())
    if tot <= 0:
        return None
    return (k / (k - 1.0)) * (1.0 - (tr / tot))

def _omega_total(S: np.ndarray, S_common: np.ndarray, w: np.ndarray) -> Optional[float]:
    tot = float(w.T @ S @ w)
    if tot <= 0:
        return None
    com = float(w.T @ S_common @ w)
    return com / tot

def _infer_general_factor(loadings: Dict[str, Dict[str, Loading]]) -> Optional[str]:
    if not loadings:
        return None
    # explicit names
    for fn in loadings.keys():
        up = fn.upper()
        if up in {"G", "GF", "GFACTOR", "GENERAL"} or "GENERAL" in up:
            return fn
    # count heuristic
    items = {it for d in loadings.values() for it in d.keys()}
    if not items:
        return None
    best = max(loadings.keys(), key=lambda k: len(loadings.get(k, {})))
    if len(loadings.get(best, {})) >= int(0.90 * len(items)):
        return best
    return None

def _is_bifactor_like(model_type: str, filename: str) -> bool:
    s = (model_type + " " + filename).lower()
    return ("bifactor" in s) or ("bi-factor" in s) or ("bif " in s)

def _assign_primary_factor(loadings: Dict[str, Dict[str, Loading]], exclude_factor: Optional[str]=None) -> Dict[str, str]:
    best: Dict[str, Tuple[str, float]] = {}
    for f, items in loadings.items():
        if exclude_factor and f == exclude_factor:
            continue
        for it, ld in items.items():
            if ld.est is None:
                continue
            val = abs(ld.est)
            if it not in best or val > best[it][1]:
                best[it] = (f, val)
    return {it: f for it, (f, _) in best.items()}

def compute_factor_indices(model: ParsedModel) -> Dict[str, Dict[str, Optional[float]]]:
    """
    Returns factor -> dict(k, AVE, alpha, omega, omegaH, omegaS)

    Notes:
    - AVE uses Fornell-Larcker: sum(lambda^2) / (sum(lambda^2)+sum(theta)).
      (This is the convergent validity AVE, not the "mean of squared loadings" variant.)
    - alpha/omega are computed for unit-weighted scores using Σ = ΛΦΛ' + Θ.
    - For bifactor-like models (detected heuristically), omegaH and omegaS are computed as:
        omegaH: proportion of score variance attributable to general factor only
        omegaS: proportion of score variance attributable to that specific factor only
      using the corresponding single-factor contribution to Σ.
    """
    out: Dict[str, Dict[str, Optional[float]]] = {}
    if not model.loadings:
        return out

    items, factors, L, Phi, Theta = _build_lambda_phi_theta(model.loadings, model.phi, model.r2)
    S = _model_implied_cov(L, Phi, Theta)
    if S is None:
        return out

    idx_f = {f:i for i,f in enumerate(factors)}
    idx_i = {it:i for i,it in enumerate(items)}

    bif = _is_bifactor_like(model.model_type, model.out_file.stem)
    g_name = _infer_general_factor(model.loadings) if bif else None
    primary = _assign_primary_factor(model.loadings, exclude_factor=g_name if bif else None)

    # common covariance for all factors
    S_common_all = (L @ Phi @ L.T)

    for f in factors:
        if bif and g_name and f == g_name:
            continue
        its = sorted([it for it, pf in primary.items() if pf == f], key=_natural_key)
        if len(its) == 0:
            continue
        rows = [idx_i[it] for it in its]
        S_sub = S[np.ix_(rows, rows)]
        S_common_sub = S_common_all[np.ix_(rows, rows)]
        w = np.ones((len(rows),), dtype=float)

        alpha = _alpha_from_cov(S_sub)
        omega = _omega_total(S_sub, S_common_sub, w)

        lam = L[rows, idx_f[f]]
        the = Theta[rows]
        ave = None
        if not np.any(np.isnan(the)):
            num = float(np.sum(lam * lam))
            den = num + float(np.sum(the))
            ave = (num / den) if den > 0 else None

        omegaH = None
        omegaS = None
        if bif and g_name and g_name in idx_f:
            jg = idx_f[g_name]
            js = idx_f[f]
            # general-only contribution (assume orthogonal; Phi may have ~0 off-diagonals in bifactor)
            S_g = np.outer(L[rows, jg], L[rows, jg])  # Var(g)=1 in STDYX
            omegaH = _omega_total(S_sub, S_g, w)
            S_s = np.outer(L[rows, js], L[rows, js])
            omegaS = _omega_total(S_sub, S_s, w)

        out[f] = {"k": float(len(its)), "AVE": ave, "alpha": alpha, "omega": omega, "omegaH": omegaH, "omegaS": omegaS}

    return out

# =============================
# Model parsing (single .out)
# =============================

def _infer_model_label_and_type(fp: Path, idx: int) -> Tuple[str, str]:
    # Model label: "Model 1", "Model 2", ... by file order
    label = f"Model {idx}"
    typ = fp.stem
    # strip common prefixes
    typ = re.sub(r"^\s*SM\s*\d+\s*[-–]\s*", "", typ, flags=re.I).strip()
    return label, typ

def parse_out_file(fp: Path, idx: int) -> ParsedModel:
    text = _read_text(fp)
    usevars = parse_usevariables(text)

    model_label, model_type = _infer_model_label_and_type(fp, idx)

    fit = parse_fit(text)
    r2 = parse_rsquare(text)
    paths = parse_stdyx_on_paths(text)
    phi = parse_stdyx_factor_correlations(text)
    loadings = parse_stdyx_loadings(text, usevars=usevars)

    return ParsedModel(
        out_file=fp,
        model_label=model_label,
        model_type=model_type,
        fit=fit,
        r2=r2,
        paths=paths,
        loadings=loadings,
        phi=phi,
        usevars=usevars,
    )

# =============================
# Excel styling helpers
# =============================

FONT_BODY  = Font(name="Times New Roman", size=10)
FONT_BOLD  = Font(name="Times New Roman", size=10, bold=True)
FONT_TITLE = Font(name="Times New Roman", size=12, bold=True)
FONT_NOTE  = Font(name="Times New Roman", size=9, italic=False)

ALIGN_LEFT   = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT  = Alignment(horizontal="right", vertical="center", wrap_text=True)

LINE = Side(style="thin", color="000000")

def border_top(): return Border(top=LINE)
def border_bottom(): return Border(bottom=LINE)
def border_top_bottom(): return Border(top=LINE, bottom=LINE)

def _set(ws, r, c, val, font=None, align=None, border: Optional[Border]=None, num_format: Optional[str]=None, fill: Optional[PatternFill]=None):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = font if font is not None else FONT_BODY
    cell.alignment = align if align is not None else ALIGN_CENTER
    if border is not None:
        cell.border = border
    if num_format is not None and val is not None:
        cell.number_format = num_format
    if fill is not None:
        cell.fill = fill
    return cell

def _apply_border_row(ws, r, c1, c2, border: Border):
    for c in range(c1, c2 + 1):
        ws.cell(r, c).border = border

def _set_col_widths(ws, widths: Dict[int, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

def apply_global_format(ws):
    max_r = ws.max_row or 1
    max_c = ws.max_column or 1
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            cell = ws.cell(r, c)
            # enforce font family/size, keep bold/italic
            f = cell.font or FONT_BODY
            cell.font = Font(
                name="Times New Roman",
                size=f.size or 10,
                bold=f.bold,
                italic=f.italic,
                underline=f.underline,
                color=f.color
            )
            if cell.alignment is None:
                cell.alignment = ALIGN_CENTER

# =============================
# Sheet 0: Instructions
# =============================

def build_instructions(ws, factor_names: List[str]):
    ws.title = "Instructions"
    _set(ws, 1, 1, "Instructions", font=FONT_TITLE, align=ALIGN_LEFT)
    _set(ws, 3, 1, "Factor label mapping (edit Column B to rename factors across the workbook):", font=FONT_BOLD, align=ALIGN_LEFT)

    _set(ws, 5, 1, "Original", font=FONT_BOLD)
    _set(ws, 5, 2, "Display label", font=FONT_BOLD)
    _apply_border_row(ws, 5, 1, 2, border_top_bottom())

    r = 6
    for f in factor_names:
        _set(ws, r, 1, f, align=ALIGN_LEFT)
        _set(ws, r, 2, f, align=ALIGN_LEFT)
        r += 1
    _apply_border_row(ws, r - 1, 1, 2, border_bottom())

    _set(ws, 3, 4, "Notes", font=FONT_BOLD, align=ALIGN_LEFT)
    note = ("Rename any factor by editing the Display label (Column B). "
            "The workbook uses Excel lookups to display the renamed labels.\n\n"
            "Reliability indices (alpha/omega/AVE) are computed from STDYX loadings, "
            "standardized factor correlations (WITH), and θ=1−R². If any component is missing, "
            "reliability values are left blank.")
    _set(ws, 4, 4, note, font=FONT_BODY, align=ALIGN_LEFT)
    ws.merge_cells(start_row=4, start_column=4, end_row=10, end_column=9)

    _set_col_widths(ws, {1: 22, 2: 34, 4: 18, 5: 18, 6: 18, 7: 18, 8: 18, 9: 18})
    ws.freeze_panes = "A6"

def _factor_display_formula(factor: str, map_range: str = "Instructions!$A$6:$B$300") -> str:
    return f'=IFERROR(VLOOKUP("{factor}",{map_range},2,FALSE),"{factor}")'

# =============================
# Sheet 1: Table 1 (Fit)
# =============================


def build_table1(ws, models: List[ParsedModel]) -> None:
    """Table 1. Structural Models (APA-style). Cosmetic formatting only."""
    ws.title = "Table 1. Structural Models"
    ws.delete_rows(1, ws.max_row)

    headers = [
        "Model", "Type", "χ²", "df", "p(χ²)", "CFI", "TLI",
        "RMSEA", "90% CI", "pclose", "SRMR", "AIC", "BIC", "aBIC",
        "Meets criteria", "Notes", "Filename"
    ]

    # Title (single merged row)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = ws.cell(1, 1, "Table 1. Structural models")
    # Bold only “Table 1.”
    title_cell.font = Font(name="Times New Roman", size=12, bold=True)
    title_cell.alignment = ALIGN_LEFT

    # Headers on row 2
    for i, h in enumerate(headers, start=1):
        _set(ws, 2, i, h, font=FONT_BOLD, align=ALIGN_CENTER)
    _apply_border_row(ws, 2, 1, len(headers), border_top_bottom())
    ws.freeze_panes = "A3"

    widths = {
        1: 10, 2: 30, 3: 10, 4: 6, 5: 10, 6: 7, 7: 7, 8: 8,
        9: 18, 10: 8, 11: 8, 12: 12, 13: 12, 14: 12, 15: 14,
        16: 32, 17: 44
    }
    _set_col_widths(ws, widths)

    r = 3
    for pm in models:
        fs = pm.fit
        meets = classify_meets_criteria(fs)

        notes = []
        if not fs.terminated:
            notes.append("Did not terminate normally")
        note = "; ".join(notes) if notes else ""

        ci = ""
        if fs.rmsea_ci_lo is not None and fs.rmsea_ci_hi is not None:
            # requested spacing: [0.123 - 0.246]
            ci = f"[{fs.rmsea_ci_lo:.3f} - {fs.rmsea_ci_hi:.3f}]"

        _set(ws, r, 1, pm.model_label, align=ALIGN_LEFT)
        _set(ws, r, 2, pm.model_type, align=ALIGN_LEFT)

        c_chi = _set(ws, r, 3, fs.chi2, align=ALIGN_RIGHT, num_format="0.000")
        _set(ws, r, 4, fs.df, align=ALIGN_RIGHT, num_format="0")
        c_pchi = _set(ws, r, 5, fs.p, align=ALIGN_RIGHT, num_format="0.0000")
        _set(ws, r, 6, fs.cfi, align=ALIGN_RIGHT, num_format="0.000")
        _set(ws, r, 7, fs.tli, align=ALIGN_RIGHT, num_format="0.000")

        c_rm = _set(ws, r, 8, fs.rmsea, align=ALIGN_RIGHT, num_format="0.000")
        _set(ws, r, 9, ci, align=ALIGN_CENTER)
        c_pclose = _set(ws, r, 10, fs.pclose, align=ALIGN_RIGHT, num_format="0.0000")
        _set(ws, r, 11, fs.srmr, align=ALIGN_RIGHT, num_format="0.000")
        _set(ws, r, 12, fs.aic, align=ALIGN_RIGHT, num_format="0.000")
        _set(ws, r, 13, fs.bic, align=ALIGN_RIGHT, num_format="0.000")
        _set(ws, r, 14, fs.abic, align=ALIGN_RIGHT, num_format="0.000")
        _set(ws, r, 15, meets, align=ALIGN_CENTER)
        _set(ws, r, 16, note, align=ALIGN_LEFT)
        _set(ws, r, 17, pm.out_file.name, align=ALIGN_LEFT)

        # Cosmetic emphasis:
        # - Bold χ² value if non-significant
        if fs.p is not None and fs.p >= 0.05 and fs.chi2 is not None:
            c_chi.font = FONT_BOLD
        # - Bold RMSEA value if pclose is non-significant (>= .05)
        if fs.pclose is not None and fs.pclose >= 0.05 and fs.rmsea is not None:
            c_rm.font = FONT_BOLD

        r += 1

    if r > 3:
        _apply_border_row(ws, r - 1, 1, len(headers), border_bottom())

    # Footnote
    note_r = r + 1
    ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=len(headers))
    note = (
        "Note. χ² = chi-square; df = degrees of freedom; RMSEA = root mean square error of approximation; "
        "CI = confidence interval; SRMR = standardized root mean square residual; aBIC = sample-size adjusted BIC. "
        "Bold χ² indicates p ≥ .05; bold RMSEA indicates pclose ≥ .05."
    )
    _set(ws, note_r, 1, note, font=FONT_NOTE, align=ALIGN_LEFT)



def build_table2(ws, models: List[ParsedModel]) -> None:
    """Creates per-model APA tables of standardized structural paths (β, SE, p)."""
    ws.title = "Std Structural Paths"
    ws.delete_rows(1, ws.max_row)

    columns = ["From", "To", "β (STDYX)", "SE", "p", "Sig.", "R² (To)"]
    widths = {1: 18, 2: 18, 3: 10, 4: 8, 5: 10, 6: 6}
    _set_col_widths(ws, widths)

    r = 1
    table_no = 2  # Table 2..n
    for pm in models:
        # Skip if no paths
        if not pm.paths:
            continue

        # Single-line title row (merged)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(columns))
        title_txt = f"Table {table_no}. Standardized structural paths for {pm.model_label} ({pm.model_type})"
        title_cell = ws.cell(r, 1, title_txt)
        title_cell.font = Font(name="Times New Roman", size=12, bold=True)
        title_cell.alignment = ALIGN_LEFT
        r += 1

        # Header row
        for j, h in enumerate(columns, start=1):
            _set(ws, r, j, h, font=FONT_BOLD, align=ALIGN_CENTER)
        _apply_border_row(ws, r, 1, len(columns), border_top_bottom())
        r += 1

        # Body
        for pth in pm.paths:
            _set(ws, r, 1, pth.iv, align=ALIGN_LEFT)
            _set(ws, r, 2, pth.dv, align=ALIGN_LEFT)

            c_beta = _set(ws, r, 3, pth.est, align=ALIGN_RIGHT, num_format="0.000")
            _set(ws, r, 4, pth.se, align=ALIGN_RIGHT, num_format="0.000")
            _set(ws, r, 5, pth.p, align=ALIGN_RIGHT, num_format="0.0000")

            sig = pth.sig or ""
            c_sig = _set(ws, r, 6, sig, align=ALIGN_CENTER)

            r2_to = pm.r2.get(pth.dv)
            _set(ws, r, 7, r2_to, align=ALIGN_RIGHT, num_format="0.0000")

            # Bold significant β only (not SE)
            if pth.p is not None and pth.p < P_BOLD_CUT and pth.est is not None:
                c_beta.font = FONT_BOLD

            r += 1

        _apply_border_row(ws, r - 1, 1, len(columns), border_bottom())

        # Model spacer
        r += 2
        table_no += 1

    ws.freeze_panes = "A1"


def build_reliability(ws, models: List[ParsedModel]):
    ws.title = "Reliability"
    ws.freeze_panes = "A5"

    _set(ws, 1, 1, "Reliability (from STDYX loadings, WITH correlations, and θ=1−R²)", font=FONT_TITLE, align=ALIGN_LEFT)

    headers = ["Model", "Factor", "k", "AVE", "α", "ω", "ωH", "ωS", "Notes"]
    for c, h in enumerate(headers, start=1):
        _set(ws, 4, c, h, font=FONT_BOLD, align=ALIGN_CENTER)
    _apply_border_row(ws, 4, 1, len(headers), border_top_bottom())

    _set_col_widths(ws, {1: 10, 2: 22, 3: 5, 4: 8, 5: 8, 6: 8, 7: 8, 8: 8, 9: 40})

    r = 5
    for pm in models:
        idxs = compute_factor_indices(pm)
        if not idxs:
            _set(ws, r, 1, pm.model_label, align=ALIGN_LEFT)
            _set(ws, r, 2, "", align=ALIGN_LEFT)
            _set(ws, r, 9, "Reliability not computed (missing loadings and/or R² and/or factor correlations).", align=ALIGN_LEFT)
            r += 1
            continue

        for f, vals in sorted(idxs.items(), key=lambda kv: kv[0].lower()):
            _set(ws, r, 1, pm.model_label, align=ALIGN_LEFT)
            _set(ws, r, 2, _factor_display_formula(f), align=ALIGN_LEFT)
            _set(ws, r, 3, vals.get("k"), align=ALIGN_RIGHT, num_format="0")
            _set(ws, r, 4, vals.get("AVE"), align=ALIGN_RIGHT, num_format="0.000")
            _set(ws, r, 5, vals.get("alpha"), align=ALIGN_RIGHT, num_format="0.000")
            _set(ws, r, 6, vals.get("omega"), align=ALIGN_RIGHT, num_format="0.000")
            _set(ws, r, 7, vals.get("omegaH"), align=ALIGN_RIGHT, num_format="0.000")
            _set(ws, r, 8, vals.get("omegaS"), align=ALIGN_RIGHT, num_format="0.000")
            note = ""
            if any(v is None for v in [vals.get("alpha"), vals.get("omega")]):
                note = "Missing θ for one or more indicators (R² not printed for all items)."
            _set(ws, r, 9, note, align=ALIGN_LEFT)
            r += 1

        r += 1

    if r > 5:
        _apply_border_row(ws, r - 1, 1, len(headers), border_bottom())

    note_r = r + 1
    ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=len(headers))
    note = ("Note. AVE uses Fornell-Larcker: Σλ² / (Σλ² + Σθ). Alpha and omega are computed for unit-weighted subscale scores "
            "using Σ = ΛΦΛ' + Θ, where Λ and Φ are taken from STDYX output and θ=1−R² from the R-SQUARE section.")
    _set(ws, note_r, 1, note, font=FONT_NOTE, align=ALIGN_LEFT)

# =============================
# Figures (matplotlib)
# =============================

def _sig(p: Optional[float]) -> bool:
    return (p is not None) and (p < P_BOLD_CUT)


def draw_sem_figure(pm: ParsedModel, out_png: Path) -> Optional[Path]:
    """
    Cleaner, more publishable SEM-style diagram using matplotlib.

    Design goals (cosmetic only):
    - larger canvas and more vertical spacing to prevent label collisions
    - ellipses for nodes (APA-ish SEM look)
    - curved/staggered arrows per DV to separate path labels
    - label only significant paths by default (or when the model is small)
    """
    try:
        import matplotlib.pyplot as plt
        from matplotlib.patches import Ellipse, FancyArrowPatch
    except Exception:
        return None

    if not pm.paths:
        return None

    dvs = sorted({p.dv for p in pm.paths}, key=str.lower)
    ivs = sorted({p.iv for p in pm.paths}, key=str.lower)

    # Bigger base sizes than v5, then scale modestly with node counts
    fig_w = 13 + 0.6 * max(0, len(dvs) - 3)
    fig_h = 8 + 0.7 * max(len(ivs), len(dvs)) / 2.0
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=DRAW_FIG_DPI)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()

    # Normalized coordinate system
    left_x, right_x = 0.12, 0.88
    top_y, bot_y = 0.94, 0.06

    def spaced_positions(n: int) -> List[float]:
        if n <= 1:
            return [(top_y + bot_y) / 2.0]
        return [top_y - i * (top_y - bot_y) / (n - 1) for i in range(n)]

    iv_y = spaced_positions(len(ivs))
    dv_y = spaced_positions(len(dvs))

    # Node geometry in axes coordinates
    node_w = 0.22
    node_h = 0.08
    text_fs = max(8, DRAW_BASE_FONTSIZE - int(0.12 * (len(ivs) + len(dvs))))

    # Draw nodes (ellipses)
    iv_pos: Dict[str, Tuple[float, float]] = {}
    for name, y in zip(ivs, iv_y):
        iv_pos[name] = (left_x, y)
        e = Ellipse((left_x, y), width=node_w, height=node_h, facecolor="white", edgecolor="black", lw=1.4)
        ax.add_patch(e)
        ax.text(left_x, y, _display_name(name), ha="center", va="center", fontsize=text_fs, fontweight="bold")

    dv_pos: Dict[str, Tuple[float, float]] = {}
    for name, y in zip(dvs, dv_y):
        dv_pos[name] = (right_x, y)
        e = Ellipse((right_x, y), width=node_w, height=node_h, facecolor="white", edgecolor="black", lw=1.4)
        ax.add_patch(e)
        r2 = pm.r2.get(name)
        label = _display_name(name)
        if r2 is not None:
            label = f"{label}\nR²={r2:.2f}"
        ax.text(right_x, y, label, ha="center", va="center", fontsize=text_fs, fontweight="bold")

    # Edge label policy
    total_edges = len(pm.paths)
    hide_nonsig = DRAW_HIDE_NONSIG_LABELS or (total_edges > DRAW_MAX_LABELS)

    # Group edges by DV, keep stable order
    by_dv: Dict[str, List[PathCoef]] = {}
    for pth in pm.paths:
        by_dv.setdefault(pth.dv, []).append(pth)
    for dv in by_dv:
        by_dv[dv].sort(key=lambda p: (p.iv.lower(), p.dv.lower()))

    # Helper: start/end points offset by ellipse radius
    x1_start = left_x + node_w / 2.0 - 0.01
    x2_end = right_x - node_w / 2.0 + 0.01

    for dv in dvs:
        plist = by_dv.get(dv, [])
        m = len(plist)
        for k, pth in enumerate(plist):
            _, y1 = iv_pos.get(pth.iv, (left_x, 0.5))
            _, y2 = dv_pos.get(dv, (right_x, 0.5))

            is_sig = _sig(pth.p)
            color = "black" if is_sig else "0.75"
            lw = 1.8 if is_sig else 1.0
            ls = "-" if is_sig else "--"

            # Curvature per edge to separate lines and labels
            # small positive/negative radii around 0
            if m <= 1:
                rad = 0.0
            else:
                span = (k - (m - 1) / 2.0)
                rad = float(span) * 0.08  # stronger separation than v5

            arrow = FancyArrowPatch(
                (x1_start, y1),
                (x2_end, y2),
                arrowstyle="-|>",
                mutation_scale=14,
                lw=lw,
                linestyle=ls,
                color=color,
                connectionstyle=f"arc3,rad={rad}",
                shrinkA=0,
                shrinkB=0,
            )
            ax.add_patch(arrow)

            # Labels (significant by default, or all when small)
            if (not hide_nonsig) or is_sig:
                if pth.est is not None and pth.se is not None:
                    lab = f"β={pth.est:.2f}, SE={pth.se:.2f}"
                elif pth.est is not None:
                    lab = f"β={pth.est:.2f}"
                else:
                    lab = ""
                if lab:
                    # place label at mid-point with a small vertical nudge aligned to curvature
                    y_mid = (y1 + y2) / 2.0 + (0.02 * (1 if rad >= 0 else -1))
                    x_mid = (x1_start + x2_end) / 2.0
                    ax.text(
                        x_mid,
                        y_mid,
                        lab,
                        fontsize=max(7, text_fs - 1),
                        ha="center",
                        va="center",
                        color=color,
                        bbox=dict(boxstyle="round,pad=0.20", fc="white", ec="none", alpha=0.90),
                    )

    title = f"{pm.model_label}: {pm.model_type}"
    ax.text(0.5, 0.99, title, ha="center", va="top", fontsize=text_fs + 3, fontweight="bold")

    fig.savefig(out_png, dpi=DRAW_FIG_DPI, bbox_inches="tight")
    plt.close(fig)
    return out_png


def _display_name(name: str) -> str:
    # will be displayed via mapping in Excel; figure uses raw name
    return name

def build_figures(ws, models: List[ParsedModel], folder: Path):
    ws.title = "Figures"
    _set(ws, 1, 1, "Figures (auto-generated)", font=FONT_TITLE, align=ALIGN_LEFT)
    r = 3
    for pm in models:
        _set(ws, r, 1, f"{pm.model_label}: {pm.model_type}", font=FONT_BOLD, align=ALIGN_LEFT)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        r += 1

        png = folder / f"{pm.model_label.replace(' ', '_')}_sem.png"
        out = draw_sem_figure(pm, png)
        if out and out.exists():
            try:
                img = XLImage(str(out))
                img.width = int(img.width * 0.85)
                img.height = int(img.height * 0.85)
                ws.add_image(img, f"A{r}")
                r += int(max(18, img.height / 18))  # crude row jump
            except Exception:
                _set(ws, r, 1, "Figure could not be embedded (image created on disk).", align=ALIGN_LEFT)
                r += 2
        else:
            _set(ws, r, 1, "No figure generated (no standardized paths found).", align=ALIGN_LEFT)
            r += 2

        r += 2

    _set_col_widths(ws, {1: 18, 2: 18, 3: 18, 4: 18, 5: 18, 6: 18, 7: 18, 8: 18})

# =============================
# Main
# =============================

def collect_factor_names(models: List[ParsedModel]) -> List[str]:
    names: Set[str] = set()
    for m in models:
        # from structural vars
        for p in m.paths:
            names.add(p.iv)
            names.add(p.dv)
        # from measurement factors
        for f in m.loadings.keys():
            names.add(f)
    return sorted(names, key=lambda x: x.lower())

def main() -> None:
    if getattr(sys, "frozen", False):
    # Running as a PyInstaller EXE: use the folder where the EXE lives
        folder = Path(sys.executable).resolve().parent
    else:
        # Running as a normal .py file: use the folder where the script lives
        folder = Path(__file__).resolve().parent
    setup_logging(folder)
    logging.info('Input folder: %s', folder)
    out_files = sorted(folder.glob("*.out"), key=lambda p: p.name.lower())

    if not out_files:
        print(f"No .out files found in: {folder}")
        return

    print(f"Found {len(out_files)} Mplus output file(s) in: {folder}")
    logging.info("Found %d Mplus output file(s)", len(out_files))

    models: List[ParsedModel] = []
    for i, fp in enumerate(out_files, start=1):
        print(f"Parsing: {fp.name}")
        logging.info("Parsing: %s", fp.name)
        models.append(parse_out_file(fp, idx=i))

    factor_names = collect_factor_names(models)

    wb = Workbook()
    ws0 = wb.active
    build_instructions(ws0, factor_names)

    ws1 = wb.create_sheet()
    build_table1(ws1, models)

    ws2 = wb.create_sheet()
    build_table2(ws2, models)

    ws3 = wb.create_sheet()
    build_reliability(ws3, models)

    ws4 = wb.create_sheet()
    build_figures(ws4, models, folder)

    # global font enforcement
    for ws in wb.worksheets:
        apply_global_format(ws)

    out_path = folder / OUTPUT_FILENAME
    try:
        wb.save(out_path)
    except PermissionError:
        # Common on Windows if the file is open in Excel.
        alt = out_path.with_name(out_path.stem + "_NEW" + out_path.suffix)
        wb.save(alt)
        out_path = alt

    print(f"Wrote: {out_path}")
    logging.info("Wrote workbook: %s", out_path)
    logging.info("Done.")

if __name__ == "__main__":
    try:
        main()
    except FileNotFoundError:
        print("ERROR: No Mplus .out files were found in the script folder.")
        print("Place this program in the folder containing your Mplus structural model .out files, then run it again.")
        raise
    except PermissionError:
        print("ERROR: The output Excel file appears to be open or locked.")
        print("Close the workbook and run again, or use the *_NEW file that was created.")
        raise
    except Exception:
        print("ERROR: Something went wrong while generating the Excel report.")
        print("Please share the 'run.log' file in this folder for troubleshooting.")
        raise
