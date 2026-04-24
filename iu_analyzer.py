"""
I-U Review Workstation · v4 — 3-Lane Architecture (Codex-revised)

Lane 1: classify_family()        — triage file into one of 5 families
Lane 2: family-specific extract  — only fill metrics where the family supports it
Lane 3: (UI layer in app.py)     — human review + override with audit trail

Families:
  "pitting"           — clear passive region + breakdown + hysteresis loop
  "transpassive"      — no breakdown, reverse returns along forward
  "prep_then_rp"      — conditioning scan followed by real RP phase
  "multi_cycle"       — multiple forward/reverse cycles
  "malformed"         — unusable / noisy / too-short

Metrics only populated when family warrants it:
  OCP       → always (blank only if no stable plateau at all)
  RPP       → ONLY for "pitting" with visible closure; else blank
  E_pit     → ONLY for "pitting"; else blank
  i_pass    → ONLY for "pitting" / "transpassive" (not multi-cycle or malformed)
  Hysteresis area → ONLY for "pitting"
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import numpy as np


# ═══════════════════════════════════════════════════════════════════════════
# Data classes
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class IUCurveData:
    time_s: np.ndarray
    potential_V: np.ndarray
    current_A: np.ndarray
    current_density_A_m2: np.ndarray
    n_rp: int = 300
    probenflaeche_mm2: Optional[float] = None
    filename: str = ""


@dataclass
class IUAnalysis:
    # Family classification (Lane 1)
    family: Optional[str] = None          # pitting | transpassive | prep_then_rp | multi_cycle | malformed
    family_conf: float = 0.0
    family_reason: str = ""
    # Extracted metrics (Lane 2) — blank if family doesn't support
    ocp_mV: Optional[float] = None
    rpp_mV: Optional[float] = None
    i_pass_mA_cm2: Optional[float] = None
    e_pit_mV: Optional[float] = None
    hysteresis_area: Optional[float] = None
    corrosion_rate_mm_per_year: Optional[float] = None
    # Confidence per metric (0-1)
    ocp_conf: float = 0.0
    rpp_conf: float = 0.0
    e_pit_conf: float = 0.0
    # Triage status derived from all confidences
    status: str = "red"                   # green | yellow | red
    status_reason: str = ""
    # Diagnostics for UI
    ocp_window_start: Optional[int] = None
    ocp_window_end: Optional[int] = None
    vertex_index: Optional[int] = None
    plateaus: list = field(default_factory=list)   # [(start, end, mean_V), ...]
    notes: list = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════════════════
# Parsing (unchanged from v3)
# ═══════════════════════════════════════════════════════════════════════════

def parse_had(path: Path) -> dict:
    meta = {}
    for enc in ("latin-1", "cp1252", "utf-8"):
        try:
            lines = path.read_text(encoding=enc).splitlines()
            break
        except Exception:
            continue
    else:
        return meta
    for line in lines:
        if "Anzahl Werte RP" in line:
            m = re.search(r":\s*(\d+)", line)
            if m: meta["n_rp"] = int(m.group(1))
        elif "Probenfl" in line:
            m = re.search(r":\s*([\d.]+)", line)
            if m: meta["probenflaeche_mm2"] = float(m.group(1))
    return meta


def parse_asc(asc_path: Path, had_path: Optional[Path] = None) -> IUCurveData:
    arr = np.loadtxt(asc_path)
    if arr.ndim != 2 or arr.shape[1] < 2:
        raise ValueError(f"ASC shape unexpected: {arr.shape}")
    time_s = arr[:, 0]
    potential_V = arr[:, 1]
    current_A = arr[:, 2] if arr.shape[1] >= 3 else np.zeros_like(time_s)
    j_Am2 = arr[:, 3] if arr.shape[1] >= 4 else np.zeros_like(time_s)
    meta = {"n_rp": 300, "probenflaeche_mm2": None}
    if had_path and had_path.exists():
        meta.update(parse_had(had_path))
    return IUCurveData(
        time_s=time_s, potential_V=potential_V,
        current_A=current_A, current_density_A_m2=j_Am2,
        n_rp=meta["n_rp"], probenflaeche_mm2=meta["probenflaeche_mm2"],
        filename=asc_path.name,
    )


# ═══════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════

def detect_flat_plateaus(pot_V: np.ndarray, win: int = 50,
                         std_threshold_V: float = 0.005,
                         search_fraction: float = 0.25) -> list[tuple[int, int, float]]:
    """Find flat plateaus in the first `search_fraction` of the trace.

    Returns list of (start_idx, end_idx, mean_V) tuples, sorted by position (earliest first).
    """
    if len(pot_V) < win * 2:
        return []
    search_end = max(win * 4, int(len(pot_V) * search_fraction))
    search = pot_V[:search_end]
    half = win // 2
    starts = list(range(0, len(search) - win, half))
    stds = np.array([search[s:s+win].std() for s in starts])
    flat = stds < std_threshold_V

    plateaus = []
    i = 0
    while i < len(flat):
        if flat[i]:
            j = i
            while j < len(flat) and flat[j]:
                j += 1
            start_pt = starts[i]
            end_pt = starts[j - 1] + win
            mean_V = float(np.median(pot_V[start_pt:end_pt]))
            plateaus.append((start_pt, min(end_pt, len(pot_V) - 1), mean_V))
            i = j
        else:
            i += 1
    return plateaus


def find_vertex(pot_V: np.ndarray, search_start: int) -> int:
    if search_start >= len(pot_V) - 50:
        return len(pot_V) - 1
    pot = pot_V[search_start:]
    window = 21
    kernel = np.ones(window) / window
    if len(pot) < window:
        return search_start + int(np.argmax(pot))
    smooth = np.convolve(pot, kernel, mode="same")
    return search_start + int(np.argmax(smooth))


def count_direction_changes(pot_V: np.ndarray, threshold_V: float = 0.050) -> int:
    """Count potential direction reversals (vertices) — indicates multi-cycle data."""
    if len(pot_V) < 200:
        return 0
    # Use decimated signal to avoid noise
    decim = pot_V[::10]
    # Smooth
    win = max(5, len(decim) // 100)
    kernel = np.ones(win) / win
    smooth = np.convolve(decim, kernel, mode="same")
    # Find local maxima separated by > threshold_V in potential range
    from numpy.lib.stride_tricks import sliding_window_view
    if len(smooth) < 50:
        return 0
    win_size = max(20, len(smooth) // 20)
    # Peak detection via simple argmax in sliding windows
    peaks = 0
    prev_peak_val = -np.inf
    for i in range(0, len(smooth) - win_size, win_size):
        window_vals = smooth[i:i+win_size]
        peak_val = window_vals.max()
        if peak_val > prev_peak_val + threshold_V:
            peaks += 1
        prev_peak_val = peak_val
    return max(0, peaks - 1)


# ═══════════════════════════════════════════════════════════════════════════
# Lane 1 — Family classification
# ═══════════════════════════════════════════════════════════════════════════

def classify_family(data: IUCurveData) -> tuple[str, float, str, list]:
    """Classify the curve file into one of 5 families.

    Returns: (family_name, confidence, reason, plateaus_found)
    """
    n_total = len(data.potential_V)
    plateaus = detect_flat_plateaus(data.potential_V, win=50,
                                    std_threshold_V=0.005, search_fraction=0.25)

    # Basic sanity checks
    if n_total < 500:
        return "malformed", 0.9, f"only {n_total} data points (too short)", plateaus

    # Check for multi-cycle (2+ vertices)
    n_cycles = count_direction_changes(data.potential_V, threshold_V=0.500)
    if n_cycles >= 2:
        return "multi_cycle", 0.8, f"detected {n_cycles} potential reversals", plateaus

    # Detect preparation scan: large excursion in first part (>100 mV swing) before stable plateau
    first_third = data.potential_V[:n_total // 3]
    if len(first_third) > 100:
        first_range_mV = (first_third.max() - first_third.min()) * 1000
        if len(plateaus) >= 2 and first_range_mV > 200:
            # Significant excursion + 2+ plateaus = probably has prep scan
            return "prep_then_rp", 0.75, f"{first_range_mV:.0f}mV swing + {len(plateaus)} plateaus", plateaus

    # Rest: pitting vs transpassive — determined by hysteresis loop
    # Quick test: find vertex, then compare reverse-sweep j to forward-sweep j at mid-potential
    if plateaus:
        first_plateau_end = plateaus[0][1]
    else:
        first_plateau_end = min(300, n_total // 20)

    vertex = find_vertex(data.potential_V, first_plateau_end)
    if vertex - first_plateau_end < 100 or n_total - vertex < 100:
        return "malformed", 0.7, "too few samples around vertex", plateaus

    fwd_pot = data.potential_V[first_plateau_end:vertex]
    fwd_j = np.abs(data.current_density_A_m2[first_plateau_end:vertex]) * 0.1
    rev_pot = data.potential_V[vertex:]
    rev_j = np.abs(data.current_density_A_m2[vertex:]) * 0.1

    # Mid-potential (avg of vertex and OCP-ish)
    mid_pot = (fwd_pot.min() + fwd_pot.max()) / 2

    # At mid-potential, find j on each sweep
    try:
        fwd_mid_j = np.interp(mid_pot, fwd_pot, fwd_j)
        rev_sorted = np.argsort(rev_pot)
        rev_mid_j = np.interp(mid_pot, rev_pot[rev_sorted], rev_j[rev_sorted])
        hyst_ratio = rev_mid_j / max(fwd_mid_j, 1e-6)
    except Exception:
        hyst_ratio = 1.0

    if hyst_ratio > 3:
        return "pitting", 0.85, f"clear hysteresis loop (reverse j = {hyst_ratio:.1f}× forward at mid-E)", plateaus
    elif hyst_ratio > 1.3:
        return "pitting", 0.6, f"weak hysteresis loop (ratio {hyst_ratio:.1f})", plateaus
    else:
        return "transpassive", 0.8, f"no hysteresis (reverse ≈ forward, ratio {hyst_ratio:.2f})", plateaus


# ═══════════════════════════════════════════════════════════════════════════
# Lane 2 — Family-specific extractors (only fill what's supported)
# ═══════════════════════════════════════════════════════════════════════════

def extract_ocp_from_plateaus(plateaus: list, family: str,
                              pot_V: np.ndarray) -> tuple[Optional[float], float, int, int, str]:
    """Pick the right plateau and return (ocp_mV, confidence, start, end, note)."""
    if not plateaus:
        return None, 0.0, 0, 0, "no flat plateau found"

    if family == "prep_then_rp":
        # Pick 2nd plateau (skip preparation)
        if len(plateaus) >= 2:
            start, end, mean_V = plateaus[1]
            return round(mean_V * 1000, 2), 0.7, start, end, "used 2nd plateau (skip prep)"
        else:
            start, end, mean_V = plateaus[0]
            return round(mean_V * 1000, 2), 0.5, start, end, "only 1 plateau despite prep family"

    # Default: pick first (earliest) plateau
    start, end, mean_V = plateaus[0]
    length = end - start
    if length > 200:
        conf = 0.9
    elif length > 100:
        conf = 0.75
    elif length > 50:
        conf = 0.55
    else:
        conf = 0.35
    return round(mean_V * 1000, 2), conf, start, end, f"first plateau ({length} pts)"


def _analyze_pitting(data: IUCurveData, ocp_window_end: int, vertex: int,
                     material_rho: float = 7.85, material_M: float = 55.85,
                     material_z: int = 2) -> dict:
    """Only for clear pitting curves. Returns metrics dict."""
    fwd_pot = data.potential_V[ocp_window_end:vertex]
    fwd_j = np.abs(data.current_density_A_m2[ocp_window_end:vertex]) * 0.1
    rev_pot = data.potential_V[vertex:]
    rev_j = np.abs(data.current_density_A_m2[vertex:]) * 0.1

    result = {"i_pass_mA_cm2": None, "e_pit_mV": None, "e_pit_conf": 0.0,
              "rpp_mV": None, "rpp_conf": 0.0,
              "hysteresis_area": None, "corrosion_rate_mm_per_year": None}

    if len(fwd_pot) < 100 or len(rev_pot) < 100:
        return result

    # i_pass — 25th percentile in first half of forward sweep (before pitting)
    passive_mask = fwd_j < np.percentile(fwd_j, 50)
    if passive_mask.sum() > 30:
        i_pass = float(np.percentile(fwd_j[passive_mask], 25))
        result["i_pass_mA_cm2"] = round(i_pass, 6)

        # Corrosion rate via Faraday
        i_corr_uA = i_pass * 1000
        EW = material_M / material_z
        result["corrosion_rate_mm_per_year"] = round(3.27e-3 * i_corr_uA * EW / material_rho, 6)

    # E_pit — first point where d(log j)/dE > 0.05 AND j > 10× i_pass
    if result["i_pass_mA_cm2"] is not None:
        log_j = np.log10(np.maximum(fwd_j, 1e-8))
        sw = 11
        log_j_smooth = np.convolve(log_j, np.ones(sw)/sw, mode="same")
        dE = np.gradient(fwd_pot)
        dlogj = np.gradient(log_j_smooth)
        with np.errstate(divide="ignore", invalid="ignore"):
            slope = np.where(np.abs(dE) > 1e-6, dlogj / dE, 0.0)
        pit_candidates = np.where(
            (slope > 0.05) & (fwd_j > result["i_pass_mA_cm2"] * 10)
        )[0]
        if len(pit_candidates) > 0:
            idx = int(pit_candidates[0])
            result["e_pit_mV"] = round(float(fwd_pot[idx]) * 1000, 2)
            result["e_pit_conf"] = 0.8

    # RPP — hysteresis-close detection (only if E_pit was found → real pitting curve)
    if result["e_pit_mV"] is not None:
        try:
            fwd_sort_idx = np.argsort(fwd_pot)
            fwd_pot_s = fwd_pot[fwd_sort_idx]
            fwd_j_s = fwd_j[fwd_sort_idx]
            _, uniq = np.unique(fwd_pot_s, return_index=True)
            fwd_pot_u = fwd_pot_s[uniq]
            fwd_j_u = fwd_j_s[uniq]
            j_fwd_at_rev = np.interp(rev_pot, fwd_pot_u, fwd_j_u,
                                     left=np.nan, right=np.nan)
            delta = rev_j - j_fwd_at_rev
            tolerance = np.maximum(j_fwd_at_rev * 0.1, 0.002)
            sustained = 0
            for i in range(len(rev_j)):
                if np.isnan(j_fwd_at_rev[i]):
                    continue
                if delta[i] > tolerance[i] * 2:
                    sustained += 1
                elif sustained >= 20 and delta[i] < tolerance[i]:
                    result["rpp_mV"] = round(float(rev_pot[i]) * 1000, 2)
                    result["rpp_conf"] = 0.7
                    break
        except Exception:
            pass

    # Hysteresis area
    try:
        e_min = max(fwd_pot.min(), rev_pot.min())
        e_max = min(fwd_pot.max(), rev_pot.max())
        if e_max > e_min:
            e_grid = np.linspace(e_min, e_max, 500)
            fwd_s = np.argsort(fwd_pot)
            rev_s = np.argsort(rev_pot)
            j_fwd_i = np.interp(e_grid, fwd_pot[fwd_s], fwd_j[fwd_s])
            j_rev_i = np.interp(e_grid, rev_pot[rev_s], rev_j[rev_s])
            area = float(np.trapezoid(np.abs(j_rev_i - j_fwd_i), e_grid))
            result["hysteresis_area"] = round(area, 6)
    except Exception:
        pass

    return result


def _analyze_transpassive(data: IUCurveData, ocp_window_end: int, vertex: int,
                          material_rho: float = 7.85, material_M: float = 55.85,
                          material_z: int = 2) -> dict:
    """For transpassive curves: only i_pass + corrosion rate. No RPP. No E_pit."""
    fwd_pot = data.potential_V[ocp_window_end:vertex]
    fwd_j = np.abs(data.current_density_A_m2[ocp_window_end:vertex]) * 0.1

    result = {"i_pass_mA_cm2": None, "e_pit_mV": None, "e_pit_conf": 0.0,
              "rpp_mV": None, "rpp_conf": 0.0,
              "hysteresis_area": None, "corrosion_rate_mm_per_year": None}

    if len(fwd_pot) < 100:
        return result

    # i_pass = 25th percentile of passive region
    if len(fwd_j) > 50:
        passive_mask = fwd_j < np.percentile(fwd_j, 40)
        if passive_mask.sum() > 30:
            i_pass = float(np.percentile(fwd_j[passive_mask], 25))
            result["i_pass_mA_cm2"] = round(i_pass, 6)
            i_corr_uA = i_pass * 1000
            EW = material_M / material_z
            result["corrosion_rate_mm_per_year"] = round(3.27e-3 * i_corr_uA * EW / material_rho, 6)

    return result


# ═══════════════════════════════════════════════════════════════════════════
# Main dispatcher
# ═══════════════════════════════════════════════════════════════════════════

def analyze(data: IUCurveData, material_rho_g_cm3: float = 7.85,
            material_M_g_mol: float = 55.85, material_z: int = 2) -> IUAnalysis:
    a = IUAnalysis()

    # Lane 1: classify family
    family, family_conf, family_reason, plateaus = classify_family(data)
    a.family = family
    a.family_conf = family_conf
    a.family_reason = family_reason
    a.plateaus = [(p[0], p[1], round(p[2]*1000, 2)) for p in plateaus]

    # If malformed or multi_cycle, bail out early with minimal metrics
    if family in ("malformed",):
        a.status = "red"
        a.status_reason = family_reason
        a.notes.append(f"family={family}: {family_reason} — needs manual review")
        return a

    # OCP always tried — extract from plateaus with family-aware logic
    ocp, ocp_conf, w_start, w_end, ocp_note = extract_ocp_from_plateaus(
        plateaus, family, data.potential_V
    )
    a.ocp_mV = ocp
    a.ocp_conf = ocp_conf
    a.ocp_window_start = w_start
    a.ocp_window_end = w_end
    a.notes.append(f"OCP: {ocp_note}")

    # Find vertex
    vertex = find_vertex(data.potential_V, w_end)
    a.vertex_index = vertex

    # Lane 2: dispatch to family-specific extractor
    if family == "pitting":
        metrics = _analyze_pitting(data, w_end, vertex,
                                   material_rho_g_cm3, material_M_g_mol, material_z)
    elif family == "transpassive":
        metrics = _analyze_transpassive(data, w_end, vertex,
                                        material_rho_g_cm3, material_M_g_mol, material_z)
    elif family == "prep_then_rp":
        # Treat as pitting (with skip-prep OCP already applied)
        metrics = _analyze_pitting(data, w_end, vertex,
                                   material_rho_g_cm3, material_M_g_mol, material_z)
    elif family == "multi_cycle":
        # Don't auto-extract metrics for multi-cycle; needs manual segmentation
        metrics = {"i_pass_mA_cm2": None, "e_pit_mV": None, "e_pit_conf": 0.0,
                   "rpp_mV": None, "rpp_conf": 0.0,
                   "hysteresis_area": None, "corrosion_rate_mm_per_year": None}
        a.notes.append("multi_cycle: manual segmentation required — no auto metrics")
    else:
        metrics = {"i_pass_mA_cm2": None, "e_pit_mV": None, "e_pit_conf": 0.0,
                   "rpp_mV": None, "rpp_conf": 0.0,
                   "hysteresis_area": None, "corrosion_rate_mm_per_year": None}

    a.i_pass_mA_cm2 = metrics["i_pass_mA_cm2"]
    a.e_pit_mV = metrics["e_pit_mV"]
    a.e_pit_conf = metrics["e_pit_conf"]
    a.rpp_mV = metrics["rpp_mV"]
    a.rpp_conf = metrics["rpp_conf"]
    a.hysteresis_area = metrics["hysteresis_area"]
    a.corrosion_rate_mm_per_year = metrics["corrosion_rate_mm_per_year"]

    # Lane 3 (partial): compute triage status
    min_conf = min(a.ocp_conf, a.e_pit_conf if a.e_pit_mV else 1.0,
                   a.rpp_conf if a.rpp_mV else 1.0,
                   a.family_conf)
    if min_conf >= 0.75 and family in ("pitting", "transpassive"):
        a.status = "green"
        a.status_reason = "all confidences ≥ 0.75"
    elif min_conf >= 0.55 or family == "prep_then_rp":
        a.status = "yellow"
        a.status_reason = f"min confidence {min_conf:.2f} (review recommended)"
    else:
        a.status = "red"
        a.status_reason = f"min confidence {min_conf:.2f} (manual required)"

    return a


def reanalyze_with_overrides(data: IUCurveData, original: IUAnalysis,
                              overrides: dict,
                              material_rho_g_cm3: float = 7.85,
                              material_M_g_mol: float = 55.85,
                              material_z: int = 2) -> IUAnalysis:
    """
    When the user overrides OCP (or vertex position), recompute all downstream
    metrics (i_pass, E_pit, RPP, hysteresis area, corrosion rate) using the new
    OCP/vertex positions.

    Overrides dict can contain:
      - ocp_mV:  float → we use this value + re-derive ocp_window_end from
                  nearest-matching-potential time index
      - rpp_mV, e_pit_mV: used directly in the result (no downstream effect)
      - family: str → force family classification (overrides auto)
    """
    import copy
    a = copy.copy(original)

    # If OCP was overridden → recompute everything downstream
    if "ocp_mV" in overrides and overrides["ocp_mV"] is not None:
        new_ocp_V = overrides["ocp_mV"] / 1000.0
        # Find the time index where potential is closest to new OCP
        # (restrict to first 20% of data so we don't pick a sweep point)
        search_end = max(1000, len(data.potential_V) // 5)
        residuals = np.abs(data.potential_V[:search_end] - new_ocp_V)
        best_idx = int(np.argmin(residuals))
        # Use a small window around best_idx as the new RP region
        new_window_end = min(best_idx + 200, len(data.potential_V) - 1)
        a.ocp_mV = overrides["ocp_mV"]
        a.ocp_window_end = new_window_end
        a.ocp_conf = 1.0  # user-set = full confidence

        # Re-find vertex after new window
        new_vertex = find_vertex(data.potential_V, new_window_end)
        a.vertex_index = new_vertex

        # Re-run family-specific extraction
        family = overrides.get("family", a.family)
        if family == "pitting" or family == "prep_then_rp":
            metrics = _analyze_pitting(data, new_window_end, new_vertex,
                                       material_rho_g_cm3, material_M_g_mol, material_z)
        elif family == "transpassive":
            metrics = _analyze_transpassive(data, new_window_end, new_vertex,
                                            material_rho_g_cm3, material_M_g_mol, material_z)
        else:
            metrics = {"i_pass_mA_cm2": None, "e_pit_mV": None, "e_pit_conf": 0.0,
                       "rpp_mV": None, "rpp_conf": 0.0,
                       "hysteresis_area": None, "corrosion_rate_mm_per_year": None}

        a.i_pass_mA_cm2 = metrics["i_pass_mA_cm2"]
        a.corrosion_rate_mm_per_year = metrics["corrosion_rate_mm_per_year"]
        a.hysteresis_area = metrics["hysteresis_area"]
        # Only overwrite E_pit / RPP if they weren't also manually overridden
        if "e_pit_mV" not in overrides:
            a.e_pit_mV = metrics["e_pit_mV"]
            a.e_pit_conf = metrics["e_pit_conf"]
        if "rpp_mV" not in overrides:
            a.rpp_mV = metrics["rpp_mV"]
            a.rpp_conf = metrics["rpp_conf"]

        a.notes.append(f"recomputed with OCP override @ {overrides['ocp_mV']} mV")

    # Apply direct value overrides (E_pit, RPP, family)
    if "rpp_mV" in overrides and overrides["rpp_mV"] is not None:
        a.rpp_mV = overrides["rpp_mV"]
        a.rpp_conf = 1.0
    if "e_pit_mV" in overrides and overrides["e_pit_mV"] is not None:
        a.e_pit_mV = overrides["e_pit_mV"]
        a.e_pit_conf = 1.0
    if "family" in overrides and overrides["family"]:
        a.family = overrides["family"]
        a.family_conf = 1.0
        a.family_reason = "user override"

    # Recompute triage status based on final values
    if a.ocp_mV is not None and (a.i_pass_mA_cm2 is not None or a.family == "malformed"):
        if all(k in overrides for k in []):  # placeholder
            pass
        # If any override applied, mark as "override-complete" = green
        if overrides:
            a.status = "green"
            a.status_reason = "user overrides applied"
        elif min(a.ocp_conf, a.family_conf) >= 0.75:
            a.status = "green"
            a.status_reason = "high auto-confidence"
        elif min(a.ocp_conf, a.family_conf) >= 0.55:
            a.status = "yellow"
            a.status_reason = "moderate auto-confidence"
        else:
            a.status = "red"
            a.status_reason = "low auto-confidence"

    return a


def filename_to_supplier_medium(fname: str) -> tuple[Optional[str], Optional[str]]:
    parts = fname.replace(".ASC", "").split("_")
    supplier = None
    for p in parts:
        if p in ("P1", "P2", "P3", "P4"):
            supplier = p
            break
    fname_lower = fname.lower()
    if "sea salt" in fname_lower or "sea water" in fname_lower:
        medium = "Sea Water"
    elif "0.6m nacl" in fname_lower or "mgcl2" in fname_lower:
        medium = "NaCl+MgCl2"
    elif "1m nacl" in fname_lower:
        medium = "1 MNaCl"
    else:
        medium = None
    return supplier, medium
