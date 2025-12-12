"""
Grade-range suggestion for rejected cells — two methods:
A) equal-width bins
B) k-means clustering (k=6)

Input:
  rejected_cells: list of dicts:
    [{"cell_id":"C001","measured_voltage":3.65,"measured_current":0.7}, ...]
Returns:
  {
    "grades": [
      {"grade_name":"Grade 1","vmin":2.50,"vmax":3.21,"count":123,"pct":12.3},
      ...
    ],
    "total_cells": N,
    "accepted_count": M,
    "accepted_pct": P,
    "ignored_outliers_count": O
  }
Notes:
 - Uses IQR method to remove extreme outliers (configurable)
 - Rounds ranges to `round_digits` decimals (default 2)
 - KMeans requires scikit-learn; code will ask to pip install if missing
"""

from typing import List, Dict, Tuple, Optional
import math


class GradeSuggestionEngine:
    """
    Class for generating grade range suggestions for rejected cells.
    Provides two methods: equal-width binning and k-means clustering.
    """

    def __init__(self, grade_count: int = 6, iqr_multiplier: float = 1.5, round_digits: int = 2):
        """
        Initialize the grade suggestion engine.

        Args:
            grade_count: Number of grades to suggest (default 6)
            iqr_multiplier: IQR multiplier for outlier detection (default 1.5)
            round_digits: Decimal places for rounding (default 2)
        """
        self.grade_count = grade_count
        self.iqr_multiplier = iqr_multiplier
        self.round_digits = round_digits

    def suggest_both_methods(self, rejected_cells: List[Dict], random_state: Optional[int] = 42) -> Dict:
        """
        Generate suggestions using both methods.

        Args:
            rejected_cells: List of dicts with voltage/current data
            random_state: Random seed for k-means (default 42)

        Returns:
            Dict with both equal_width and kmeans results
        """
        equal_width_result = self.suggest_ranges_equal_width(rejected_cells)
        try:
            kmeans_result = self.suggest_ranges_kmeans(rejected_cells, random_state=random_state)
        except RuntimeError as e:
            kmeans_result = {"error": str(e), "grades": [], "total_cells": 0,
                             "accepted_count": 0, "accepted_pct": 0.0, "ignored_outliers_count": 0}

        return {
            "equal_width": equal_width_result,
            "kmeans": kmeans_result
        }

    def suggest_ranges_equal_width(self, rejected_cells: List[Dict]) -> Dict:
        """Generate grade ranges using equal-width binning."""
        volts = self._extract_voltages(rejected_cells)
        total = len(volts)
        if total == 0:
            return {"grades": [], "total_cells": 0, "accepted_count": 0, "accepted_pct": 0.0,
                    "ignored_outliers_count": 0}

        filtered_volts, kept_idx = self._iqr_filter(volts, iqr_multiplier=self.iqr_multiplier)
        ignored = total - len(filtered_volts)
        if not filtered_volts:
            return {"grades": [], "total_cells": total, "accepted_count": 0, "accepted_pct": 0.0,
                    "ignored_outliers_count": ignored}

        vmin = min(filtered_volts)
        vmax = max(filtered_volts)
        if vmin == vmax:
            bins = [(vmin, vmax) for _ in range(self.grade_count)]
        else:
            width = (vmax - vmin) / self.grade_count
            bins = []
            for i in range(self.grade_count):
                bmin = vmin + i * width
                bmax = (vmin + (i + 1) * width) if i < self.grade_count - 1 else vmax
                bins.append((bmin, bmax))

        bins = self._make_non_overlapping(bins, round_digits=self.round_digits)

        grades = []
        for idx, (mn, mx) in enumerate(bins):
            count = sum(1 for v in filtered_volts if mn <= v <= mx)
            pct = 100.0 * count / total if total > 0 else 0.0
            grades.append({
                "grade_name": f"Grade {idx + 1}",
                "vmin": round(mn, self.round_digits),
                "vmax": round(mx, self.round_digits),
                "count": count,
                "pct": round(pct, 2)
            })
        accepted_count = sum(g["count"] for g in grades)
        accepted_pct = round(100.0 * accepted_count / total, 2) if total > 0 else 0.0
        return {
            "grades": grades,
            "total_cells": total,
            "accepted_count": accepted_count,
            "accepted_pct": accepted_pct,
            "ignored_outliers_count": ignored
        }

    def suggest_ranges_kmeans(self, rejected_cells: List[Dict], random_state: Optional[int] = 42) -> Dict:
        """Generate grade ranges using k-means clustering."""
        volts = self._extract_voltages(rejected_cells)
        total = len(volts)
        if total == 0:
            return {"grades": [], "total_cells": 0, "accepted_count": 0, "accepted_pct": 0.0,
                    "ignored_outliers_count": 0}
        filtered_volts, kept_idx = self._iqr_filter(volts, iqr_multiplier=self.iqr_multiplier)
        ignored = total - len(filtered_volts)
        if not filtered_volts:
            return {"grades": [], "total_cells": total, "accepted_count": 0, "accepted_pct": 0.0,
                    "ignored_outliers_count": ignored}

        X = [[v] for v in filtered_volts]

        try:
            from sklearn.cluster import KMeans
        except Exception as e:
            raise RuntimeError("scikit-learn is required for kmeans. Install with: pip install scikit-learn") from e

        if len(filtered_volts) < self.grade_count:
            unique_sorted = sorted(set(filtered_volts))
            bins = []
            for i, uv in enumerate(unique_sorted):
                mn = uv - 0.0001
                mx = uv + 0.0001
                bins.append((mn, mx))
            while len(bins) < self.grade_count:
                bins.append((unique_sorted[-1] + 0.001 * len(bins), unique_sorted[-1] + 0.002 * len(bins)))
            bins = self._make_non_overlapping(bins, round_digits=self.round_digits)
        else:
            km = KMeans(n_clusters=self.grade_count, random_state=random_state, n_init='auto')
            labels = km.fit_predict(X)
            cluster_ranges = []
            for k in range(self.grade_count):
                cluster_vs = [filtered_volts[i] for i, lab in enumerate(labels) if lab == k]
                if cluster_vs:
                    cluster_ranges.append((min(cluster_vs), max(cluster_vs)))
            cluster_ranges = sorted(cluster_ranges, key=lambda x: x[0])
            bins = self._make_non_overlapping(cluster_ranges, round_digits=self.round_digits)

        grades = []
        for idx, (mn, mx) in enumerate(bins):
            count = sum(1 for v in filtered_volts if mn <= v <= mx)
            pct = 100.0 * count / total if total > 0 else 0.0
            grades.append({
                "grade_name": f"Grade {idx + 1}",
                "vmin": round(mn, self.round_digits),
                "vmax": round(mx, self.round_digits),
                "count": count,
                "pct": round(pct, 2)
            })
        accepted_count = sum(g["count"] for g in grades)
        accepted_pct = round(100.0 * accepted_count / total, 2) if total > 0 else 0.0
        return {
            "grades": grades,
            "total_cells": total,
            "accepted_count": accepted_count,
            "accepted_pct": accepted_pct,
            "ignored_outliers_count": ignored
        }

    @staticmethod
    def _extract_voltages(rejected_cells: List[Dict]) -> List[float]:
        volts = []
        for c in rejected_cells:
            v = c.get("measured_voltage", None)
            if v is None:
                v = c.get("measured_voltage") or c.get("voltage")
            try:
                volts.append(float(v))
            except Exception:
                continue
        return volts

    @staticmethod
    def _iqr_filter(volts: List[float], iqr_multiplier: float = 1.5) -> Tuple[List[float], List[int]]:
        """Return (filtered_volts, indices_kept). Removes points outside [Q1 - k*IQR, Q3 + k*IQR]."""
        if not volts:
            return [], []
        sorted_idx = sorted(range(len(volts)), key=lambda i: volts[i])
        sorted_v = [volts[i] for i in sorted_idx]
        n = len(sorted_v)

        def _quantile(arr, q):
            pos = (len(arr) - 1) * q
            lo = math.floor(pos)
            hi = math.ceil(pos)
            if lo == hi:
                return arr[int(pos)]
            frac = pos - lo
            return arr[lo] * (1 - frac) + arr[hi] * frac

        q1 = _quantile(sorted_v, 0.25)
        q3 = _quantile(sorted_v, 0.75)
        iqr = q3 - q1
        low = q1 - iqr_multiplier * iqr
        high = q3 + iqr_multiplier * iqr
        kept = [i for i, v in enumerate(volts) if low <= v <= high]
        filtered = [volts[i] for i in kept]
        return filtered, kept

    @staticmethod
    def _make_non_overlapping(sorted_ranges: List[Tuple[float, float]], round_digits: int = 2) -> List[
        Tuple[float, float]]:
        """
        Takes list of (min,max) possibly touching/overlapping, sorts them by min,
        ensures non-overlap by shrinking tiny gaps or adjusting boundaries so ranges
        are consecutive. Rounds results.
        """
        if not sorted_ranges:
            return []
        sorted_ranges = sorted(sorted_ranges, key=lambda x: x[0])
        out = []
        prev_max = None
        for (mn, mx) in sorted_ranges:
            mn = float(mn);
            mx = float(mx)
            if prev_max is None:
                curr_min = mn
            else:
                curr_min = max(mn, prev_max + 10 ** (-round_digits))
            curr_max = max(curr_min, mx)
            prev_max = curr_max
            out.append((round(curr_min, round_digits), round(curr_max, round_digits)))
        return out

# # Legacy functions for backward compatibility
# def suggest_ranges_equal_width(
#     rejected_cells: List[Dict],
#     grade_count: int = 6,
#     iqr_multiplier: float = 1.5,
#     round_digits: int = 2
# ) -> Dict:
#     """Legacy function - use GradeSuggestionEngine class instead."""
#     engine = GradeSuggestionEngine(grade_count, iqr_multiplier, round_digits)
#     return engine.suggest_ranges_equal_width(rejected_cells)


# def suggest_ranges_kmeans(
#     rejected_cells: List[Dict],
#     grade_count: int = 6,
#     iqr_multiplier: float = 1.5,
#     round_digits: int = 2,
#     random_state: Optional[int] = 42
# ) -> Dict:
#     """Legacy function - use GradeSuggestionEngine class instead."""
#     engine = GradeSuggestionEngine(grade_count, iqr_multiplier, round_digits)
#     return engine.suggest_ranges_kmeans(rejected_cells, random_state)


# def _extract_voltages(rejected_cells: List[Dict]) -> List[float]:
#     return GradeSuggestionEngine._extract_voltages(rejected_cells)


# def _iqr_filter(volts: List[float], iqr_multiplier: float = 1.5) -> Tuple[List[float], List[int]]:
#     return GradeSuggestionEngine._iqr_filter(volts, iqr_multiplier)


# def _make_non_overlapping(sorted_ranges: List[Tuple[float,float]], round_digits:int=2) -> List[Tuple[float,float]]:
#     return GradeSuggestionEngine._make_non_overlapping(sorted_ranges, round_digits)


# # Utility to pretty-print the report
# def print_report(result: Dict, title:str="Suggested Grade Ranges"):
#     print(f"=== {title} ===")
#     print(f"Total rejected cells analyzed : {result['total_cells']}")
#     print(f"Ignored (outliers)           : {result['ignored_outliers_count']}")
#     print()
#     for g in result["grades"]:
#         print(f"{g['grade_name']}: {g['vmin']}V - {g['vmax']}V  ->  {g['count']} cells  ({g['pct']}%)")
#     print()
#     print(f"Total accepted by these ranges: {result['accepted_count']} ({result['accepted_pct']}%)")
#     print("===============================")


# # Example usage with synthetic data
# if __name__ == "__main__":
#     import random
#     random.seed(0)
#     # simulate many rejected voltages with three clusters + outliers
#     volts = ([random.gauss(3.2,0.08) for _ in range(1500)]
#              + [random.gauss(4.1,0.05) for _ in range(2500)]
#              + [random.gauss(5.6,0.07) for _ in range(1200)]
#              + [6.8, 7.0, 1.2]  # outliers
#              )
#     rejected_cells = [{"cell_id": f"C{i+1:05d}", "measured_voltage": v, "measured_current": 0.5} for i,v in enumerate(volts)]

#     # Using the new class
#     engine = GradeSuggestionEngine()
#     results = engine.suggest_both_methods(rejected_cells)

#     print_report(results["equal_width"], "Equal-width bins (6)")
#     if "error" not in results["kmeans"]:
#         print_report(results["kmeans"], "KMeans clusters (6)")
#     else:
#         print("KMeans unavailable:", results["kmeans"]["error"])
