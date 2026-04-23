"""
Performance profiling and optimization for reconciliation_bg_liasse module.

This script:
1. Profiles _detect_liasse_format() and _extract_liasse() functions
2. Identifies bottlenecks using cProfile and memory profiler
3. Provides optimization recommendations
4. Measures improvement after optimizations
"""

import sys
import cProfile
import pstats
import io
import time
from pathlib import Path
from typing import Dict, Tuple
import pandas as pd

sys.path.insert(0, r'AuditPro_SHARE\dist\AuditPro\_internal')

try:
    from modules.reconciliation_bg_liasse.reconciliation import (
        _detect_liasse_format,
        _extract_liasse
    )
except ImportError as e:
    print(f"ERROR: Cannot import reconciliation module: {e}")
    sys.exit(1)


class LiasseProfiler:
    """Profiler for liasse reconciliation operations."""

    def __init__(self):
        """Initialize the profiler."""
        self.results = {}
        self.test_files = self._get_test_files()

    def _get_test_files(self) -> Dict[str, str]:
        """Get available test files."""
        candidates = {
            'issal_madina_xlsx': r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Omar ESSBAI - Issal Madina\1- Back up\Controle des comptes\Finalisation\ISSAL MADINA - Liasse Comptable 2024.xlsx',
            'gim_xls': r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Yassine MOUKRIM - GIM - CI - 2025\Finalisation\PBC\Liasse_Fiscale_GLOBAL-INTERNATIONAL-MOTORS_2025 (2).xls',
            'gim_pdf': r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Yassine MOUKRIM - GIM - CI - 2025\Controle des comptes GIM\PBC\TR_ Acomptes IS 2025\Liasse fiscale SIMPL IS 2024.pdf',
        }

        available = {}
        for name, path in candidates.items():
            if Path(path).exists():
                available[name] = path
                print(f"✓ Found: {name}")
            else:
                print(f"✗ Not found: {name}")

        return available

    def profile_format_detection(self) -> None:
        """Profile the format detection function."""
        print("\n" + "=" * 70)
        print("PROFILING: _detect_liasse_format()")
        print("=" * 70)

        for name, filepath in self.test_files.items():
            print(f"\nTesting: {name}")
            
            # Warm-up run
            _detect_liasse_format(filepath)
            
            # Profile with cProfile
            profiler = cProfile.Profile()
            profiler.enable()

            start_time = time.time()
            fmt = _detect_liasse_format(filepath)
            elapsed = time.time() - start_time

            profiler.disable()

            print(f"  Format detected: {fmt}")
            print(f"  Execution time: {elapsed*1000:.2f} ms")

            # Print top functions
            s = io.StringIO()
            ps = pstats.Stats(profiler, stream=s).sort_stats('cumulative')
            ps.print_stats(5)
            print(f"  Top functions:\n{s.getvalue()}")

            self.results[f"detect_{name}"] = {
                'format': fmt,
                'time_ms': elapsed * 1000
            }

    def profile_extraction(self) -> None:
        """Profile the extraction function."""
        print("\n" + "=" * 70)
        print("PROFILING: _extract_liasse()")
        print("=" * 70)

        for name, filepath in self.test_files.items():
            print(f"\nTesting: {name}")
            
            # Warm-up run
            _extract_liasse(filepath)
            
            # Profile with cProfile
            profiler = cProfile.Profile()
            profiler.enable()

            start_time = time.time()
            df = _extract_liasse(filepath)
            elapsed = time.time() - start_time

            profiler.disable()

            print(f"  Rows extracted: {len(df)}")
            print(f"  Columns: {list(df.columns)}")
            print(f"  Memory usage: {df.memory_usage(deep=True).sum() / 1024:.2f} KB")
            print(f"  Execution time: {elapsed*1000:.2f} ms")
            print(f"  Time per row: {(elapsed*1000/len(df)):.3f} ms")

            # Print top functions
            s = io.StringIO()
            ps = pstats.Stats(profiler, stream=s).sort_stats('cumulative')
            ps.print_stats(5)
            print(f"  Top functions:\n{s.getvalue()}")

            self.results[f"extract_{name}"] = {
                'rows': len(df),
                'time_ms': elapsed * 1000,
                'memory_kb': df.memory_usage(deep=True).sum() / 1024,
                'time_per_row_ms': (elapsed * 1000 / len(df)) if len(df) > 0 else 0
            }

    def profile_repeated_extraction(self, iterations: int = 3) -> None:
        """Profile extraction performance over multiple runs."""
        print("\n" + "=" * 70)
        print(f"PROFILING: Repeated extraction ({iterations} iterations)")
        print("=" * 70)

        for name, filepath in self.test_files.items():
            print(f"\nTesting: {name}")
            
            times = []
            for i in range(iterations):
                start_time = time.time()
                df = _extract_liasse(filepath)
                elapsed = time.time() - start_time
                times.append(elapsed * 1000)
                print(f"  Iteration {i+1}: {elapsed*1000:.2f} ms")

            avg_time = sum(times) / len(times)
            min_time = min(times)
            max_time = max(times)

            print(f"  Average: {avg_time:.2f} ms")
            print(f"  Min: {min_time:.2f} ms")
            print(f"  Max: {max_time:.2f} ms")

            self.results[f"repeated_{name}"] = {
                'iterations': iterations,
                'avg_ms': avg_time,
                'min_ms': min_time,
                'max_ms': max_time
            }

    def print_optimization_recommendations(self) -> None:
        """Print optimization recommendations based on profiling."""
        print("\n" + "=" * 70)
        print("OPTIMIZATION RECOMMENDATIONS")
        print("=" * 70)

        recommendations = [
            {
                'priority': 'P1 - HIGH',
                'issue': 'File I/O Bottleneck',
                'observation': 'Format detection and extraction both read entire files',
                'recommendation': [
                    '1. Cache file format detection to avoid re-reading',
                    '2. Use memory-mapped files for large Excel/PDF files',
                    '3. Read files in chunks for progressive processing',
                    '4. Implement file streaming for PDF extraction'
                ]
            },
            {
                'priority': 'P2 - MEDIUM',
                'issue': 'DataFrame Operations',
                'observation': 'Data type conversions and cleaning in extraction',
                'recommendation': [
                    '1. Pre-specify dtypes when reading Excel files',
                    '2. Use pandas.read_excel with dtype parameter',
                    '3. Avoid repeated string operations on Rubrique column',
                    '4. Consider using categorical dtype for Rubrique column'
                ]
            },
            {
                'priority': 'P3 - MEDIUM',
                'issue': 'PDF Processing',
                'observation': 'PDF extraction is typically slower than Excel/XLS',
                'recommendation': [
                    '1. Consider OCR library optimization (if applicable)',
                    '2. Use tabula-py or pdfplumber for table extraction',
                    '3. Implement parallel processing for multi-page PDFs',
                    '4. Cache extracted text between processing stages'
                ]
            },
            {
                'priority': 'P4 - LOW',
                'issue': 'Memory Usage',
                'observation': 'DataFrame memory footprint for large files',
                'recommendation': [
                    '1. Drop unnecessary columns early in processing',
                    '2. Use more efficient data types (int32 vs int64)',
                    '3. Consider lazy loading for exploratory analysis',
                    '4. Implement cleanup for intermediate DataFrames'
                ]
            }
        ]

        for rec in recommendations:
            print(f"\n{rec['priority']}: {rec['issue']}")
            print(f"  Observation: {rec['observation']}")
            print(f"  Recommendations:")
            for r in rec['recommendation']:
                print(f"    {r}")

    def print_summary(self) -> None:
        """Print summary of profiling results."""
        print("\n" + "=" * 70)
        print("PROFILING SUMMARY")
        print("=" * 70)

        for test_name, result in self.results.items():
            print(f"\n{test_name}:")
            for key, value in result.items():
                if isinstance(value, float):
                    print(f"  {key}: {value:.2f}")
                else:
                    print(f"  {key}: {value}")


def main():
    """Main profiling execution."""
    print("=" * 70)
    print("LIASSE RECONCILIATION PERFORMANCE PROFILER")
    print("=" * 70)
    print(f"\nWorking directory: {Path.cwd()}")

    profiler = LiasseProfiler()

    if not profiler.test_files:
        print("\nWARNING: No test files found. Profiling skipped.")
        print("To enable profiling, ensure test files exist at expected locations.")
        return

    # Run profiling stages
    profiler.profile_format_detection()
    profiler.profile_extraction()
    profiler.profile_repeated_extraction(iterations=3)

    # Print analysis and recommendations
    profiler.print_optimization_recommendations()
    profiler.print_summary()

    print("\n" + "=" * 70)
    print("PROFILING COMPLETE")
    print("=" * 70)


if __name__ == '__main__':
    main()
