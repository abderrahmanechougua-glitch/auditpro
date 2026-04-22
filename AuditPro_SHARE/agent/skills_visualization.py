"""
Data Visualization Skill Implementation for AuditPro modules.

Provides chart and visualization capabilities:
- TVA trends visualization
- GL reconciliation charts
- Invoice distribution analysis
- Audit summary visualizations
"""
from __future__ import annotations
from pathlib import Path
from typing import Optional, Dict, Any, List
import logging
import pandas as pd

logger = logging.getLogger(__name__)


class DataVisualizationSkillHandler:
    """Handles data visualization for audit modules."""
    
    def __init__(self):
        self.matplotlib_available = True
        self.plotly_available = True
    
    def create_bar_chart(
        self,
        data: Dict | pd.DataFrame,
        title: str,
        x_label: str = "",
        y_label: str = "",
        output_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Create a bar chart for audit data visualization.
        
        Args:
            data: Data for chart (dict or DataFrame)
            title: Chart title
            x_label: X-axis label
            y_label: Y-axis label
            output_path: Optional path to save chart as image
        
        Returns:
            Dict with chart metadata and data
        """
        result = {
            "success": False,
            "chart_type": "bar",
            "title": title,
            "output_path": output_path
        }
        
        try:
            import matplotlib.pyplot as plt
            
            # Convert data to DataFrame if needed
            if isinstance(data, dict):
                df = pd.DataFrame(list(data.items()), columns=[x_label or "Category", y_label or "Value"])
            else:
                df = data
            
            # Create figure
            fig, ax = plt.subplots(figsize=(12, 6))
            df.plot(kind='bar', ax=ax, color='steelblue')
            
            ax.set_title(title, fontsize=14, fontweight='bold')
            ax.set_xlabel(x_label, fontsize=12)
            ax.set_ylabel(y_label, fontsize=12)
            plt.tight_layout()
            
            # Save if path provided
            if output_path:
                Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                plt.savefig(output_path, dpi=300, bbox_inches='tight')
                result["output_path"] = str(output_path)
            
            plt.close()
            result["success"] = True
            
        except Exception as e:
            logger.error(f"Error creating bar chart: {e}")
            result["error"] = str(e)
        
        return result
    
    def create_line_chart(
        self,
        data: pd.DataFrame,
        title: str,
        x_col: str,
        y_cols: List[str],
        output_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Create line chart for trend analysis (TVA, GL balances over time).
        
        Args:
            data: DataFrame with data
            title: Chart title
            x_col: Column for X-axis (e.g., month)
            y_cols: Column(s) for Y-axis values
            output_path: Optional path to save
        
        Returns:
            Dict with chart metadata
        """
        result = {
            "success": False,
            "chart_type": "line",
            "title": title,
            "output_path": output_path
        }
        
        try:
            import matplotlib.pyplot as plt
            
            fig, ax = plt.subplots(figsize=(12, 6))
            
            for y_col in y_cols:
                ax.plot(data[x_col], data[y_col], marker='o', label=y_col, linewidth=2)
            
            ax.set_title(title, fontsize=14, fontweight='bold')
            ax.set_xlabel(x_col, fontsize=12)
            ax.set_ylabel("Value", fontsize=12)
            ax.legend()
            ax.grid(True, alpha=0.3)
            plt.tight_layout()
            
            if output_path:
                Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                plt.savefig(output_path, dpi=300, bbox_inches='tight')
                result["output_path"] = str(output_path)
            
            plt.close()
            result["success"] = True
            
        except Exception as e:
            logger.error(f"Error creating line chart: {e}")
            result["error"] = str(e)
        
        return result
    
    def create_pie_chart(
        self,
        data: Dict[str, float],
        title: str,
        output_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Create pie chart for distribution analysis (invoice by supplier, etc.).
        
        Args:
            data: Dict of {label: value}
            title: Chart title
            output_path: Optional path to save
        
        Returns:
            Dict with chart metadata
        """
        result = {
            "success": False,
            "chart_type": "pie",
            "title": title,
            "output_path": output_path
        }
        
        try:
            import matplotlib.pyplot as plt
            
            fig, ax = plt.subplots(figsize=(10, 8))
            
            labels = list(data.keys())
            values = list(data.values())
            
            ax.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
            ax.set_title(title, fontsize=14, fontweight='bold')
            
            plt.tight_layout()
            
            if output_path:
                Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                plt.savefig(output_path, dpi=300, bbox_inches='tight')
                result["output_path"] = str(output_path)
            
            plt.close()
            result["success"] = True
            
        except Exception as e:
            logger.error(f"Error creating pie chart: {e}")
            result["error"] = str(e)
        
        return result
    
    def create_comparison_chart(
        self,
        data_n: pd.Series,
        data_n_minus_1: pd.Series,
        title: str,
        output_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Create side-by-side comparison chart (N vs N-1).
        
        Args:
            data_n: Current year data
            data_n_minus_1: Prior year data
            title: Chart title
            output_path: Optional path to save
        
        Returns:
            Dict with chart metadata
        """
        result = {
            "success": False,
            "chart_type": "comparison",
            "title": title,
            "output_path": output_path
        }
        
        try:
            import matplotlib.pyplot as plt
            import numpy as np
            
            fig, ax = plt.subplots(figsize=(12, 6))
            
            x = np.arange(len(data_n))
            width = 0.35
            
            bars1 = ax.bar(x - width/2, data_n, width, label='N', color='steelblue')
            bars2 = ax.bar(x + width/2, data_n_minus_1, width, label='N-1', color='lightcoral')
            
            ax.set_title(title, fontsize=14, fontweight='bold')
            ax.set_ylabel('Amount', fontsize=12)
            ax.legend()
            ax.grid(True, alpha=0.3, axis='y')
            
            plt.tight_layout()
            
            if output_path:
                Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                plt.savefig(output_path, dpi=300, bbox_inches='tight')
                result["output_path"] = str(output_path)
            
            plt.close()
            result["success"] = True
            
        except Exception as e:
            logger.error(f"Error creating comparison chart: {e}")
            result["error"] = str(e)
        
        return result
    
    def generate_summary_stats(
        self,
        df: pd.DataFrame,
        numeric_cols: List[str]
    ) -> Dict[str, Any]:
        """
        Generate summary statistics for visualization context.
        
        Args:
            df: DataFrame to analyze
            numeric_cols: Columns to analyze
        
        Returns:
            Dict with summary statistics
        """
        result = {
            "success": False,
            "statistics": {}
        }
        
        try:
            for col in numeric_cols:
                if col in df.columns:
                    col_data = pd.to_numeric(df[col], errors='coerce')
                    result["statistics"][col] = {
                        "count": int(col_data.count()),
                        "mean": float(col_data.mean()),
                        "std": float(col_data.std()),
                        "min": float(col_data.min()),
                        "max": float(col_data.max()),
                        "median": float(col_data.median()),
                        "sum": float(col_data.sum())
                    }
            
            result["success"] = True
            
        except Exception as e:
            logger.error(f"Error generating summary stats: {e}")
            result["error"] = str(e)
        
        return result


# Singleton instance
_viz_handler = DataVisualizationSkillHandler()


def get_visualization_skill_handler() -> DataVisualizationSkillHandler:
    """Get the data visualization skill handler instance."""
    return _viz_handler
