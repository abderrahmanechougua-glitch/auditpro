"""
Tableau d'aperçu des données avant exécution.
"""
from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView
from PyQt6.QtCore import Qt
import pandas as pd


class PreviewTable(QTableWidget):
    """Affiche un DataFrame dans un QTableWidget formaté."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.verticalHeader().setDefaultSectionSize(28)
        self.horizontalHeader().setStretchLastSection(True)

    def load_dataframe(self, df: pd.DataFrame, max_rows: int = 10):
        """Charge un DataFrame et l'affiche."""
        if df is None or df.empty:
            self.setRowCount(0)
            self.setColumnCount(0)
            return

        display_df = df.head(max_rows)

        self.setRowCount(len(display_df))
        self.setColumnCount(len(display_df.columns))
        self.setHorizontalHeaderLabels([str(c) for c in display_df.columns])

        for i in range(len(display_df)):
            for j in range(len(display_df.columns)):
                val = display_df.iat[i, j]
                text = "" if pd.isna(val) else str(val)
                item = QTableWidgetItem(text)

                # Aligner les nombres à droite
                if isinstance(val, (int, float)):
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                    if isinstance(val, float):
                        item.setText(f"{val:,.2f}")

                self.setItem(i, j, item)

        # Auto-resize colonnes
        self.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )

    def clear_table(self):
        self.setRowCount(0)
        self.setColumnCount(0)
