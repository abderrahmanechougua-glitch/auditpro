"""
Thème visuel QSS pour AuditPro.
Couleurs violettes et blanches avec accents.
"""

from core.config import COLORS

C = COLORS

STYLESHEET = f"""

/* ── Global ──────────────────────────────────────────── */
QWidget {{
    font-family: "Segoe UI", "Roboto", sans-serif;
    font-size: 13px;
    color: {C['text_primary']};
}}

QMainWindow {{
    background-color: {C['bg_main']};
}}

/* ── Sidebar gauche ──────────────────────────────────── */
#SideBar {{
    background-color: {C['bg_sidebar']};
    min-width: 240px;
    max-width: 240px;
}}

#SideBar QLabel#AppTitle {{
    color: {C['text_on_dark']};
    font-size: 18px;
    font-weight: bold;
    padding: 20px 16px 4px 16px;
}}

#SideBar QLabel#AppSubtitle {{
    color: {C['primary_light']};
    font-size: 11px;
    padding: 0px 16px 16px 16px;
}}

#SideBar QLabel#SectionLabel {{
    color: {C['primary_light']};
    font-size: 10px;
    font-weight: bold;
    text-transform: uppercase;
    padding: 16px 16px 6px 16px;
    letter-spacing: 1px;
}}

/* Boutons modules dans la sidebar */
QPushButton#ModuleButton {{
    background-color: transparent;
    color: #ee82ee;
    border: none;
    border-radius: 8px;
    padding: 10px 16px;
    text-align: left;
    font-size: 13px;
}}

QPushButton#ModuleButton:hover {{
    background-color: {C['primary_dark']};
    color: #ee82ee;
}}

QPushButton#ModuleButton:checked {{
    background-color: {C['primary']};
    color: #ee82ee;
    font-weight: bold;
}}

/* ── Zone centrale (Workspace) ───────────────────────── */
#Workspace {{
    background-color: {C['bg_main']};
}}

#WorkspaceTitle {{
    font-size: 22px;
    font-weight: bold;
    color: {C['text_primary']};
    padding: 4px 0px;
}}

#WorkspaceDesc {{
    font-size: 13px;
    color: {C['text_secondary']};
    padding-bottom: 12px;
}}

/* Drop zone */
#DropZone {{
    background-color: {C['bg_panel']};
    border: 2px dashed {C['border']};
    border-radius: 12px;
    min-height: 120px;
}}

#DropZone:hover {{
    border-color: {C['primary']};
    background-color: #F8F0FF;
}}

#DropZoneLabel {{
    color: {C['text_secondary']};
    font-size: 14px;
}}

/* Cards / Panels */
#Card {{
    background-color: {C['bg_panel']};
    border: 1px solid {C['border']};
    border-radius: 10px;
    padding: 16px;
}}

/* ── Boutons d'action ────────────────────────────────── */
QPushButton#PrimaryButton {{
    background-color: {C['primary']};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px 24px;
    font-size: 14px;
    font-weight: bold;
    min-width: 140px;
}}

QPushButton#PrimaryButton:hover {{
    background-color: {C['primary_dark']};
}}

QPushButton#PrimaryButton:disabled {{
    background-color: {C['border']};
    color: {C['text_secondary']};
}}

QPushButton#SecondaryButton {{
    background-color: transparent;
    color: {C['primary']};
    border: 2px solid {C['primary']};
    border-radius: 8px;
    padding: 9px 24px;
    font-size: 13px;
    font-weight: bold;
}}

QPushButton#SecondaryButton:hover {{
    background-color: #F0E6F6;
}}

QPushButton#DangerButton {{
    background-color: {C['danger']};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 9px 20px;
    font-size: 13px;
}}

/* ── Panneau assistant (droite) ──────────────────────── */
#AssistantPanel {{
    background-color: {C['bg_panel']};
    border-left: 1px solid {C['border']};
    min-width: 280px;
    max-width: 280px;
}}

#AssistantTitle {{
    font-size: 14px;
    font-weight: bold;
    color: {C['primary']};
    padding: 16px 16px 8px 16px;
}}

#AssistantMessage {{
    color: {C['text_primary']};
    font-size: 12px;
    padding: 4px 16px;
    line-height: 1.5;
}}

#AssistantSection {{
    color: {C['text_secondary']};
    font-size: 10px;
    font-weight: bold;
    text-transform: uppercase;
    padding: 12px 16px 4px 16px;
}}

/* ── Tableau aperçu ──────────────────────────────────── */
QTableWidget {{
    background-color: {C['bg_panel']};
    border: 1px solid {C['border']};
    border-radius: 6px;
    gridline-color: {C['border']};
    font-size: 11px;
}}

QTableWidget::item {{
    padding: 4px 8px;
}}

QTableWidget QHeaderView::section {{
    background-color: {C['primary']};
    color: white;
    font-weight: bold;
    font-size: 11px;
    padding: 6px 8px;
    border: none;
}}

/* ── Historique ──────────────────────────────────────── */
QListWidget#HistoryList {{
    background-color: transparent;
    border: none;
    font-size: 11px;
}}

QListWidget#HistoryList::item {{
    padding: 6px 16px;
    color: {C['text_primary']};
    border-bottom: 1px solid {C['border']};
}}

QListWidget#HistoryList::item:hover {{
    background-color: #F0E6F6;
}}

/* ── Barre de progression ────────────────────────────── */
QProgressBar {{
    border: none;
    border-radius: 6px;
    background-color: {C['border']};
    text-align: center;
    font-size: 11px;
    min-height: 20px;
}}

QProgressBar::chunk {{
    background-color: {C['primary']};
    border-radius: 6px;
}}

/* ── Inputs ──────────────────────────────────────────── */
QLineEdit, QComboBox, QSpinBox, QDateEdit {{
    border: 1px solid {C['border']};
    border-radius: 6px;
    padding: 8px 12px;
    background-color: {C['bg_panel']};
    font-size: 13px;
}}

QLineEdit:focus, QComboBox:focus {{
    border-color: {C['primary']};
}}

/* ── Scrollbar ───────────────────────────────────────── */
QScrollBar:vertical {{
    width: 8px;
    background: transparent;
}}

QScrollBar::handle:vertical {{
    background: {C['border']};
    border-radius: 4px;
    min-height: 30px;
}}

QScrollBar::handle:vertical:hover {{
    background: {C['text_secondary']};
}}

QScrollBar::add-line, QScrollBar::sub-line {{
    height: 0px;
}}

/* ── Status bar ──────────────────────────────────────── */
QStatusBar {{
    background-color: {C['bg_panel']};
    border-top: 1px solid {C['border']};
    font-size: 11px;
    color: {C['text_secondary']};
}}
"""
