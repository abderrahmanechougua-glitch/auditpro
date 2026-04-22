"""
Thèmes visuels QSS pour AuditPro.
"""

from core.config import COLORS


LIGHT_COLORS = dict(COLORS)
LIGHT_COLORS.update({
    "module_button_text": "#EE82EE",
    "banner_bg": "#F8F0FF",
    "chip_bg": "#F5F1FA",
    "tab_bg": "#EEE4F7",
    "tab_bg_alt": "#EFE7F7",
    "secondary_hover": "#F0E6F6",
    "drop_hover": "#F8F0FF",
    "drop_active": "#F3E8FF",
})

DIM_COLORS = dict(COLORS)
DIM_COLORS.update({
    "primary": "#305A78",
    "primary_dark": "#203C50",
    "primary_light": "#4D7A97",
    "accent": "#D97706",
    "success": "#2E8B57",
    "warning": "#D97706",
    "danger": "#C2410C",
    "bg_main": "#EEF3F7",
    "bg_panel": "#FCFDFE",
    "bg_sidebar": "#1F3140",
    "text_primary": "#14212B",
    "text_secondary": "#52606D",
    "text_on_dark": "#E7EEF4",
    "border": "#CFD8E3",
    "module_button_text": "#CFE8FF",
    "banner_bg": "#EAF4FB",
    "chip_bg": "#EDF3F8",
    "tab_bg": "#E4EEF6",
    "tab_bg_alt": "#E8F1F8",
    "secondary_hover": "#E4EEF6",
    "drop_hover": "#E8F1F8",
    "drop_active": "#DCEBF7",
})


def _build_stylesheet(c):
    return f"""

/* ── Global ──────────────────────────────────────────── */
QWidget {{
    font-family: "Segoe UI", "Roboto", sans-serif;
    font-size: 13px;
    color: {c['text_primary']};
}}

QMainWindow {{
    background-color: {c['bg_main']};
}}

/* ── Sidebar gauche ──────────────────────────────────── */
#SideBar {{
    background-color: {c['bg_sidebar']};
    min-width: 240px;
}}

#SideBar QLabel#AppTitle {{
    color: {c['text_on_dark']};
    font-size: 18px;
    font-weight: bold;
    padding: 20px 16px 4px 16px;
}}

#SideBar QLabel#AppSubtitle {{
    color: {c['primary_light']};
    font-size: 11px;
    padding: 0px 16px 16px 16px;
}}

#SideBar QLabel#SectionLabel {{
    color: {c['primary_light']};
    font-size: 10px;
    font-weight: bold;
    text-transform: uppercase;
    padding: 16px 16px 6px 16px;
    letter-spacing: 1px;
}}

/* Boutons modules dans la sidebar */
QPushButton#ModuleButton {{
    background-color: transparent;
    color: {c['module_button_text']};
    border: none;
    border-radius: 8px;
    padding: 10px 16px;
    text-align: left;
    font-size: 13px;
}}

QPushButton#ModuleButton:hover {{
    background-color: {c['primary_dark']};
    color: {c['module_button_text']};
}}

QPushButton#ModuleButton:checked {{
    background-color: {c['primary']};
    color: {c['module_button_text']};
    font-weight: bold;
}}

/* ── Zone centrale (Workspace) ───────────────────────── */
#Workspace {{
    background-color: {c['bg_main']};
}}

#WorkspaceTitle {{
    font-size: 22px;
    font-weight: bold;
    color: {c['text_primary']};
    padding: 4px 0px;
}}

#WorkspaceDesc {{
    font-size: 13px;
    color: {c['text_secondary']};
    padding-bottom: 12px;
}}

QLabel#ModuleHintBanner {{
    background-color: {c['banner_bg']};
    border: 1px solid {c['primary_light']};
    border-radius: 8px;
    padding: 8px 12px;
    color: {c['primary_dark']};
    font-size: 12px;
}}

QLabel#DropFeedback {{
    background-color: {c['drop_active']};
    border: 1px solid {c['primary_light']};
    border-radius: 8px;
    padding: 8px 12px;
    color: {c['primary_dark']};
    font-size: 12px;
    font-weight: bold;
}}

QFrame#WorkspaceStateStrip {{
    background-color: {c['bg_panel']};
    border: 1px solid {c['border']};
    border-radius: 10px;
}}

QLabel#WorkspaceStateChip {{
    background-color: {c['chip_bg']};
    border: 1px solid {c['border']};
    border-radius: 8px;
    padding: 4px 8px;
    font-size: 11px;
    color: {c['text_primary']};
}}

QLabel#SectionTitle {{
    font-weight: bold;
    font-size: 14px;
    color: {c['text_primary']};
}}

QTabWidget#AnalysisTabs::pane {{
    border: 1px solid {c['border']};
    background-color: {c['bg_panel']};
    border-radius: 8px;
    top: -1px;
}}

QTabWidget#AnalysisTabs QTabBar::tab {{
    background-color: {c['tab_bg']};
    color: {c['text_primary']};
    border: 1px solid {c['border']};
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    padding: 6px 12px;
    min-width: 90px;
}}

QTabWidget#AnalysisTabs QTabBar::tab:selected {{
    background-color: {c['bg_panel']};
    color: {c['primary_dark']};
    font-weight: bold;
}}

/* Drop zone */
#DropZone {{
    background-color: {c['bg_panel']};
    border: 2px dashed {c['border']};
    border-radius: 12px;
    min-height: 120px;
}}

#DropZone:hover {{
    border-color: {c['primary']};
    background-color: {c['drop_hover']};
}}

#DropZoneLabel {{
    color: {c['text_secondary']};
    font-size: 14px;
}}

/* Cards / Panels */
#Card {{
    background-color: {c['bg_panel']};
    border: 1px solid {c['border']};
    border-radius: 10px;
    padding: 16px;
}}

/* ── Boutons d'action ────────────────────────────────── */
QPushButton#PrimaryButton {{
    background-color: {c['primary']};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px 24px;
    font-size: 14px;
    font-weight: bold;
    min-width: 140px;
}}

QPushButton#PrimaryButton:hover {{
    background-color: {c['primary_dark']};
}}

QPushButton#PrimaryButton:disabled {{
    background-color: {c['border']};
    color: {c['text_secondary']};
}}

QPushButton#SecondaryButton {{
    background-color: transparent;
    color: {c['primary']};
    border: 2px solid {c['primary']};
    border-radius: 8px;
    padding: 9px 24px;
    font-size: 13px;
    font-weight: bold;
}}

QPushButton#SecondaryButton:hover {{
    background-color: {c['secondary_hover']};
}}

QPushButton#DangerButton {{
    background-color: {c['danger']};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 9px 20px;
    font-size: 13px;
}}

/* ── Panneau assistant (droite) ──────────────────────── */
#AssistantPanel {{
    background-color: {c['bg_panel']};
    border-left: 1px solid {c['border']};
    min-width: 260px;
}}

/* ── Shell principal (barre contexte + splitter) ─────── */
QFrame#AuditContextBar {{
    background-color: {c['bg_panel']};
    border-bottom: 1px solid {c['border']};
}}

QLabel#ContextChipModule,
QLabel#ContextChipProfile,
QLabel#ContextChipFile,
QLabel#ContextChipRows,
QLabel#ContextChipDetection,
QLabel#ContextChipStatus {{
    background-color: {c['chip_bg']};
    border: 1px solid {c['border']};
    border-radius: 10px;
    padding: 4px 8px;
    font-size: 11px;
    color: {c['text_primary']};
}}

QLabel#ContextChipStatus {{
    border-color: {c['primary']};
    color: {c['primary_dark']};
    font-weight: bold;
}}

QPushButton#ContextActionPrimary {{
    background-color: {c['primary']};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 6px 12px;
    font-size: 12px;
    font-weight: bold;
}}

QPushButton#ContextActionPrimary:hover {{
    background-color: {c['primary_dark']};
}}

QPushButton#ContextActionPrimary:disabled {{
    background-color: {c['border']};
    color: {c['text_secondary']};
}}

QPushButton#ContextActionSecondary {{
    background-color: transparent;
    color: {c['primary']};
    border: 1px solid {c['primary']};
    border-radius: 8px;
    padding: 6px 12px;
    font-size: 12px;
    font-weight: bold;
}}

QPushButton#ContextActionSecondary:hover {{
    background-color: {c['secondary_hover']};
}}

QPushButton#ContextActionSecondary:disabled {{
    color: {c['text_secondary']};
    border-color: {c['border']};
}}

QFrame#RightRail {{
    background-color: {c['bg_panel']};
    border-left: 1px solid {c['border']};
}}

QTabWidget#RightRailTabs::pane {{
    border: none;
}}

QTabWidget#RightRailTabs QTabBar::tab {{
    background-color: {c['tab_bg_alt']};
    color: {c['text_primary']};
    border: 1px solid {c['border']};
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    padding: 6px 12px;
    min-width: 90px;
}}

QTabWidget#RightRailTabs QTabBar::tab:selected {{
    background-color: {c['bg_panel']};
    color: {c['primary_dark']};
    font-weight: bold;
}}

QSplitter#MainContentSplitter::handle {{
    background-color: {c['border']};
}}

QSplitter#MainContentSplitter::handle:horizontal {{
    width: 2px;
}}

#AssistantTitle {{
    font-size: 14px;
    font-weight: bold;
    color: {c['primary']};
    padding: 16px 16px 8px 16px;
}}

#AssistantMessage {{
    color: {c['text_primary']};
    font-size: 12px;
    padding: 4px 16px;
    line-height: 1.5;
}}

#AssistantSection {{
    color: {c['text_secondary']};
    font-size: 10px;
    font-weight: bold;
    text-transform: uppercase;
    padding: 12px 16px 4px 16px;
}}

/* ── Tableau aperçu ──────────────────────────────────── */
QTableWidget {{
    background-color: {c['bg_panel']};
    border: 1px solid {c['border']};
    border-radius: 6px;
    gridline-color: {c['border']};
    font-size: 11px;
}}

QTableWidget::item {{
    padding: 4px 8px;
}}

QTableWidget QHeaderView::section {{
    background-color: {c['primary']};
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
    color: {c['text_on_dark']};
    border-bottom: 1px solid {c['primary_dark']};
}}

QListWidget#HistoryList::item:hover {{
    background-color: {c['primary_dark']};
}}

/* ── Barre de progression ────────────────────────────── */
QProgressBar {{
    border: none;
    border-radius: 6px;
    background-color: {c['border']};
    text-align: center;
    font-size: 11px;
    min-height: 20px;
}}

QProgressBar::chunk {{
    background-color: {c['primary']};
    border-radius: 6px;
}}

/* ── Inputs ──────────────────────────────────────────── */
QLineEdit, QComboBox, QSpinBox, QDateEdit {{
    border: 1px solid {c['border']};
    border-radius: 6px;
    padding: 8px 12px;
    background-color: {c['bg_panel']};
    font-size: 13px;
}}

QLineEdit:focus, QComboBox:focus {{
    border-color: {c['primary']};
}}

/* ── Scrollbar ───────────────────────────────────────── */
QScrollBar:vertical {{
    width: 8px;
    background: transparent;
}}

QScrollBar::handle:vertical {{
    background: {c['border']};
    border-radius: 4px;
    min-height: 30px;
}}

QScrollBar::handle:vertical:hover {{
    background: {c['text_secondary']};
}}

QScrollBar::add-line, QScrollBar::sub-line {{
    height: 0px;
}}

/* ── Status bar ──────────────────────────────────────── */
QStatusBar {{
    background-color: {c['bg_panel']};
    border-top: 1px solid {c['border']};
    font-size: 11px;
    color: {c['text_secondary']};
}}
"""


STYLESHEET = _build_stylesheet(LIGHT_COLORS)
DIM_STYLESHEET = _build_stylesheet(DIM_COLORS)


def get_stylesheet(theme_name: str = "light") -> str:
    """Retourne la feuille de style selon le thème demandé."""
    if theme_name == "dim":
        return DIM_STYLESHEET
    return STYLESHEET
