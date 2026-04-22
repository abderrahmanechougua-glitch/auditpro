"""
Thèmes visuels QSS pour AuditPro.
"""

from core.config import COLORS


LIGHT_COLORS = dict(COLORS)
LIGHT_COLORS.update({
    "module_button_text": "#4B5563",
    "banner_bg": "#EEF2FF",
    "chip_bg": "#F3F4F6",
    "tab_bg": "#F3F4F6",
    "tab_bg_alt": "#F8FAFC",
    "secondary_hover": "#EEF2FF",
    "drop_hover": "#F8FAFF",
    "drop_active": "#EEF2FF",
})

DIM_COLORS = dict(COLORS)
DIM_COLORS.update({
    "primary": "#B882EE",
    "primary_dark": "#C86FD0",
    "primary_light": "#F2A5F2",
    "accent": "#FDF4FF",
    "success": "#10B981",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "bg_main": "#F3F4F6",
    "bg_panel": "#FFFFFF",
    "bg_sidebar": "#F9FAFB",
    "text_primary": "#111827",
    "text_secondary": "#6B7280",
    "text_on_dark": "#111827",
    "border": "#D1D5DB",
    "module_button_text": "#374151",
    "banner_bg": "#E0E7FF",
    "chip_bg": "#E5E7EB",
    "tab_bg": "#E5E7EB",
    "tab_bg_alt": "#ECEFF3",
    "secondary_hover": "#FDF4FF",
    "drop_hover": "#FDF4FF",
    "drop_active": "#FAE8FF",
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
    color: {c['text_primary']};
    font-size: 18px;
    font-weight: bold;
    padding: 20px 16px 4px 16px;
}}

#SideBar QLabel#AppSubtitle {{
    color: {c['text_secondary']};
    font-size: 11px;
    padding: 0px 16px 16px 16px;
}}

#SideBar QLabel#SectionLabel {{
    color: {c['text_secondary']};
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
    min-height: 18px;
    text-align: left;
    font-size: 13px;
}}

QPushButton#ModuleButton:hover {{
    background-color: {c['accent']};
    color: {c['text_primary']};
}}

QPushButton#ModuleButton:checked {{
    background-color: {c['primary']};
    color: white;
    font-weight: bold;
    border-left: 3px solid {c['primary_dark']};
}}

QPushButton#ModuleButton:focus {{
    border: 2px solid {c['primary_light']};
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

QPushButton#StepButton {{
    background-color: #F8FAFC;
    color: {c['text_primary']};
    border: 1px solid {c['border']};
    border-radius: 8px;
    padding: 8px 16px;
    text-align: left;
    font-size: 13px;
}}

QPushButton#StepButton:hover {{
    background-color: {c['accent']};
    border-color: {c['primary']};
}}

QPushButton#StepButton:pressed {{
    background-color: {c['drop_active']};
}}

QPushButton#StepButton:focus {{
    border: 2px solid {c['primary_light']};
}}

QPushButton#StepButton:checked {{
    background-color: {c['primary']};
    color: white;
    border-color: {c['primary_dark']};
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
    font-size: 15px;
    color: {c['text_primary']};
}}

QLabel#InputLabel {{
    font-weight: 600;
    font-size: 12px;
    color: {c['text_primary']};
    padding-top: 2px;
}}

QLabel#InputTip {{
    color: {c['text_secondary']};
    font-size: 11px;
    margin-left: 8px;
    padding-bottom: 4px;
}}

QLabel#PreviewMeta {{
    background-color: {c['chip_bg']};
    border: 1px solid {c['border']};
    border-radius: 8px;
    padding: 2px 8px;
    font-size: 11px;
    color: {c['text_secondary']};
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
    min-height: 38px;
}}

QPushButton#PrimaryButton:hover {{
    background-color: {c['primary_dark']};
}}

QPushButton#PrimaryButton:pressed {{
    background-color: {c['primary_dark']};
    padding-top: 11px;
    padding-bottom: 9px;
}}

QPushButton#PrimaryButton:focus {{
    border: 2px solid {c['primary_light']};
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
    min-height: 36px;
}}

QPushButton#SecondaryButton:hover {{
    background-color: {c['secondary_hover']};
}}

QPushButton#SecondaryButton:pressed {{
    background-color: {c['drop_active']};
}}

QPushButton#SecondaryButton:focus {{
    border-color: {c['primary_dark']};
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
    border-radius: 8px;
    gridline-color: {c['border']};
    font-size: 12px;
    alternate-background-color: {c['chip_bg']};
    selection-background-color: {c['accent']};
    selection-color: {c['text_primary']};
}}

QTableWidget::item:hover {{
    background-color: {c['drop_hover']};
}}

QTableWidget::item {{
    padding: 4px 8px;
}}

QTableWidget QHeaderView::section {{
    background-color: {c['primary']};
    color: white;
    font-weight: bold;
    font-size: 12px;
    padding: 8px 10px;
    border: none;
}}

QFrame#ResultPanel {{
    background-color: {c['bg_panel']};
    border: 1px solid {c['border']};
    border-radius: 10px;
}}

QFrame#ResultPanel[state="success"] {{
    background-color: #ECFDF5;
    border: 1px solid #A7F3D0;
}}

QFrame#ResultPanel[state="warning"] {{
    background-color: #FFFBEB;
    border: 1px solid #FDE68A;
}}

QFrame#ResultPanel[state="error"] {{
    background-color: #FEF2F2;
    border: 1px solid #FECACA;
}}

QFrame#ResultPanel[state="success"] QLabel#ResultTitle {{
    color: #047857;
}}

QFrame#ResultPanel[state="warning"] QLabel#ResultTitle {{
    color: #B45309;
}}

QFrame#ResultPanel[state="error"] QLabel#ResultTitle {{
    color: #B91C1C;
}}

QLabel#ResultTitle {{
    font-size: 14px;
    font-weight: 700;
    color: {c['text_primary']};
}}

QLabel#ResultBody {{
    font-size: 12px;
    color: {c['text_primary']};
    line-height: 1.4;
}}

QLabel#ResultMeta {{
    font-size: 11px;
    color: {c['text_secondary']};
    border-top: 1px solid {c['border']};
    padding-top: 6px;
}}

/* ── Historique ──────────────────────────────────────── */
QListWidget#HistoryList {{
    background-color: transparent;
    border: none;
    font-size: 11px;
}}

QListWidget#HistoryList::item {{
    padding: 6px 16px;
    color: {c['text_primary']};
    border-bottom: 1px solid {c['border']};
}}

QListWidget#HistoryList::item:hover {{
    background-color: {c['accent']};
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
    min-height: 20px;
}}

QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDateEdit:focus {{
    border-color: {c['primary']};
    background-color: #FFFFFF;
}}

QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled, QDateEdit:disabled {{
    background-color: #F9FAFB;
    color: {c['text_secondary']};
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
