"""
DRK Dienstkleidung - QSS Stylesheet
Farben: DRK-Blau #2F4B5D, Weiß, Hellgrau
"""

MAIN_STYLE = """
/* === Allgemein === */
QMainWindow, QDialog {
    background-color: #F0F2F5;
}

QWidget {
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 13px;
    color: #333333;
}

/* === Sidebar === */
QWidget#sidebar {
    background-color: #2F4B5D;
    min-width: 240px;
    max-width: 240px;
}

QLabel#sidebar_title {
    color: white;
    font-size: 14px;
    font-weight: bold;
    padding: 18px 16px 4px 16px;
}

QLabel#sidebar_subtitle {
    color: rgba(255,255,255,0.65);
    font-size: 10px;
    padding: 0 16px 14px 16px;
}

QFrame#sidebar_divider {
    background-color: rgba(255,255,255,0.2);
    max-height: 1px;
    margin: 0 12px 8px 12px;
}

QPushButton#nav_btn {
    background-color: transparent;
    color: rgba(255,255,255,0.8);
    border: none;
    text-align: left;
    padding: 11px 16px;
    font-size: 13px;
    border-radius: 0;
    border-left: 3px solid transparent;
}

QPushButton#nav_btn:hover {
    background-color: rgba(255,255,255,0.12);
    color: white;
}

QPushButton#nav_btn:checked {
    background-color: rgba(255,255,255,0.18);
    color: white;
    font-weight: bold;
    border-left: 3px solid white;
}

QLabel#sidebar_version {
    color: rgba(255,255,255,0.4);
    font-size: 10px;
    padding: 8px 16px;
}

/* === Content Bereich === */
QWidget#content_area {
    background-color: #F0F2F5;
}

/* === Seitentitel === */
QLabel#page_title {
    font-size: 22px;
    font-weight: bold;
    color: #2F4B5D;
    padding-bottom: 2px;
}

QLabel#page_subtitle {
    font-size: 12px;
    color: #888888;
    padding-bottom: 12px;
}

/* === Stat-Karten === */
QFrame#stat_card {
    background-color: white;
    border-radius: 8px;
    border: 1px solid #E8E8E8;
}

QLabel#stat_value {
    font-size: 30px;
    font-weight: bold;
    color: #2F4B5D;
}

QLabel#stat_label {
    font-size: 11px;
    color: #888888;
    font-weight: bold;
    text-transform: uppercase;
}

QFrame#stat_card_warn {
    background-color: #FFF8E1;
    border-radius: 8px;
    border: 1px solid #FFE082;
}

QLabel#stat_value_warn {
    font-size: 30px;
    font-weight: bold;
    color: #F57C00;
}

/* === Tabellen === */
QTableWidget {
    background-color: white;
    gridline-color: #F0F0F0;
    border: 1px solid #E0E0E0;
    border-radius: 6px;
    selection-background-color: #2F4B5D;
    selection-color: white;
    alternate-background-color: #FAFAFA;
    outline: none;
}

QTableWidget::item {
    padding: 7px 12px;
    border-bottom: 1px solid #F5F5F5;
}

QTableWidget::item:selected {
    background-color: #2F4B5D;
    color: white;
}

QHeaderView::section {
    background-color: #2F4B5D;
    color: white;
    padding: 9px 12px;
    border: none;
    border-right: 1px solid rgba(255,255,255,0.1);
    font-weight: bold;
    font-size: 12px;
}

QHeaderView::section:last {
    border-right: none;
}

/* === Buttons === */
QPushButton#btn_primary {
    background-color: #2F4B5D;
    color: white;
    border: none;
    border-radius: 5px;
    padding: 9px 22px;
    font-size: 13px;
    font-weight: bold;
    min-width: 110px;
}

QPushButton#btn_primary:hover {
    background-color: #3A5E72;
}

QPushButton#btn_primary:pressed {
    background-color: #1E3444;
}

QPushButton#btn_primary:disabled {
    background-color: #BBBBBB;
    color: #F5F5F5;
}

QPushButton#btn_secondary {
    background-color: white;
    color: #2F4B5D;
    border: 2px solid #2F4B5D;
    border-radius: 5px;
    padding: 7px 20px;
    font-size: 13px;
    min-width: 100px;
}

QPushButton#btn_secondary:hover {
    background-color: #DCE8F0;
}

QPushButton#btn_secondary:pressed {
    background-color: #C4D8E8;
}

QPushButton#btn_danger {
    background-color: white;
    color: #CC3300;
    border: 2px solid #CC3300;
    border-radius: 5px;
    padding: 7px 20px;
    font-size: 13px;
}

QPushButton#btn_danger:hover {
    background-color: #CC3300;
    color: white;
}

QPushButton#btn_icon {
    background-color: transparent;
    color: #2F4B5D;
    border: none;
    padding: 4px 8px;
    font-size: 12px;
    border-radius: 3px;
}

QPushButton#btn_icon:hover {
    background-color: #DCE8F0;
}

/* === Formularelemente === */
QComboBox, QLineEdit, QDateEdit, QSpinBox, QTextEdit {
    border: 1px solid #CCCCCC;
    border-radius: 5px;
    padding: 7px 10px;
    background-color: white;
    font-size: 13px;
    min-height: 34px;
    selection-background-color: #2F4B5D;
}

QComboBox:focus, QLineEdit:focus, QDateEdit:focus, QSpinBox:focus, QTextEdit:focus {
    border: 2px solid #2F4B5D;
    padding: 6px 9px;
}

QComboBox::drop-down {
    border: none;
    padding-right: 4px;
    width: 24px;
}

QComboBox::down-arrow {
    image: none;
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 7px solid #555555;
}

QComboBox::down-arrow:on {
    border-top: none;
    border-bottom: 7px solid #2F4B5D;
}

QDateEdit::drop-down, QSpinBox::drop-down {
    border: none;
    width: 22px;
}

QDateEdit::down-arrow, QSpinBox::down-arrow {
    image: none;
    width: 0;
    height: 0;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 6px solid #555555;
}

QDateEdit::up-arrow, QSpinBox::up-arrow {
    image: none;
    width: 0;
    height: 0;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-bottom: 6px solid #555555;
}

QDateEdit::up-button, QDateEdit::down-button,
QSpinBox::up-button, QSpinBox::down-button {
    width: 20px;
    border: none;
    background-color: transparent;
}

/* === GroupBox === */
QGroupBox {
    font-weight: bold;
    font-size: 13px;
    color: #555555;
    border: 1px solid #E0E0E0;
    border-radius: 6px;
    margin-top: 14px;
    padding-top: 14px;
    background-color: white;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 8px;
    color: #2F4B5D;
}

/* === Statusleiste === */
QStatusBar {
    background-color: #EAEAEA;
    color: #777777;
    font-size: 11px;
    border-top: 1px solid #E0E0E0;
}

/* === Tabs === */
QTabWidget::pane {
    border: 1px solid #E0E0E0;
    border-radius: 0 6px 6px 6px;
    background-color: white;
    padding: 16px;
}

QTabBar::tab {
    background-color: #E8E8E8;
    color: #666666;
    padding: 9px 24px;
    border: 1px solid #E0E0E0;
    border-bottom: none;
    border-radius: 5px 5px 0 0;
    font-size: 13px;
    margin-right: 2px;
}

QTabBar::tab:selected {
    background-color: white;
    color: #2F4B5D;
    font-weight: bold;
}

QTabBar::tab:hover:!selected {
    background-color: #F0F0F0;
    color: #2F4B5D;
}

/* === ScrollBar === */
QScrollBar:vertical {
    border: none;
    background-color: #F5F5F5;
    width: 8px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background-color: #CCCCCC;
    border-radius: 4px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background-color: #2F4B5D;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}

QScrollBar:horizontal {
    border: none;
    background-color: #F5F5F5;
    height: 8px;
}

QScrollBar::handle:horizontal {
    background-color: #CCCCCC;
    border-radius: 4px;
    min-width: 30px;
}

/* === Trennlinie === */
QFrame[frameShape="4"], QFrame[frameShape="5"] {
    color: #E0E0E0;
}

/* === Label in Formular === */
QLabel#form_label {
    font-size: 12px;
    color: #555555;
    font-weight: bold;
}

/* === Info-Banner === */
QFrame#info_banner {
    background-color: #E3F2FD;
    border-radius: 6px;
    border-left: 4px solid #1976D2;
    padding: 8px;
}

QFrame#warn_banner {
    background-color: #FFF3E0;
    border-radius: 6px;
    border-left: 4px solid #F57C00;
    padding: 8px;
}

QFrame#success_banner {
    background-color: #E8F5E9;
    border-radius: 6px;
    border-left: 4px solid #388E3C;
    padding: 8px;
}
"""
