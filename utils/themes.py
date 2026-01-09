LIGHT_THEME = """
    QMainWindow, QDialog {
        background-color: #f8fafc;
    }
    QWidget#CentralWidget {
        background-color: #f8fafc;
    }
    QFrame#HeaderFrame {
        background: white;
        border-bottom: 1px solid #e2e8f0;
        border-bottom-left-radius: 15px;
        border-bottom-right-radius: 15px;
    }
    QLabel {
        color: #1e293b;
        font-family: 'Microsoft YaHei UI';
    }
    QLabel[class="task-title"] {
        color: #1e293b;
        font-weight: bold;
    }
    QLabel[class="task-meta"] {
        color: #64748b;
    }
    QScrollArea {
        background: transparent;
        border: none;
    }
    QWidget#ScrollContent {
        background: transparent;
    }
    /* 任务卡片样式 */
    QFrame[class="TaskCard"] {
        background: white;
        border: 1px solid #e2e8f0;
    }
    QFrame#UploadArea {
        border: 3px dashed #cbd5e1;
        border-radius: 10px;
        background: white;
    }
    QFrame#UploadArea:hover {
        border-color: #2563eb;
        background: #f8fafc;
    }
    QFrame[class="FileQueueItem"] {
        background-color: white;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px;
        margin: 2px 0;
    }
    QFrame[class="FileQueueItem"] QPushButton {
        background: #f1f5f9;
        border: 1px solid #e2e8f0;
        color: #64748b;
    }
    QFrame[class="FileQueueItem"] QPushButton:hover {
        background: #fee2e2;
        color: #ef4444;
        border-color: #fca5a5;
    }
    QFrame[class="FileQueueItem"]:hover {
        border-color: #2563eb;
        background-color: #f8fafc;
    }
    QLabel#LogPanel {
        background: #f8fafc;
        color: #1e293b;
        border-right: none;
        padding: 12px 16px;
        border-radius: 6px;
        border: 1px solid #e2e8f0;
        border-left: 4px solid #10b981;
    }
    /* 步骤图标下方的文字样式 */
    QLabel[class="step-label"] {
        font-size: 11px;
        color: #64748b;
        padding: 2px 4px;
        border-radius: 4px;
    }
    QLabel[class="step-label"][status="processing"] { color: #2563eb; font-weight: bold; }
    QLabel[class="step-label"][status="done"] { color: #10b981; font-weight: bold; }
    QLabel[class="step-label"][status="finished"] { color: #3b82f6; font-weight: bold; }
    QLabel[class="step-label"][status="error"] { color: #ef4444; font-weight: bold; }

    /* 表格样式 */
    QTableWidget {
        background-color: white;
        gridline-color: #e2e8f0;
        border: 1px solid #e2e8f0;
        color: #1e293b;
    }
    QHeaderView::section {
        background-color: #f1f5f9;
        color: #475569;
        padding: 4px;
        border: 1px solid #e2e8f0;
    }
    QGroupBox {
        font-weight: bold;
        color: #0ea5e9;
        border: 2px solid #0ea5e9;
        border-radius: 10px;
        margin-top: 12px;
        padding-top: 10px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 15px;
        top: 2px;
        padding: 0 5px;
        background-color: #f8fafc;
    }
    QLineEdit, QComboBox {
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        padding: 5px;
        background: white;
        color: #1e293b;
    }
    QPushButton {
        border-radius: 6px;
        padding: 8px 16px;
        color: #1e293b;
        background: #f1f5f9;
        border: 1px solid #e2e8f0;
    }
    QPushButton#UploadBtn {
        background: #2563eb;
        color: white;
        font-weight: bold;
    }
    QPushButton#UploadBtn:hover {
        background: #1d4ed8;
    }
    QPushButton#ThemeBtn {
        background: #f1f5f9;
        color: #475569;
        border: 1px solid #cbd5e1;
    }
    QPushButton#ThemeBtn:hover {
        background: #e2e8f0;
    }

    /* 复选框/单选框样式 */
    QCheckBox, QRadioButton {
        spacing: 8px;
        color: #1e293b;
    }
    QCheckBox::indicator, QRadioButton::indicator {
        width: 18px;
        height: 18px;
        border: 1px solid #cbd5e1;
        background: white;
    }
    QCheckBox::indicator { border-radius: 4px; }
    QRadioButton::indicator { border-radius: 9px; }
    QCheckBox::indicator:checked {
        background: #8b5cf6;
        border-color: #7c4dff;
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjQiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBvbHlsaW5lIHBvaW50cz0iMjAgNiA5IDE3IDQgMTIiPjwvcG9seWxpbmU+PC9zdmc+);
    }
    QRadioButton::indicator:checked {
        background: #10b981;
        border-color: #059669;
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJ3aGl0ZSI+PGNpcmNsZSBjeD0iMTIiIGN5PSIxMiIgcj0iNiI+PC9jaXJjbGU+PC9zdmc+);
    }

    /* 下拉框进阶样式 */
    QComboBox::drop-down {
        subcontrol-origin: padding;
        subcontrol-position: center right;
        width: 24px;
        border-left: none;
    }
    QComboBox::down-arrow {
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiM2NDc0OGIiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cG9seWxpbmUgcG9pbnRzPSI2IDkgMTIgMTUgMTggOSI+PC9wb2x5bGluZT48L3N2Zz4=);
        width: 14px;
        height: 14px;
    }
    QComboBox QAbstractItemView {
        background-color: white;
        color: #1e293b;
        border: 1px solid #cbd5e1;
        selection-background-color: #f1f5f9;
        outline: none;
    }
"""

DARK_THEME = """
    QMainWindow, QDialog {
        background-color: #1f2937;
    }
    QWidget#CentralWidget {
        background-color: #1f2937;
    }
    QFrame#HeaderFrame {
        background: #111827;
        border-bottom: 1px solid #374151;
        border-bottom-left-radius: 15px;
        border-bottom-right-radius: 15px;
    }
    QLabel {
        color: #f3f4f6;
        font-family: 'Microsoft YaHei UI';
    }
    QLabel[class="task-title"] {
        color: #f9fafb;
        font-weight: bold;
    }
    QLabel[class="task-meta"] {
        color: #9ca3af;
    }
    /* 修正滚动区域背景 */
    QScrollArea {
        background: transparent;
        border: none;
    }
    QScrollArea > QWidget > QWidget {
        background-color: transparent;
    }
    QWidget#ScrollContent {
        background: transparent;
    }
    /* 任务卡片深色适配 */
    QFrame[class="TaskCard"] {
        background: transparent;
        border: 1px solid #374151;
    }
    QFrame#UploadArea {
        border: 3px dashed #4b5563;
        border-radius: 10px;
        background: #111827;
    }
    QFrame#UploadArea:hover {
        border-color: #3b82f6;
        background: #0f172a;
    }
    QFrame[class="FileQueueItem"] {
        background-color: #111827;
        border: 1px solid #374151;
        border-radius: 8px;
        padding: 12px;
        margin: 2px 0;
    }
    QFrame[class="FileQueueItem"] QPushButton {
        background: #374151;
        border: 1px solid #4b5563;
        color: #9ca3af;
    }
    QFrame[class="FileQueueItem"] QPushButton:hover {
        background: #4c0519;
        color: #fb7185;
        border-color: #fb7185;
    }
    QFrame[class="FileQueueItem"]:hover {
        border-color: #3b82f6;
        background-color: #1f2937;
    }
    QFrame[class="TaskCard"] QLabel {
        color: #f3f4f6;
    }
    QFrame[class="TaskCard"] #LogPanel {
        background: #111827;
        color: #d1d5db;
        padding: 12px 16px;
        border-radius: 6px;
        border: 1px solid #374151;
        border-left: 4px solid #059669;
    }
    /* 步骤图标下方的文字样式 - 深色背景增强 */
    QLabel[class="step-label"] {
        font-size: 11px;
        color: #9ca3af;
        background: #1f2937; /* 深色背景 */
        padding: 3px 6px;
        border-radius: 4px;
    }
    QLabel[class="step-label"][status="processing"] { color: #60a5fa; font-weight: bold; background: #1e3a8a; }
    QLabel[class="step-label"][status="done"] { color: #34d399; font-weight: bold; background: #064e3b; }
    QLabel[class="step-label"][status="finished"] { color: #93c5fd; font-weight: bold; background: #1e3a8a; }
    QLabel[class="step-label"][status="error"] { color: #fb7185; font-weight: bold; background: #4c0519; }

    /* 表格样式 */
    QTableWidget {
        background-color: #374151;
        gridline-color: #4b5563;
        border: 1px solid #4b5563;
        color: #f3f4f6;
    }
    QTableWidget::item {
        color: #f3f4f6;
    }
    QHeaderView::section {
        background-color: #4b5563;
        color: #f9fafb;
        padding: 4px;
        border: 1px solid #374151;
    }
    QGroupBox {
        font-weight: bold;
        color: #38bdf8;
        border: 2px solid #38bdf8;
        border-radius: 10px;
        margin-top: 12px;
        padding-top: 10px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 15px;
        top: 2px;
        padding: 0 5px;
        background-color: #1f2937;
    }
    QLineEdit, QComboBox {
        border: 1px solid #4b5563;
        border-radius: 6px;
        padding: 5px;
        background: #111827;
        color: #f3f4f6;
    }
    QPushButton {
        border-radius: 6px;
        padding: 8px 16px;
        color: #f3f4f6;
        background: #374151;
        border: 1px solid #4b5563;
    }
    QPushButton#UploadBtn {
        background: #2563eb;
        color: white;
        font-weight: bold;
    }
    QPushButton#UploadBtn:hover {
        background: #1d4ed8;
    }
    QTextBrowser {
        background-color: #111827;
        color: #f3f4f6;
        border: 1px solid #374151;
    }
    QPushButton#ThemeBtn {
        background: #374151;
        color: #d1d5db;
        border: 1px solid #4b5563;
    }
    QPushButton#ThemeBtn:hover {
        background: #4b5563;
    }

    /* 解决 QScrollArea 和内容区域背景问题 */
    QScrollArea, QScrollArea QWidget {
        background-color: #1f2937;
    }
    QScrollArea {
        border: none;
    }

    /* 复选框/单选框样式 */
    QCheckBox, QRadioButton {
        spacing: 8px;
        color: #f3f4f6;
    }
    QCheckBox::indicator, QRadioButton::indicator {
        width: 20px;
        height: 20px;
        border: 2px solid #4b5563;
        background: #111827;
    }
    QCheckBox::indicator { border-radius: 4px; }
    QRadioButton::indicator { border-radius: 10px; }
    QCheckBox::indicator:checked {
        background: #2563eb;
        border-color: #3b82f6;
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IndoaXRlIiBzdHJva2Utd3lkdGg9IjQiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBvbHlsaW5lIHBvaW50cz0iMjAgNiA5IDE3IDQgMTIiPjwvcG9seWxpbmU+PC9zdmc+);
    }
    QRadioButton::indicator:checked {
        background: #10b981;
        border-color: #34d399;
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0Ij48Y2lyY2xlIGN4PSIxMiIgY3k9IjEyIiByPSI1IiBmaWxsPSJ3aGl0ZSIvPjwvc3ZnPg==);
    }

    /* 下拉框进阶样式 */
    QComboBox::drop-down {
        subcontrol-origin: padding;
        subcontrol-position: center right;
        width: 24px;
        border-left: none;
    }
    QComboBox::down-arrow {
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiM5Y2EzYWYiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cG9seWxpbmUgcG9pbnRzPSI2IDkgMTIgMTUgMTggOSI+PC9wb2x5bGluZT48L3N2Zz4=);
        width: 14px;
        height: 14px;
    }
    QComboBox QAbstractItemView {
        background-color: #1f2937;
        color: #f3f4f6;
        border: 1px solid #4b5563;
        selection-background-color: #374151;
        outline: none;
    }
"""
