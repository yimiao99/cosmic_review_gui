from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QFrame,
    QApplication,
)
from PySide6.QtCore import Signal, Qt, QRectF
from PySide6.QtGui import QCursor, QPainter, QPen, QColor, QFont


class StepCircle(QWidget):
    """环形进度图标"""

    circle_clicked = Signal(int)

    def __init__(self, step_num, parent=None):
        super().__init__(parent)
        self.step_num = step_num
        self.status = "pending"
        self.progress = 0
        self.setFixedSize(40, 40)
        self.setCursor(QCursor(Qt.PointingHandCursor))

    def mousePressEvent(self, event):
        self.circle_clicked.emit(self.step_num)
        super().mousePressEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(4, 4, 32, 32)

        palette = self.palette()
        border_color = palette.color(self.foregroundRole())
        border_color.setAlpha(50)  # 设置较低透明度作为背景环

        painter.setPen(QPen(border_color, 2))
        painter.drawEllipse(rect)

        # 待处理颜色
        color = palette.color(self.foregroundRole())
        color.setAlpha(120)  # 稍微提高透明度更清晰
        text = str(self.step_num)

        if self.status == "done":
            color = QColor("#10b981")
            text = "✓"
        elif self.status == "finished":
            color = QColor("#3b82f6")  # 蓝色，表示该步骤技术上已跑完，但结果待最终确认
            text = "✓"
        elif self.status == "skipped":
            color = QColor("#94a3b8")  # 灰色，表示跳过
            text = "○"
        elif self.status == "fail":
            color = QColor("#ef4444")
        elif self.status == "warn":
            color = QColor("#ef4444")
            text = "!"
        elif self.status == "processing":
            color = QColor("#3b82f6")

        if self.status != "pending" or self.progress > 0:
            if self.status in ["done", "fail", "warn", "finished", "skipped"]:
                # 已结束的状态：不画进度弧线，只画实心圆
                painter.setBrush(color)
                painter.setPen(Qt.NoPen)
                painter.drawEllipse(rect)
                painter.setPen(QPen(Qt.white, 2))
            else:
                # 进行中的状态：画进度弧线
                painter.setPen(QPen(color, 3))
                span_angle = -int(self.progress * 3.6 * 16)
                painter.drawArc(rect, 90 * 16, span_angle)
                painter.setPen(QPen(color, 2))
        else:
            painter.setPen(QPen(QColor("#cbd5e1"), 2))

        font = QFont("Microsoft YaHei UI")
        font.setBold(True)
        font.setPointSize(10)
        painter.setFont(font)
        painter.setPen(
            Qt.white
            if self.status in ["done", "fail", "warn", "skipped", "finished"]
            else palette.color(self.foregroundRole())
        )
        painter.drawText(rect, Qt.AlignCenter, text)


class StepNode(QWidget):
    """单个步骤节点 (支持分离点击)"""

    node_clicked = Signal(int)
    label_clicked = Signal(int)

    def __init__(self, step_num, label, status="pending", parent=None):
        super().__init__(parent)
        self.step_num = step_num
        self.status = status
        # 弹性宽度，降低高度限制
        self.setMinimumWidth(90)
        self.setMaximumWidth(140)
        self.setMinimumHeight(100)

        layout = QVBoxLayout(self)
        # 使用顶部居中对齐
        layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        layout.setSpacing(6)  # 减少圆圈和文字的间距
        layout.setContentsMargins(0, 10, 0, 0)

        self.circle = StepCircle(step_num)
        self.circle.circle_clicked.connect(self.node_clicked.emit)
        layout.addWidget(self.circle, alignment=Qt.AlignCenter)

        self.label_widget = QLabel(label)
        self.label_widget.setAlignment(Qt.AlignCenter)
        self.label_widget.setProperty("class", "step-label")
        # 保持字体大小，优化行高
        self.label_widget.setStyleSheet(
            "font-family: 'Microsoft YaHei UI'; font-size: 11px; font-weight: 500; line-height: 1.2;"
        )
        self.label_widget.setWordWrap(True)
        self.label_widget.setCursor(QCursor(Qt.PointingHandCursor))
        # 允许标签根据需要扩展高度
        layout.addWidget(self.label_widget)

        self.update_status(status)

    def update_status(self, status):
        self.status = status
        self.circle.status = status
        if status == "done":
            self.circle.progress = 0  # 完成后清除环形进度

        # 强制触发一次文字颜色更新
        self.set_progress(self.circle.progress)
        self.circle.update()

    def set_progress(self, value):
        self.circle.progress = value
        if value > 0:
            if self.status in ["pending", "processing"]:
                self.status = "processing"
                self.circle.status = "processing"
        elif value == 0:
            if self.status == "processing":
                self.status = "pending"
                self.circle.status = "pending"

        # 处理状态显示
        if self.status == "processing":
            self.label_widget.setProperty("status", "processing")
        elif self.status == "done":
            self.label_widget.setProperty("status", "done")
        elif self.status == "finished":
            self.label_widget.setProperty("status", "finished")
        elif self.status == "skipped":
            self.label_widget.setProperty("status", "skipped")
        elif self.status in ["fail", "warn"]:
            self.label_widget.setProperty("status", "error")
        else:
            self.label_widget.setProperty("status", "pending")

        # 强制触发样式更新
        self.label_widget.style().unpolish(self.label_widget)
        self.label_widget.style().polish(self.label_widget)

        self.circle.update()
        self.update()

    def mousePressEvent(self, event):
        child = self.childAt(event.pos())
        if child == self.label_widget:
            self.label_clicked.emit(self.step_num)
        super().mousePressEvent(event)


class StepsWidget(QWidget):
    """步骤进度条组件 (信号源)"""

    node_clicked = Signal(int)
    label_clicked = Signal(int)

    def __init__(self, steps_data, parent=None):
        super().__init__(parent)
        self.steps_data = steps_data

        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)  # 进一步压缩外边距

        self.container = QWidget()
        # 调整容器高度，确保紧致
        self.container.setMinimumHeight(120)
        self.container.setFixedHeight(120)

        self.update_theme_style()

        container_layout = QHBoxLayout(self.container)
        container_layout.setSpacing(0)
        container_layout.setContentsMargins(5, 0, 5, 0)

        self.step_nodes = []
        for i, (label, status) in enumerate(steps_data, 1):
            node = StepNode(i, label, status)
            node.node_clicked.connect(self.node_clicked.emit)
            node.label_clicked.connect(self.label_clicked.emit)
            self.step_nodes.append(node)
            container_layout.addWidget(node)

            if i < len(steps_data):
                line = QFrame()
                line.setFrameShape(QFrame.HLine)
                # 关键：通过 margin-top 将连接线对齐到圆圈中心 (10px margin + 20px radius - 1px half-height = 29px)
                line.setStyleSheet(
                    "background: palette(mid); margin-top: 30px; border: none; min-height: 2px; max-height: 2px;"
                )
                line.setFixedHeight(32)  # 30px top + 2px line
                container_layout.addWidget(line, stretch=1, alignment=Qt.AlignTop)
        layout.addWidget(self.container)

    def update_theme_style(self):
        """刷新主题样式"""
        from extend.matcher_config import MatcherConfig

        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)

        # 备用检测
        if not is_dark:
            qss = QApplication.instance().styleSheet() or ""
            is_dark = "background-color: #1f2937" in qss

        # 深色模式去掉底色（透明），浅色模式保持中灰色
        bg_color = "transparent" if is_dark else "#e5e7eb"
        self.setStyleSheet(f"background: {bg_color};")
        if hasattr(self, "container"):
            self.container.setStyleSheet(f"background: {bg_color};")

        # 刷新所有圆圈和文字的状态颜色
        if hasattr(self, "step_nodes"):
            for node in self.step_nodes:
                node.update()
                node.label_widget.style().unpolish(node.label_widget)
                node.label_widget.style().polish(node.label_widget)

    def set_step_status(self, step_num, status):
        """更新指定步骤的状态 (1-indexed)"""
        if 1 <= step_num <= len(self.step_nodes):
            node = self.step_nodes[step_num - 1]
            node.update_status(status)
            # 如果是完成状态，清除进度显示
            if status in ["done", "fail", "warn", "finished"]:
                node.circle.progress = 0  # 显式清零
                node.circle.update()

    def set_step_progress(self, step_num, progress):
        """更新指定步骤的进度 (1-indexed)"""
        if 1 <= step_num <= len(self.step_nodes):
            # 遍历所有节点，确保只有当前节点在转
            for i, node in enumerate(self.step_nodes):
                curr_idx = i + 1
                if curr_idx == step_num:
                    node.set_progress(progress)
                else:
                    # 非当前节点：如果是在进行中，强制回归 pending 或保持之前的结果状态
                    # 关键是清除它们的进度数值，防止画出环
                    node.circle.progress = 0
                    if node.status == "processing":
                        # 只有在还没有进入最终状态(done/fail/warn/finished)时才重置
                        node.status = "pending"
                        node.circle.status = "pending"
                    node.update()

    def clear_all_progress(self):
        """清除所有步骤的进度环"""
        for node in self.step_nodes:
            node.circle.progress = 0
            if node.status == "processing":
                node.status = "pending"
                node.circle.status = "pending"
            node.update_status(node.status)  # 刷新一次样式
