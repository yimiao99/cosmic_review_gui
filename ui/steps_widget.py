from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QFrame,
    QApplication,
)
from PySide6.QtCore import Signal, Qt, QRectF, QTimer
from PySide6.QtGui import QCursor, QPainter, QPen, QColor, QFont


class StepCircle(QWidget):
    """环形进度图标"""

    circle_clicked = Signal(int)

    def __init__(self, step_num, parent=None):
        super().__init__(parent)
        self.step_num = step_num
        self.status = "pending"
        self.progress = 0
        self.setFixedSize(44, 44)  # 稍微大一点
        self.setCursor(QCursor(Qt.PointingHandCursor))

    def mousePressEvent(self, event):
        self.circle_clicked.emit(self.step_num)
        super().mousePressEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 居中绘制
        rect = QRectF(4, 4, 36, 36)

        # 核心颜色定义 (移除黄色/警告，统一归为异常/红色)
        color_map = {
            "pending": QColor("#94a3b8"),
            "processing": QColor("#3b82f6"),
            "done": QColor("#10b981"),
            "finished": QColor("#10b981"),
            "skipped": QColor("#64748b"),
            "fail": QColor("#ef4444"),
            "error": QColor("#ef4444"),
            "warn": QColor("#ef4444"),  # 警告与错误均显示红色
        }

        main_color = color_map.get(self.status, color_map["pending"])

        # 背景圆环 (玻璃感/空心感)
        painter.setPen(
            QPen(QColor(main_color.red(), main_color.green(), main_color.blue(), 40), 3)
        )
        painter.drawEllipse(rect)

        if self.status in ["done", "finished", "fail", "warn", "skipped"]:
            # 填充圆
            painter.setBrush(main_color)
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(rect)

            # 中心文字/图标
            text = str(self.step_num)
            if self.status in ["done", "finished"]:
                text = "✓"
            elif self.status in ["fail", "warn"]:
                text = "✕"  # 警告也显示 X
            elif self.status == "skipped":
                text = "-"

            painter.setPen(QPen(Qt.white, 2.5))
            font = QFont("Microsoft YaHei UI", 11, QFont.Bold)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignCenter, text)
        else:
            # 等待或处理中
            # 绘制一层极淡的底色增加体积感
            painter.setBrush(
                QColor(main_color.red(), main_color.green(), main_color.blue(), 15)
            )
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(rect)

            if self.progress > 0:
                # 绘制进度弧线 (外圈)
                painter.setPen(QPen(main_color, 4, Qt.SolidLine, Qt.RoundCap))
                span_angle = -int(self.progress * 3.6 * 16)
                painter.drawArc(rect, 90 * 16, span_angle)

            # 中心数字
            painter.setPen(QPen(main_color, 2))
            font = QFont("Segoe UI", 11, QFont.DemiBold)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignCenter, str(self.step_num))


class StepLine(QWidget):
    """动态连接线进度条"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.progress = 0  # 0 to 100
        self.setMinimumWidth(20)
        self.setFixedHeight(44)

        # 动画相关
        self._glow_timer = QTimer(self)
        self._glow_timer.timeout.connect(self.update)
        self._glow_step = 0

    def set_progress(self, val):
        if int(self.progress) != int(val):
            self.progress = max(0, min(100, val))
            # 如果正在进度中且没跑完，启动动画
            if 0 < self.progress < 100:
                if not self._glow_timer.isActive():
                    self._glow_timer.start(50)
            else:
                self._glow_timer.stop()
            self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 预加载主题色
        from extend.matcher_config import MatcherConfig

        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)

        # y 轴对齐：逻辑上圆圈中心在顶部的 10px margin + 22px 半径 = 32px
        # StepLine 作为 44px 高的 widget，如果顶对齐，其中线在 22px。
        # 因此我们需要在绘制时向下偏移 10px。
        y_offset = 10
        y = self.height() / 2 + y_offset

        h = 6
        line_rect = QRectF(0, y - h / 2, self.width(), h)

        # 1. 绘制背景线 (槽)
        bg_color = (
            QColor(255, 255, 255, 30) if is_dark else QColor(0, 0, 0, 60)
        )  # 加深 Light Mode 背景槽，增强对比度
        painter.setBrush(bg_color)
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(line_rect, h / 2, h / 2)

        # 2. 绘制进度填充
        if self.progress > 0:
            from PySide6.QtGui import QLinearGradient

            fill_width = (self.progress / 100.0) * self.width()
            fill_rect = QRectF(0, y - h / 2, fill_width, h)

            if self.progress >= 100:
                # 完成色：翠绿色渐变 (统一使用 #10b981)
                grad = QLinearGradient(fill_rect.left(), 0, fill_rect.right(), 0)
                grad.setColorAt(0, QColor("#10b981"))
                grad.setColorAt(1, QColor("#10b981"))
            else:
                # 进行色：科技蓝渐变，带流动能量感
                self._glow_step = (self._glow_step + 1) % 40
                shift = self._glow_step / 40.0

                grad = QLinearGradient(fill_rect.left(), 0, fill_rect.right(), 0)
                grad.setColorAt(max(0, shift - 0.3), QColor("#3b82f6"))
                grad.setColorAt(shift, QColor("#93c5fd"))
                grad.setColorAt(min(1, shift + 0.3), QColor("#3b82f6"))

            painter.setBrush(grad)
            painter.drawRoundedRect(fill_rect, h / 2, h / 2)

            # 3. 发光效果 (能量核流过感)
            if self.progress < 100:
                glow_color = QColor(147, 197, 253, 60)
                painter.setBrush(glow_color)
                painter.drawRoundedRect(fill_rect.adjusted(-1, -1, 1, 1), h, h)


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
        if status == "done" or status == "finished":
            self.circle.progress = 0  # 完成后清除环形进度

        # 优化文字颜色与对比度 (Light/Dark Mode)
        from extend.matcher_config import MatcherConfig

        is_dark = MatcherConfig.load().get("theme", {}).get("is_dark", False)

        if status in ["done", "finished"]:
            text_color = "#10b981"
        elif status in ["fail", "error", "warn"]:
            text_color = "#ef4444"
        elif status == "processing":
            text_color = "#3b82f6"
        else:
            text_color = "#475569" if not is_dark else "#94a3b8"

        self.label_widget.setStyleSheet(
            f"""
            QLabel {{
                font-family: 'Microsoft YaHei UI';
                font-size: 11px;
                font-weight: 600;
                line-height: 1.2;
                color: {text_color};
            }}
        """
        )

        self.circle.update()
        self.update()

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
        layout.setContentsMargins(0, 0, 0, 0)

        self.container = QWidget()
        self.container.setFixedHeight(110)

        self.update_theme_style()

        container_layout = QHBoxLayout(self.container)
        container_layout.setSpacing(0)
        container_layout.setContentsMargins(20, 0, 20, 0)  # 左右留白增加呼吸感

        self.step_nodes = []
        self.step_lines = []
        for i, (label, status) in enumerate(steps_data, 1):
            node = StepNode(i, label, status)
            node.node_clicked.connect(self.node_clicked.emit)
            node.label_clicked.connect(self.label_clicked.emit)
            self.step_nodes.append(node)
            container_layout.addWidget(node)

            if i < len(steps_data):
                line = StepLine()
                self.step_lines.append(line)
                # 采用顶对齐，具体坐标由 StepLine 内部 y_offset 补偿对齐圆心
                container_layout.addWidget(line, stretch=1, alignment=Qt.AlignTop)
        layout.addWidget(self.container)

    def set_total_progress(self, total_val):
        """基于全局 0-100 的百分比更新所有线段和节点状态"""
        num_lines = len(self.step_lines)
        if num_lines == 0:
            return

        # 增加容错：如果接近 100，直接设为 100
        if total_val >= 99.8:
            total_val = 100.0

        # 采用与 ValidationWorker 任务分配点一致的非线性分段 (优化权重：给耗时长的步骤 [模板/层级/过程] 更多空间)
        # 1->2 (30%), 2->3 (3%), 3->4 (3%), 4->5 (3%), 5->6 (31%), 6->7 (25%)
        breakpoints = [0, 30, 33, 36, 39, 70, 95]

        for i, line in enumerate(self.step_lines):
            if i >= len(breakpoints) - 1:
                # 兜底处理：如果线段多于断点，剩下的平分 96-100
                start_progress = 96
                end_progress = 100
            else:
                start_progress = breakpoints[i]
                end_progress = breakpoints[i + 1]

            if total_val >= end_progress:
                line.set_progress(100)
            elif total_val <= start_progress:
                line.set_progress(0)
            else:
                # 在区间内
                span = end_progress - start_progress
                if span <= 0:
                    span = 1
                local_p = (total_val - start_progress) / span * 100
                line.set_progress(local_p)

    def update_theme_style(self):
        """同步全局主题色"""
        from extend.matcher_config import MatcherConfig

        is_dark = MatcherConfig.load().get("theme", {}).get("is_dark", False)

        # 彻底移除底色，采用全透明设计，消除用户所谓的“顽固底色”
        bg_color = "transparent"
        self.setStyleSheet(f"background: {bg_color}; border: none;")
        if hasattr(self, "container"):
            self.container.setStyleSheet(f"background: {bg_color}; border: none;")

        # 刷新所有圆圈和文字的状态颜色
        if hasattr(self, "step_nodes"):
            for node in self.step_nodes:
                node.update_status(node.status)  # 触发颜色与对比度重算

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
