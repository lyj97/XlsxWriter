"""Local Excel structure alignment and viewing application."""

import json
import os
import tempfile
from pathlib import Path
import sys

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QProcess, Qt, QTimer
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QFileDialog, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton, QTableView,
    QVBoxLayout, QWidget, QSplitter, QTableWidget, QTableWidgetItem, QTextEdit,
)


LOAD_TIMEOUT_MS = 30000


def column_title(index):
    title = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        title = chr(65 + remainder) + title
    return title


class SheetModel(QAbstractTableModel):
    def __init__(self, sheet, parent=None):
        super().__init__(parent)
        self.sheet = sheet

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.sheet["rows"])

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.sheet["columns"]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            row = self.sheet["rows"][index.row()]
            return row[index.column()] if index.column() < len(row) else ""
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return (column_title(section) if orientation == Qt.Orientation.Horizontal
                    else str(section + 1))
        return None


class Viewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel 结构对齐工具")
        self.resize(1000, 700)
        self.sheets = []
        self.model = None
        self.failure = None
        self.filename = ""
        self.open_button = QPushButton("打开工作簿…")
        self.cancel_button = QPushButton("取消加载")
        self.cancel_button.setEnabled(False)
        self.selector = QComboBox()
        self.selector.setMinimumWidth(200)
        self.selector.setEnabled(False)
        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索当前工作表…")
        self.search.setClearButtonEnabled(True)
        self.search.setEnabled(False)
        self.search.setMinimumWidth(180)
        self.search_previous = QPushButton("上一个")
        self.search_next = QPushButton("下一个")
        self.search_previous.setEnabled(False)
        self.search_next.setEnabled(False)
        self.search_result = QLabel("输入关键词")
        self.search_result.setMinimumWidth(70)
        self.search_matches = []
        self.search_position = -1
        self.table = QTableView()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setWordWrap(False)
        self.table.horizontalHeader().setDefaultSectionSize(120)
        self.status = QLabel("请选择 .xlsx 或 .xlsm 文件。公式显示为文本。")
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        self.status.setWordWrap(True)
        controls = QHBoxLayout()
        controls.addWidget(self.open_button)
        controls.addWidget(self.cancel_button)
        controls.addWidget(QLabel("工作表："))
        controls.addWidget(self.selector)
        controls.addWidget(QLabel("搜索："))
        controls.addWidget(self.search)
        controls.addWidget(self.search_previous)
        controls.addWidget(self.search_next)
        controls.addWidget(self.search_result)
        controls.addStretch()
        layout = QVBoxLayout()
        layout.addLayout(controls)
        self.analysis = None
        self.staging = None
        self.job = None
        self.analyze_button = QPushButton("分析结构")
        self.analyze_button.setEnabled(False)
        self.groups = QComboBox()
        self.scope_dirty = False
        self.regenerate_button = QPushButton("重新生成计划")
        self.regenerate_button.setEnabled(False)
        self.members = QTableWidget(0, 1)
        self.members.setHorizontalHeaderLabels(["参与对齐的工作表（至少两张）"])
        self.members.setMaximumHeight(120)
        self.equivalences = QTableWidget(0, 2)
        self.equivalences.setHorizontalHeaderLabels(["视为同一项目", "候选原因（不修改原标签）"])
        self.equivalences.setMaximumHeight(100)
        self.equivalences.setColumnWidth(0, 380)
        for table in (self.members, self.equivalences):
            table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            table.horizontalHeader().setStretchLastSection(True)
        self.save_button = QPushButton("另存对齐副本…")
        self.save_button.setEnabled(False)
        self.cancel_job_button = QPushButton("取消任务")
        self.cancel_job_button.setEnabled(False)
        review_controls = QHBoxLayout()
        for widget in (self.analyze_button, self.groups, self.regenerate_button, self.save_button, self.cancel_job_button):
            review_controls.addWidget(widget)
        self.actions = QTableWidget(0, 8)
        self.actions.setHorizontalHeaderLabels(["应用", "工作表", "操作", "源行", "预计新行", "项目 / 上下文", "置信度", "原因"])
        self.actions.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.actions.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.actions.horizontalHeader().setStretchLastSection(True)
        self.summary = QLabel("未分析")
        self.summary.setWordWrap(True)
        self.summary.setTextFormat(Qt.TextFormat.PlainText)
        self.report = QTextEdit()
        self.report.setReadOnly(True)
        self.report.setMaximumHeight(110)
        self.report.setPlainText("仅支持保守的精确顺序并集；勾选部分操作时不保证完整对齐。\n"
                                 "openpyxl 无法保证绘图、外部链接等 Excel 对象完整往返。请保留原件并在 Excel 中复核。")
        review = QWidget()
        review_layout = QVBoxLayout(review)
        review_layout.addLayout(review_controls)
        review_layout.addWidget(QLabel("一次另存仅应用当前组的所选工作表；其他组需打开副本后继续。"))
        review_layout.addWidget(self.members)
        review_layout.addWidget(self.equivalences)
        review_layout.addWidget(self.summary)
        review_layout.addWidget(self.actions)
        review_layout.addWidget(self.report)
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.addWidget(self.table)
        splitter.addWidget(review)
        splitter.setSizes([400, 300])
        layout.addWidget(splitter)
        layout.addWidget(self.status)
        central = QWidget()
        central.setLayout(layout)
        self.setCentralWidget(central)

        self.worker = QProcess(self)
        self.worker.finished.connect(self._job_finished)
        self.worker.errorOccurred.connect(self._job_error)
        self.job_timer = QTimer(self)
        self.job_timer.setSingleShot(True)
        self.job_timer.timeout.connect(lambda: self.cancel_job("失败：任务超时，已停止（分析 120 秒，保存 10 分钟）。"))
        self.analyze_button.clicked.connect(self.start_analysis)
        self.regenerate_button.clicked.connect(self.regenerate_plan)
        self.members.itemChanged.connect(self.scope_changed)
        self.equivalences.itemChanged.connect(self.scope_changed)
        self.groups.currentIndexChanged.connect(self.show_plan)
        self.actions.itemChanged.connect(self.selection_changed)
        self.actions.cellClicked.connect(self.navigate_action)
        self.save_button.clicked.connect(self.save_aligned)
        self.cancel_job_button.clicked.connect(lambda: self.cancel_job("未完成：任务已取消。"))
        self.process = QProcess(self)
        self.process.finished.connect(self._finished)
        self.process.errorOccurred.connect(self._process_error)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self._timeout)
        self.open_button.clicked.connect(self.choose_file)
        self.cancel_button.clicked.connect(self.cancel_load)
        self.selector.currentIndexChanged.connect(self.show_sheet)
        self.search.textChanged.connect(self.update_search)
        self.search.returnPressed.connect(self.next_search_match)
        self.search_previous.clicked.connect(self.previous_search_match)
        self.search_next.clicked.connect(self.next_search_match)
        QShortcut(QKeySequence.StandardKey.Find, self, activated=self.search.setFocus)

    def choose_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "打开 Excel 工作簿", "", "Excel 工作簿 (*.xlsx *.xlsm)"
        )
        if filename:
            self.load_file(filename)

    def load_file(self, filename):
        if self.process.state() != QProcess.ProcessState.NotRunning:
            return
        if self.worker.state() != QProcess.ProcessState.NotRunning:
            return
        self.analysis = None
        self.groups.clear()
        self.actions.setRowCount(0)
        self.summary.setText("未分析")
        self.analyze_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.filename = str(filename)
        self.failure = None
        self.open_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.selector.clear()
        self.selector.setEnabled(False)
        self.search.clear()
        self.search.setEnabled(False)
        self.sheets = []
        self.table.setModel(None)
        self.model = None
        self.setWindowTitle("Excel 结构对齐工具")
        self.status.setText("正在加载…最多等待 30 秒，可随时取消。")
        worker = str(Path(__file__).with_name("loader.py"))
        self.timer.start(LOAD_TIMEOUT_MS)
        self.process.start(sys.executable, [worker, self.filename])

    def cancel_load(self):
        self.failure = "已取消加载。"
        self.process.kill()

    def _timeout(self):
        self.failure = "加载超过 30 秒，已停止。请缩小工作簿后重试。"
        self.process.kill()

    def _idle(self):
        self.timer.stop()
        self.open_button.setEnabled(True)
        self.cancel_button.setEnabled(False)

    def _process_error(self, error):
        if error == QProcess.ProcessError.FailedToStart:
            self._idle()
            self._report_error("无法启动加载进程，请检查 Python 环境。")

    def _report_error(self, message):
        self.status.setText(message)
        QMessageBox.warning(self, "无法打开工作簿", message)

    def _finished(self, exit_code, exit_status):
        self._idle()
        output = bytes(self.process.readAllStandardOutput())
        self.process.readAllStandardError()
        if self.failure:
            self.status.setText(self.failure)
            return
        if exit_code or exit_status != QProcess.ExitStatus.NormalExit:
            self._report_error("加载进程异常退出。请检查依赖是否安装，或尝试较小的有效工作簿。")
            return
        try:
            result = json.loads(output)
            if "error" in result:
                self._report_error(result["error"])
                return
            self.sheets = result["sheets"]
        except (ValueError, KeyError):
            self._report_error("加载进程返回了无效数据。")
            return
        self.selector.addItems([sheet["name"] for sheet in self.sheets])
        self.selector.setEnabled(True)
        self.search.setEnabled(True)
        self.analyze_button.setEnabled(True)
        self.setWindowTitle(f"{Path(self.filename).name} — Excel 结构对齐工具")

    def show_sheet(self, index):
        if not 0 <= index < len(self.sheets):
            return
        sheet = self.sheets[index]
        self.model = SheetModel(sheet)
        self.table.setModel(self.model)
        rows, columns = len(sheet["rows"]), sheet["columns"]
        detail = f"{rows} 行 × {columns} 列" if rows else "空工作表"
        self.status.setText(f"{sheet['name']}：{detail}。只读；公式显示为文本。")
        self.update_search()

    def update_search(self):
        query = self.search.text().casefold()
        self.search_matches = []
        self.search_position = -1
        if query and self.model:
            for row_number, row in enumerate(self.model.sheet["rows"]):
                for column_number, value in enumerate(row):
                    if query in value.casefold():
                        self.search_matches.append((row_number, column_number))
        enabled = bool(self.search_matches)
        self.search_previous.setEnabled(enabled)
        self.search_next.setEnabled(enabled)
        if enabled:
            self.search_position = 0
            self._show_search_match()
        else:
            self.search_result.setText("无匹配" if query else "输入关键词")

    def previous_search_match(self):
        if self.search_matches:
            self.search_position = (self.search_position - 1) % len(self.search_matches)
            self._show_search_match()

    def next_search_match(self):
        if self.search_matches:
            self.search_position = (self.search_position + 1) % len(self.search_matches)
            self._show_search_match()

    def _show_search_match(self):
        row, column = self.search_matches[self.search_position]
        index = self.model.index(row, column)
        self.table.setCurrentIndex(index)
        self.table.scrollTo(index)
        self.search_result.setText(f"{self.search_position + 1} / {len(self.search_matches)}")

    def start_analysis(self):
        self.analysis = None
        self.groups.clear()
        self.actions.setRowCount(0)
        self._start_job({"mode": "analyze", "source": self.filename})

    def scope_changed(self, *args):
        self.scope_dirty = True
        self.save_button.setEnabled(False)
        count = sum(self.members.item(i, 0).checkState() == Qt.CheckState.Checked
                    for i in range(self.members.rowCount()))
        self.regenerate_button.setEnabled(count >= 2 and self.job is None)
        self.summary.setText("选择已更改，请重新生成计划。" if count >= 2 else "请至少选择两张候选工作表。")

    def regenerate_plan(self):
        if not self.regenerate_button.isEnabled():
            return
        index = self.groups.currentIndex()
        scopes = dict(self.analysis.get("scopes", {}))
        scopes[str(index)] = {
            "sheets": [self.members.item(i, 0).text() for i in range(self.members.rowCount())
                       if self.members.item(i, 0).checkState() == Qt.CheckState.Checked],
            "equivalences": [self.equivalences.item(i, 0).data(Qt.ItemDataRole.UserRole)
                             for i in range(self.equivalences.rowCount())
                             if self.equivalences.item(i, 0).checkState() == Qt.CheckState.Checked]}
        # Changing sheets invalidates old terminology candidates; regenerate them first.
        if scopes[str(index)]["sheets"] != self.analysis["groups"][index]["sheets"]:
            scopes[str(index)]["equivalences"] = []
        self._start_job({"mode": "analyze", "source": self.filename, "scopes": scopes})

    def _start_job(self, request):
        self.review_index = max(0, self.groups.currentIndex())
        self.job = request["mode"]
        self.regenerate_button.setEnabled(False)
        self.members.setEnabled(False)
        self.equivalences.setEnabled(False)
        self.job_failure = None
        self.open_button.setEnabled(False)
        self.analyze_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.groups.setEnabled(False)
        self.actions.setEnabled(False)
        self.cancel_job_button.setEnabled(True)
        self.summary.setText("正在分析…" if self.job == "analyze" else "正在保存并验证临时副本…")
        self.worker.start(sys.executable, [str(Path(__file__).with_name("alignment.py"))])
        self.worker.write(json.dumps(request, ensure_ascii=False).encode("utf-8"))
        self.worker.closeWriteChannel()
        self.job_timer.start(120000 if self.job == "analyze" else 600000)

    def selected_actions(self):
        return [i for i in range(self.actions.rowCount())
                if self.actions.item(i, 0).checkState() == Qt.CheckState.Checked]

    def show_plan(self, index):
        self.scope_dirty = False
        for table in (self.members, self.equivalences):
            table.blockSignals(True)
            table.setRowCount(0)
        self.actions.blockSignals(True)
        self.actions.setRowCount(0)
        if self.analysis and 0 <= index < len(self.analysis["groups"]):
            group = self.analysis["groups"][index]
            for name in group["members"]:
                i = self.members.rowCount()
                self.members.insertRow(i)
                item = QTableWidgetItem(name)
                item.setCheckState(Qt.CheckState.Checked if name in group["sheets"] else Qt.CheckState.Unchecked)
                self.members.setItem(i, 0, item)
            for candidate in group["candidates"]:
                i = self.equivalences.rowCount()
                self.equivalences.insertRow(i)
                item = QTableWidgetItem(" = ".join(candidate["labels"]))
                item.setData(Qt.ItemDataRole.UserRole, candidate["labels"])
                item.setCheckState(Qt.CheckState.Checked if candidate["labels"] in group["equivalences"] else Qt.CheckState.Unchecked)
                self.equivalences.setItem(i, 0, item)
                self.equivalences.setItem(i, 1, QTableWidgetItem(candidate["reason"]))
            self.regenerate_button.setEnabled(len(group["sheets"]) >= 2 and self.job is None)
            self.actions.setRowCount(len(group["actions"]))
            for i, action in enumerate(group["actions"]):
                check = QTableWidgetItem()
                check.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable |
                               (Qt.ItemFlag.ItemIsUserCheckable if not action["blocking"] else Qt.ItemFlag.NoItemFlags))
                check.setCheckState(Qt.CheckState.Checked if action["selected"] else Qt.CheckState.Unchecked)
                self.actions.setItem(i, 0, check)
                values = [action["sheet"], "插入模板行", action["old_row"], action["new_row"],
                          action["label"] + " / " + action["context"], action["confidence"], action["reason"]]
                for c, value in enumerate(values, 1):
                    self.actions.setItem(i, c, QTableWidgetItem(str(value)))
            self.report.setPlainText("\n".join(group["blockers"]) or self.analysis["warning"])
        for table in (self.members, self.equivalences):
            table.blockSignals(False)
        self.actions.blockSignals(False)
        self.selection_changed()

    def selection_changed(self, *args):
        index = self.groups.currentIndex()
        valid = self.analysis and 0 <= index < len(self.analysis["groups"])
        if not valid:
            self.save_button.setEnabled(False)
            return
        group = self.analysis["groups"][index]
        count = len(self.selected_actions())
        blocked = bool(group["blockers"])
        self.summary.setText(("已阻断" if blocked else "发现差异" if group["actions"] else "结构一致") +
                             f"；操作 {len(group['actions'])}，已选 {count}，阻断 {len(group['blockers'])}。新行号按完整选择预览。")
        self.save_button.setEnabled(bool(count and not blocked and not self.scope_dirty and self.job is None))

    def navigate_action(self, row, column):
        group = self.analysis["groups"][self.groups.currentIndex()]
        action = group["actions"][row]
        self.selector.setCurrentText(action["sheet"])
        if self.model and self.model.rowCount():
            index = self.model.index(min(action["old_row"] - 1, self.model.rowCount() - 1), action["column"] - 1)
            self.table.setCurrentIndex(index)
            self.table.selectRow(index.row())
            self.table.scrollTo(index)

    def save_aligned(self):
        if not self.save_button.isEnabled():
            return
        from alignment import validate_destination
        filename, _ = QFileDialog.getSaveFileName(self, "另存对齐副本（选择新文件名）", "",
                                                 f"Excel (*{Path(self.filename).suffix})")
        if not filename:
            return
        try:
            self.destination = validate_destination(self.filename, filename)
            fd, staging = tempfile.mkstemp(prefix=".alignment-", suffix=self.destination.suffix,
                                           dir=self.destination.parent)
            os.close(fd)
            self.staging = Path(staging)
            self._start_job({"mode": "apply", "source": self.filename, "destination": str(self.destination),
                             "staging": staging, "analysis": self.analysis,
                             "group": self.groups.currentIndex(), "selected": self.selected_actions()})
        except Exception as exc:
            self.summary.setText("失败：" + str(exc))
            self._cleanup_staging()

    def cancel_job(self, message):
        self.job_failure = message
        self.worker.kill()

    def _cleanup_staging(self):
        if self.staging:
            self.staging.unlink(missing_ok=True)
            self.staging = None

    def _job_error(self, error):
        if error == QProcess.ProcessError.FailedToStart:
            self.job_failure = "失败：无法启动工作进程。"
            self._job_finished(1, QProcess.ExitStatus.CrashExit)

    def _job_finished(self, code, status):
        from alignment import validate_destination, fingerprint
        self.job_timer.stop()
        mode, self.job = self.job, None
        self.open_button.setEnabled(True)
        self.analyze_button.setEnabled(bool(self.sheets))
        self.groups.setEnabled(True)
        self.actions.setEnabled(True)
        self.members.setEnabled(True)
        self.equivalences.setEnabled(True)
        self.cancel_job_button.setEnabled(False)
        output = bytes(self.worker.readAllStandardOutput())
        self.worker.readAllStandardError()
        try:
            if self.job_failure:
                raise ValueError(self.job_failure)
            if code or status != QProcess.ExitStatus.NormalExit:
                raise ValueError("工作进程异常退出。")
            response = json.loads(output)
            if "error" in response:
                raise ValueError(response["error"])
            result = response["result"]
            if mode == "analyze":
                self.analysis = result
                self.groups.blockSignals(True)
                self.groups.clear()
                self.groups.addItems([" / ".join(g["members"]) for g in result["groups"]])
                self.groups.setCurrentIndex(min(self.review_index, len(result["groups"]) - 1))
                self.groups.blockSignals(False)
                self.show_plan(self.groups.currentIndex())
                if not result["groups"]:
                    self.summary.setText("分析完成：未发现可对齐表头。")
            else:
                validate_destination(self.filename, self.destination)
                if fingerprint(self.filename) != self.analysis["sha256"]:
                    raise ValueError("源文件已变化；未发布副本。")
                os.replace(self.staging, self.destination)
                self.staging = None
                self.summary.setText("已保存并验证：" + str(self.destination))
                self.report.setPlainText(json.dumps(result, ensure_ascii=False, indent=2))
        except Exception as exc:
            self.summary.setText("失败：" + str(exc))
            self.report.setPlainText(str(exc))
        finally:
            self._cleanup_staging()

    def closeEvent(self, event):
        self.job_timer.stop()
        self.job_failure = "已关闭。"
        self.worker.kill()
        self.worker.waitForFinished(1000)
        self._cleanup_staging()
        self.timer.stop()
        self.failure = "已停止加载。"
        self.process.kill()
        self.process.waitForFinished(1000)
        event.accept()


def main():
    app = QApplication(sys.argv)
    window = Viewer()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
