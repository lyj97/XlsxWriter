"""Local, read-only Excel viewer. Run with: python viewer.py"""

import json
from pathlib import Path
import sys

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QProcess, Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QFileDialog, QHBoxLayout,
    QLabel, QMainWindow, QMessageBox, QPushButton, QTableView, QVBoxLayout,
    QWidget,
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
        self.setWindowTitle("Excel 查看器（只读）")
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
        controls.addStretch()
        layout = QVBoxLayout()
        layout.addLayout(controls)
        layout.addWidget(self.table)
        layout.addWidget(self.status)
        central = QWidget()
        central.setLayout(layout)
        self.setCentralWidget(central)

        self.process = QProcess(self)
        self.process.finished.connect(self._finished)
        self.process.errorOccurred.connect(self._process_error)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self._timeout)
        self.open_button.clicked.connect(self.choose_file)
        self.cancel_button.clicked.connect(self.cancel_load)
        self.selector.currentIndexChanged.connect(self.show_sheet)

    def choose_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "打开 Excel 工作簿", "", "Excel 工作簿 (*.xlsx *.xlsm)"
        )
        if filename:
            self.load_file(filename)

    def load_file(self, filename):
        if self.process.state() != QProcess.ProcessState.NotRunning:
            return
        self.filename = str(filename)
        self.failure = None
        self.open_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.selector.clear()
        self.selector.setEnabled(False)
        self.sheets = []
        self.table.setModel(None)
        self.model = None
        self.setWindowTitle("Excel 查看器（只读）")
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
        self.setWindowTitle(f"{Path(self.filename).name} — Excel 查看器（只读）")

    def show_sheet(self, index):
        if not 0 <= index < len(self.sheets):
            return
        sheet = self.sheets[index]
        self.model = SheetModel(sheet)
        self.table.setModel(self.model)
        rows, columns = len(sheet["rows"]), sheet["columns"]
        detail = f"{rows} 行 × {columns} 列" if rows else "空工作表"
        self.status.setText(f"{sheet['name']}：{detail}。只读；公式显示为文本。")

    def closeEvent(self, event):
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
