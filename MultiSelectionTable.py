from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QTableWidget, QApplication

class MultiSelectionTable(QTableWidget):
    """QTableWidget subclass that supports multi-cell Ctrl+C copying like Excel."""
    def keyPressEvent(self, event):
        super().keyPressEvent(event)
        if event.key() == Qt.Key.Key_C and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            copied_cells = sorted(self.selectedIndexes())
            if not copied_cells:
                return

            max_col = copied_cells[-1].column()
            copy_text = ""
            for idx in copied_cells:
                item = self.item(idx.row(), idx.column())
                copy_text += item.text() if item else ""
                copy_text += "\n" if idx.column() == max_col else "\t"

            QApplication.clipboard().setText(copy_text)
