from PyQt6.QtWidgets import QInputDialog, QMessageBox

def choose_column_dialog(parent, title, prompt, default=None):
    """Общий диалог для выбора колонки."""
    text, ok = QInputDialog.getText(
        parent,
        title,
        prompt,
        text=default
    )
    return text if ok and text else None

def choose_sheet_dialog(parent, items, is_multi=False):
    """Диалог для выбора листа."""
    if not items:
        QMessageBox.warning(parent, "Ошибка", "В файле нет листов!")
        return None

    if is_multi:
        item, ok = QInputDialog.getText(
            parent,
            "Выбор листов",
            f"Введите названия листов через запятую (доступные: {', '.join(items)}):"
        )
        if ok and item:
            return [i.strip() for i in item.split(",")]
    else:
        item, ok = QInputDialog.getItem(
            parent,
            "Выбор листа",
            "Выберите лист:",
            items,
            0,
            False
        )
        if ok and item:
            return item

    return None