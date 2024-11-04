# -*- coding: utf-8 -*-
from aqt import mw
from aqt.qt import *
from aqt.utils import showInfo, getFile, chooseList
import aqt
import csv
from aqt.gui_hooks import browser_will_show

def export_selected_notes_as_tsv(browser):
    selected_notes = browser.selectedNotes()

    if not selected_notes:
        showInfo("No notes selected. Please select notes to export.")
        return

    # Load configuration for presets
    config = mw.addonManager.getConfig(__name__)
    presets = config.get("presets", [])

    if not presets:
        showInfo("No presets defined. Please configure presets in addon settings.")
        return

    # Prompt user to select a preset
    preset_names = [preset["name"] for preset in presets]
    selected_preset_idx = chooseList("Select a preset for export:", preset_names)
    if selected_preset_idx == -1:
        return  # User canceled

    # Retrieve selected preset details
    selected_preset = presets[selected_preset_idx]
    export_fields = selected_preset.get("fields", [])

    if not export_fields:
        showInfo("Selected preset has no fields defined for export.")
        return

    # Fetch field names from the first selected note
    col = mw.col
    model = col.getNote(selected_notes[0]).model()
    all_fields = [f['name'] for f in model['flds']]

    # Filter and order fields according to the selected preset
    export_field_indexes = []
    for field in export_fields:
        if field in all_fields:
            export_field_indexes.append(all_fields.index(field))
        else:
            showInfo(f"Field '{field}' not found in note model.")
            return

    # Get TSV save path
    tsv_path, _ = QFileDialog.getSaveFileName(
        None, "Export TSV", "", "TSV Files (*.tsv);;All Files (*)"
    )
    if not tsv_path:
        return

    # Prepare data to write
    rows = []

    # Add header row
    header = [export_fields[i] for i in range(len(export_fields))]
    rows.append(header)

    # Collect selected notes' data
    for note_id in selected_notes:
        note = col.getNote(note_id)
        row = [note.fields[i] for i in export_field_indexes]
        rows.append(row)

    # Sort rows by the first selected field (ascending order)
    rows[1:] = sorted(rows[1:], key=lambda x: x[0])

    # Write to TSV
    with open(tsv_path, mode="w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file, delimiter="\t")
        writer.writerows(rows)

    showInfo(f"Exported {len(selected_notes)} notes to {tsv_path}")

# Hook to add the menu action to the browser window
def init_menu(browser):
    action = QAction("Export Selected Notes as TSV", browser)
    action.triggered.connect(lambda: export_selected_notes_as_tsv(browser))
    browser.form.menuEdit.addSeparator()
    browser.form.menuEdit.addAction(action)

browser_will_show.append(init_menu)
