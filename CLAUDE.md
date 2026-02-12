# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NoHands is a Chrome Extension (Manifest V3) for French property management. It parses tab-separated Excel data from three sheets (PROP, LOTS, BAIL), lets users map columns to HTML input names, and auto-fills web forms via content script injection. All documentation and UI text is in French.

**No build system, no dependencies, no tests, no linter.** Pure vanilla JS/HTML/CSS loaded directly as an unpacked Chrome extension.

## Loading the Extension

1. Navigate to `chrome://extensions/`
2. Enable Developer Mode
3. Click "Load unpacked extension" and select this directory

Generate placeholder icons if needed:
```powershell
powershell -ExecutionPolicy Bypass -File create-icons.ps1
```

## Architecture

Three-part Chrome MV3 extension:

- **popup.html / popup.js** — Main UI. Handles Excel data parsing for 3 sheet types (`parseExcelData(rawData, sheetType)`), per-sheet field mapping configuration, row selection for multi-row sheets (LOTS/BAIL), and triggering form fill via `chrome.tabs.sendMessage`.
- **content.js** — Injected into web pages. Receives merged data + mapping via message, fills form fields by input name. Handles text, select, checkbox (O/N), radio, date (DD/MM/YYYY → YYYY-MM-DD), and hidden inputs. Dispatches `input`, `change`, `blur` events for framework compatibility.
- **background.js** — Service worker. Adds "Copier le nom de l'input" context menu for discovering input names.

**Data flow:** User pastes TSV per sheet tab → `popup.js` parses into keyed objects → stored in `chrome.storage.local` → user selects LOT/BAIL row → user configures column-to-input mapping per sheet → popup merges PROP + selected LOT + selected BAIL data/mappings → sends to `content.js` → content script fills DOM inputs.

## Key Data Structures

### Sheet Type Definitions (`SHEET_TYPES` in popup.js)
- **PROP**: 26 columns, single row per proprietor (`multiRow: false`)
- **LOTS**: 13 columns, multiple rows per proprietor (`multiRow: true`)
- **BAIL**: 48 columns, multiple rows per proprietor (`multiRow: true`)

Each sheet type defines: `id`, `label`, `columnCount`, `multiRow`, `columns`, `summaryColumns`, `storageKey`, `mappingKey`, `pasteHint`.

### State
- **sheetState**: `{ PROP: { data, selectedIndex }, LOTS: { data, selectedIndex }, BAIL: { data, selectedIndex } }` — PROP.data is an object, LOTS/BAIL.data are arrays of objects
- **sheetMappings**: `{ PROP: {}, LOTS: {}, BAIL: {} }` — per-sheet column-to-input mappings
- **customFieldsArray**: `Array<{ name, value }>` — static values not from Excel

### Storage Keys (chrome.storage.local)
- `propData` / `lotsData` / `bailData` — parsed data per sheet
- `fieldMapping_PROP` / `fieldMapping_LOTS` / `fieldMapping_BAIL` — mapping configs
- `selectedLOTSIndex` / `selectedBAILIndex` — selected row index for multi-row sheets
- `customFields` — custom static fields

Legacy keys (`parsedData`, `fieldMapping`) are auto-migrated on first load.

## Notable Behaviors

- IBAN auto-formatted with spaces every 4 characters (any field with "IBAN" in label)
- Select matching uses 6-level cascade: exact value → exact text → case-insensitive value → case-insensitive text → partial text → partial value
- Checkboxes recognize O/Oui/Yes/True/1/On/Checked as truthy
- Content script guards against injection on chrome://, about://, and extension pages
- Fill button requires PROP data + PROP mapping (LOTS/BAIL are optional additions)
- Data fusion order: PROP → LOTS → BAIL (later values override for shared column names)
