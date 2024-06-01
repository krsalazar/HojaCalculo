document.addEventListener('DOMContentLoaded', () => {
    loadSheet('Sheet1');
    loadSheets();
});

let hot;  // Handsontable instance
let activeCell = null;

async function loadSheet(sheetName) {
    const response = await fetch(`/get_sheet?sheet_name=${sheetName}`);
    const data = await response.json();
    renderSheet(data, sheetName);
    // Restablecer la celda activa
    if (activeCell) {
        const hot = Handsontable.getInstance(document.getElementById('spreadsheet'));
        hot.selectCell(activeCell.row, activeCell.col);
    }
}

function renderSheet(data, sheetName) {
    const container = document.getElementById('spreadsheet');
    if (hot) {
        hot.destroy();
    }
    hot = new Handsontable(container, {
        data: data,
        rowHeaders: true,
        colHeaders: true,
        contextMenu: true,
        formulas: true,
        licenseKey: 'non-commercial-and-evaluation',
        afterChange: async (changes, source) => {
            if (source === 'edit') {
                for (let [row, col, oldValue, newValue] of changes) {
                    await setValue(sheetName, row, col, newValue);
                }
            }
        },
        afterSelection: (row, col) => {
            activeCell = { row, col };
        }
    });
}

async function setValue(sheetName, row, col, value) {
    await fetch('/set_value', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ sheet_name: sheetName, row, col, value })
    });
}

async function addSheet() {
    const name = prompt('Nombre de la nueva hoja:');
    const rows = parseInt(prompt('Número de filas:'));
    const cols = parseInt(prompt('Número de columnas:'));
    await fetch('/add_sheet', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ name, rows, cols })
    });
    loadSheets();
}

async function renameSheet() {
    const oldName = prompt('Nombre actual de la hoja:');
    const newName = prompt('Nuevo nombre de la hoja:');
    await fetch('/rename_sheet', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ old_name: oldName, new_name: newName })
    });
    loadSheets();
}

async function addRow() {
    const sheetName = getActiveSheetName();
    await fetch('/add_row', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ sheet_name: sheetName })
    });
    loadSheet(sheetName);
}

async function addCol() {
    const sheetName = getActiveSheetName();
    await fetch('/add_col', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ sheet_name: sheetName })
    });
    loadSheet(sheetName);
}

async function loadSheets() {
    const response = await fetch('/get_sheets');
    const sheets = await response.json();
    renderSheets(sheets);
}

function renderSheets(sheets) {
    const sheetsContainer = document.getElementById('sheets-container');
    sheetsContainer.innerHTML = '';
    sheets.forEach(sheetName => {
        const sheetTab = document.createElement('div');
        sheetTab.classList.add('sheet-tab');
        sheetTab.textContent = sheetName;
        sheetTab.addEventListener('click', () => loadSheet(sheetName));
        sheetsContainer.appendChild(sheetTab);
    });
    setActiveSheet(getActiveSheetName());
}

async function setActiveSheet(sheetName) {
    await fetch('/set_active_sheet', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ sheet_name: sheetName })
    });
    document.querySelectorAll('.sheet-tab').forEach(tab => {
        tab.classList.remove('active');
        if (tab.textContent === sheetName) {
            tab.classList.add('active');
        }
    });
}

function getActiveSheetName() {
    const activeTab = document.querySelector('.sheet-tab.active');
    return activeTab ? activeTab.textContent : 'Sheet1';
}