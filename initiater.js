import {Emaker} from './eMaker.js'

class ExcelBorderHighlighter {
    constructor(excelContainer) {
        this.excelContainer = excelContainer;
        this.selectedDiv = null;
        this.handleClick = this.handleClick.bind(this);
        this.addClickListener();
    }
  
    handleClick(event) {
        if (this.selectedDiv) {
            this.selectedDiv.style.border = '1px solid black';
        }
  
        const targetDiv = event.target.closest('div[aria-rowindex][aria-colindex]');
        
        if (targetDiv) {
            targetDiv.style.border = '1px solid red';
            this.selectedDiv = targetDiv;
        }
    }
  
    addClickListener() {
        this.excelContainer.addEventListener('click', this.handleClick);
    }
}

class Excel {
    constructor(rowContainer, row, col) {
        this.rowContainer = rowContainer;
        this.col = col;
        this.row = row;
        this.init();
    }

    init() {
        this.constructExcel();
    }
 
    constructExcel() {
        const excel = document.createElement('div');
        excel.className = 'excel resizable';
        excel.setAttribute('id', `rowCol${this.row}_${this.col}`);
        excel.setAttribute('role', 'gridcell');
        excel.setAttribute('aria-rowindex', this.row);
        excel.setAttribute('aria-colindex', this.col);
        excel.style.flex = '1';
        this.rowContainer.appendChild(excel);
        this.element = excel;
        new Emaker(excel, this.row, this.col);
        
    }
}

class Grid_maker {
    constructor(mainContainer, maxRow, maxCol) {
        this.mainContainer = mainContainer;
        this.maxRow = maxRow;
        this.maxCol = maxCol;
        this.currentRowCount = 0;
        this.rowArr = [];
        this.init();
    }

    addNewRow() {
        if (this.currentRowCount >= this.maxRow) {
            alert("No more rows can be added");
            return;
        }

        this.currentRowCount += 1;
        const row = document.createElement('div');
        row.className = 'row';
        row.id = `row_${this.currentRowCount}`;
        row.style.flex = '1';

        const excel = new Excel(row, this.currentRowCount, 1);
        this.rowArr[this.currentRowCount - 1] = [excel];
        this.mainContainer.appendChild(row);
    }

    addNewCol(rowNum) {
        if (rowNum > this.currentRowCount) return;
        let colCount = this.rowArr[rowNum - 1].length;
        if (colCount >= this.maxCol) {
            alert("No more columns can be added");
            return;
        }

        colCount += 1;
        const row = document.getElementById(`row_${rowNum}`);
        const excel = new Excel(row, rowNum, colCount);
        this.rowArr[rowNum - 1].push(excel);
    }

    addResizeHandles() {
        this.rowArr.forEach((row, rowIndex) => {
            const rowElement = document.getElementById(`row_${rowIndex + 1}`);

            // Add row resize handle
            if (rowIndex < this.rowArr.length - 1) {
                const rowResizeHandle = document.createElement('div');
                rowResizeHandle.className = 'row-resize-handle';
                rowElement.appendChild(rowResizeHandle);
            }

            row.forEach((cell, colIndex) => {
                // Add column resize handle
                if (colIndex < row.length - 1) {
                    const colResizeHandle = document.createElement('div');
                    colResizeHandle.className = 'col-resize-handle';
                    cell.element.appendChild(colResizeHandle);
                }
            });
        });
    }

    handleResize() {
        let isResizing = false;
        let currentElement = null;
        let startX, startY, startWidth, startHeight;
        let resizeType = '';

        const startResize = (e) => {
            if (e.target.classList.contains('col-resize-handle')) {
                currentElement = e.target.closest('.excel');
                resizeType = 'column';
                
            } else if (e.target.classList.contains('row-resize-handle')) {
                currentElement = e.target.closest('.row');
                resizeType = 'row';
            } else {
                return;
            }

            isResizing = true;
            startX = e.clientX;
            startY = e.clientY;
            startWidth = currentElement.offsetWidth;
            startHeight = currentElement.offsetHeight;

            document.addEventListener('mousemove', resize);
            document.addEventListener('mouseup', stopResize);
            e.preventDefault();
        };

        const resize = (e) => {
            if (!isResizing) return;

            if (resizeType === 'column') {
                const width = startWidth + (e.clientX - startX);
                currentElement.style.width = `${width}px`;
                currentElement.style.flex = 'none'; // Override flex settings
            } else if (resizeType === 'row') {
                const height = startHeight + (e.clientY - startY);
                // Ensure the height does not go below a minimum value (e.g., 50px)
                const newHeight = Math.max(height, 50);
                currentElement.style.height = `${newHeight}px`;
                currentElement.style.flex = 'none'; // Override flex settings
            }

            // Force a reflow/repaint to make sure the changes are applied
            currentElement.style.display = 'none';
            currentElement.offsetHeight; // Trigger reflow
            currentElement.style.display = '';
        };

        const stopResize = () => {
            isResizing = false;
            currentElement = null;
            document.removeEventListener('mousemove', resize);
            document.removeEventListener('mouseup', stopResize);
        };

        this.mainContainer.addEventListener('mousedown', startResize);
    }

    init() {
        this.mainContainer.style.display = 'flex';
        this.mainContainer.style.flexDirection = 'column';
        this.addNewRow();
        this.addNewCol(1);
        this.addNewRow();
        this.addResizeHandles();
        this.handleResize();
        this.borderHighlighter = new ExcelBorderHighlighter(this.mainContainer);
    }
}

function init(mainContainer) {
    new Grid_maker(mainContainer, 3, 3);
}

document.addEventListener('DOMContentLoaded', () => {
    let mainContainer = document.getElementById("mainContainer");
    init(mainContainer);
});

// document.querySelector('form').addEventListener('submit', async (e) => {
//     e.preventDefault();
    
//     const formData = new FormData(e.target);
//     const response = await fetch('http://localhost:5228/api/Data/uploadAndCreateTable', {
//         method: 'POST',
//         body: formData
//     });
    
//     if (response.ok) {
//         const data = await response.json();
//         showTableCreationPopup(data.columns);
//     } else {
//         alert('Error uploading file');
//     }
// });

// function showTableCreationPopup(columns) {
//     const popup = document.createElement('div');
//     popup.className = 'popup';
    
//     const form = document.createElement('form');
//     form.innerHTML = `
//         <h2>Create Table</h2>
//         <label for="tableName">Table Name:</label>
//         <input type="text" id="tableName" required>
//         <table>
//             <thead>
//                 <tr>
//                     <th>Column Name</th>
//                     <th>Type</th>
//                     <th>Allow Null</th>
//                     <th>Primary Key</th>
//                 </tr>
//             </thead>
//             <tbody>
//                 ${columns.map((column, index) => `
//                     <tr>
//                         <td>${column}</td>
//                         <td>
//                             <select name="type_${index}" required>
//                                 <option value="INT">INT</option>
//                                 <option value="VARCHAR(255)">VARCHAR(255)</option>
//                                 <option value="DATE">DATE</option>
//                                 <option value="DECIMAL(10,2)">DECIMAL(10,2)</option>
//                             </select>
//                         </td>
//                         <td>
//                             <input type="checkbox" name="null_${index}" checked>
//                         </td>
//                         <td>
//                             <input type="radio" name="primary_key" value="${index}">
//                         </td>
//                     </tr>
//                 `).join('')}
//             </tbody>
//         </table>
//         <button type="submit">Create Table</button>
//     `;
    
//     popup.appendChild(form);
//     document.body.appendChild(popup);
    
//     form.addEventListener('submit', async (e) => {
//         e.preventDefault();
        
//         const tableName = document.getElementById('tableName').value;
//         const columnDefinitions = columns.map((column, index) => ({
//             name: column,
//             type: document.querySelector(`[name="type_${index}"]`).value,
//             allowNull: document.querySelector(`[name="null_${index}"]`).checked,
//             isPrimaryKey: document.querySelector(`[name="primary_key"]:checked`)?.value == index
//         }));
        
//         const response = await fetch('http://localhost:5228/api/Data/createTable', {
//             method: 'POST',
//             headers: {
//                 'Content-Type': 'application/json'
//             },
//             body: JSON.stringify({
//                 tableName,
//                 columns: columnDefinitions
//             })
//         });
        
//         if (response.ok) {
//             alert('Table created successfully');
//             popup.remove();
//         } else {
//             alert('Error creating table');
//         }
//     });
// }


document.querySelector('form').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const formData = new FormData(e.target);
    const response = await fetch('http://localhost:5228/api/Data/uploadAndCreateTable', {
        method: 'POST',
        body: formData
    });
    
    if (response.ok) {
        const data = await response.json();
        showTableCreationPopup(data.columns,data.tempFilePath);
    } else {
        alert('Error uploading file');
    }
});

function showTableCreationPopup(columns,tempFilePath) {
    const popup = document.createElement('div');
    popup.className = 'popup';
    
    const form = document.createElement('form');
    form.innerHTML = `
        <h2>Create Table</h2>
        <label for="tableName">Table Name:</label>
        <input type="text" id="tableName" required>
        <table>
            <thead>
                <tr>
                    <th>Column Name</th>
                    <th>Type</th>
                    <th>Allow Null</th>
                    <th>Primary Key</th>
                </tr>
            </thead>
            <tbody>
                ${columns.map((column, index) => `
                    <tr>
                        <td>${column}</td>
                        <td>
                            <select name="type_${index}" required>
                                <option value="INT">INT</option>
                                <option value="VARCHAR(255)">VARCHAR(255)</option>
                                <option value="DATE">DATE</option>
                                <option value="DECIMAL(10,2)">DECIMAL(10,2)</option>
                            </select>
                        </td>
                        <td>
                            <input type="checkbox" name="null_${index}" checked>
                        </td>
                        <td>
                            <input type="radio" name="primary_key" value="${index}">
                        </td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
        <button type="submit">Create Table and Upload Data</button>
    `;
    
    popup.appendChild(form);
    document.body.appendChild(popup);
    
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const tableName = document.getElementById('tableName').value;
        const columnDefinitions = columns.map((column, index) => ({
            name: column,
            type: document.querySelector(`[name="type_${index}"]`).value,
            allowNull: document.querySelector(`[name="null_${index}"]`).checked,
            isPrimaryKey: document.querySelector(`[name="primary_key"]:checked`)?.value == index
        }));
        
        const response = await fetch('http://localhost:5228/api/Data/createTableAndUpload1', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                tableName,
                columns: columnDefinitions,
                tempFilePath: tempFilePath, 
            })
        });
        
        if (response.ok) {
            alert('Table created and data upload started');
            popup.remove();
            checkMiscellaneousRows(tableName);
        } else {
            alert('Error creating table and uploading data');
        }
    });
}


async function checkMiscellaneousRows(tableName) {
    try {
        const response = await fetch('http://localhost:5228/api/Data/checkMiscellaneousRows', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(tableName)
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        if (data.hasMiscellaneousRows) {
            showMiscellaneousRowsPopup(tableName);
        } else {
            alert('Data upload completed successfully');
        }
    } catch (error) {
        console.error('Error checking for miscellaneous rows:', error);
        alert('Error checking for miscellaneous rows');
    }
}


function showMiscellaneousRowsPopup(tableName) {
    const popup = document.createElement('div');
    popup.className = 'popup';
    
    const form = document.createElement('form');
    form.innerHTML = `
        <h2>Miscellaneous Rows Detected</h2>
        <p>Some rows have more columns than expected. How would you like to handle them?</p>
        <select id="miscAction">
            <option value="delete">Delete these rows</option>
            <option value="alterTable">Alter table to add new columns</option>
            <option value="truncate">Add only up to specified columns</option>
        </select>
        <div id="newColumnsDiv" style="display:none;">
            <label for="newColumns">New Column Names (comma-separated):</label>
            <input type="text" id="newColumns">
        </div>
        <div id="columnCountDiv" style="display:none;">
            <label for="columnCount">Number of columns to keep:</label>
            <input type="number" id="columnCount">
        </div>
        <button type="submit">Process Miscellaneous Rows</button>
    `;
    
    popup.appendChild(form);
    document.body.appendChild(popup);

    const actionSelect = document.getElementById('miscAction');
    const newColumnsDiv = document.getElementById('newColumnsDiv');
    const columnCountDiv = document.getElementById('columnCountDiv');

    actionSelect.addEventListener('change', () => {
        newColumnsDiv.style.display = actionSelect.value === 'alterTable' ? 'block' : 'none';
        columnCountDiv.style.display = actionSelect.value === 'truncate' ? 'block' : 'none';
    });
    
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const action = actionSelect.value;
        const newColumns = document.getElementById('newColumns').value.split(',').map(c => c.trim());
        const columnCount = parseInt(document.getElementById('columnCount').value);

        try {
            const response = await fetch('http://localhost:5228/api/Data/handleMiscellaneousRows', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    action,
                    tableName,
                    newColumns,
                    columnCount
                })
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            alert('Miscellaneous rows processed successfully');
            popup.remove();
        } catch (error) {
            console.error('Error processing miscellaneous rows:', error);
            alert('Error processing miscellaneous rows');
        }
    });
}