document.addEventListener("DOMContentLoaded", () => {
    const spreadsheet = document.getElementById("spreadsheet");
    const headers = document.getElementById("headers");
    const themeToggle = document.getElementById("themeToggle");
    const spreadsheetTitle = document.getElementById("spreadsheetTitle");
    const csvInput = document.getElementById("csvInput");
    const exportCsvButton = document.getElementById("exportCsv");

    const customFunctions = {}; // Stores custom functions as { FunctionName: Expression }

    // Check if the themeToggle button exists
    if (!themeToggle) {
        console.error("Theme toggle button not found in the DOM.");
        return;
    }

    // Function to toggle dark theme
    themeToggle.addEventListener("click", () => {
        document.body.classList.toggle("dark-theme");
    });

    // Function to evaluate formulas, including SUM, MIN, MAX, and AVG for rows and columns
    function evaluateFormula(cell, formula) {
        try {
            const updatedFormula = formula.trim().toUpperCase();

            // Match SUM, MIN, MAX, or AVG for a column (e.g., SUM(E)) or a row (e.g., SUM(1))
            const match = updatedFormula.match(/^(SUM|MIN|MAX|AVG)\((\w+)\)$/);
            if (match) {
                const operation = match[1]; // SUM, MIN, MAX, or AVG
                const target = match[2]; // Column letter or row number
                const values = [];

                if (isNaN(target)) {
                    // Target is a column (e.g., "E")
                    const columnIndex = target.charCodeAt(0) - 64; // Convert column letter to index (1-based)
                    Array.from(spreadsheet.rows).forEach((row, rowIndex) => {
                        if (rowIndex > 0) { // Skip the header row
                            const targetCell = row.cells[columnIndex];
                            if (targetCell) {
                                const value = parseFloat(targetCell.textContent.trim());
                                if (!isNaN(value)) {
                                    values.push(value);
                                }
                            }
                        }
                    });
                } else {
                    // Target is a row (e.g., "1")
                    const rowIndex = parseInt(target, 10); // Convert row number to index
                    const row = spreadsheet.rows[rowIndex];
                    if (row) {
                        Array.from(row.cells).forEach((targetCell, colIndex) => {
                            if (colIndex > 0) { // Skip the row label column
                                const value = parseFloat(targetCell.textContent.trim());
                                if (!isNaN(value)) {
                                    values.push(value);
                                }
                            }
                        });
                    }
                }

                // Perform the operation
                let result;
                if (operation === "SUM") {
                    result = values.reduce((sum, value) => sum + value, 0);
                } else if (operation === "MIN") {
                    result = values.length > 0 ? Math.min(...values) : "N/A";
                } else if (operation === "MAX") {
                    result = values.length > 0 ? Math.max(...values) : "N/A";
                } else if (operation === "AVG") {
                    console.log(values);
                    result = values.length > 0 ? (values.reduce((sum, value) => sum + value, 0) / values.length) : "N/A";
                }

                cell.textContent = result; // Update the cell with the calculated result
                return;
            }

            throw new Error("Unsupported formula syntax");
        } catch (error) {
            console.error(`Error evaluating formula: ${formula}`, error);
            cell.textContent = `Error: ${error.message}`; // Display the error message in the cell
        }
    }

    function evaluateGeneralFormula(cell, formula) {
        console.log(`Evaluating general formula: ${formula} in cell:`, cell);

        try {
            const rowIndex = cell.parentElement.rowIndex; // Get the row index of the cell
            const row = spreadsheet.rows[rowIndex]; // Get the row element

            const updatedFormula = formula.replace(/[A-Z]/g, column => {
                const columnIndex = column.charCodeAt(0) - 65 + 1; // Convert column letter to index
                const referencedCell = row.cells[columnIndex]; // Get the cell in the same row
                const value = referencedCell ? (parseFloat(referencedCell.textContent) || 0) : 0; // Use 0 if the cell is empty
                console.log(`Extracted value for column ${column} in row ${rowIndex}:`, value);
                return value;
            });

            console.log(`Updated formula after replacing cell references: ${updatedFormula}`);
            const result = eval(updatedFormula);
            console.log(`Result of general formula: ${result}`);
            cell.textContent = result; // Update the cell with the calculated result
        } catch (error) {
            console.error(`Error evaluating general formula: ${formula}`, error);
            cell.textContent = "ERROR"; // Display "ERROR" if the formula is invalid
        }
    }

    // Helper function to calculate sum, min, or max
    function calculate(values, operation) {
        console.log(`Calculating ${operation} for values:`, values);

        // Filter out blank or non-numerical values
        const filteredValues = values.filter(value => typeof value === "number" && !isNaN(value));
        console.log(`Filtered values for ${operation}:`, filteredValues);

        if (operation === "sum") {
            const result = filteredValues.reduce((total, value) => total + value, 0); // Sum of valid values
            console.log(`Result of SUM: ${result}`);
            return result;
        } else if (operation === "min") {
            const result = filteredValues.length > 0 ? Math.min(...filteredValues) : "N/A"; // Minimum of valid values
            console.log(`Result of MIN: ${result}`);
            return result;
        } else if (operation === "max") {
            const result = filteredValues.length > 0 ? Math.max(...filteredValues) : "N/A"; // Maximum of valid values
            console.log(`Result of MAX: ${result}`);
            return result;
        }
        console.error(`Invalid operation: ${operation}`);
        return "N/A"; // Default return value if no valid operation is provided
    }

    function adjustFontSize(cell) {
        const maxFontSize = 16; // Maximum font size in pixels
        const minFontSize = 8; // Minimum font size in pixels
        let fontSize = maxFontSize;

        // Temporarily set the font size to the maximum
        cell.style.fontSize = `${fontSize}px`;

        // Reduce the font size until the content fits or the minimum font size is reached
        while (fontSize > minFontSize && (cell.scrollWidth > cell.clientWidth || cell.scrollHeight > cell.clientHeight)) {
            fontSize--;
            cell.style.fontSize = `${fontSize}px`;
        }
    }

    // Function to handle cell blur (when editing is finished)
    function handleCellBlur(event) {
        const cell = event.target;
        const value = cell.textContent.trim();

        if (value.startsWith("=")) {
            const formula = value.slice(1); // Remove "="

            // Check if the formula contains the "=>" syntax
            const rangeMatch = formula.match(/(.*)=>(\d+)$/);
            if (rangeMatch) {
                const operation = rangeMatch[1].trim(); // Extract the operation (e.g., "B+C*D")
                const targetRow = parseInt(rangeMatch[2], 10); // Extract the target row (e.g., "5")
                executeOperationInRange(cell, operation, targetRow);
            } else if (/^(SUM|MIN|MAX|AVG)\(/i.test(formula)) {
                // If the formula starts with SUM, MIN, MAX, or AVG, call evaluateFormula
                evaluateFormula(cell, formula);
            } else if (/^[A-Z]+\d+([\+\-\*\/][A-Z]+\d+)+$/i.test(formula)) {
                // If the formula matches an arithmetic pattern (e.g., A1+B1), treat it as an arithmetic formula
                evaluateFormula(cell, formula);
            } else {
                // Otherwise, treat it as a general formula
                evaluateGeneralFormula(cell, formula);
            }
        }

        // Adjust the font size of the cell after updating its content
        adjustFontSize(cell);
    }

    function executeOperationInRange(cell, operation, targetRow) {
        const startRow = cell.parentElement.rowIndex; // Get the current row index
        const endRow = Math.min(targetRow, spreadsheet.rows.length - 1); // Ensure the target row is within bounds

        for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            const row = spreadsheet.rows[rowIndex];
            const updatedFormula = operation.replace(/[A-Z]/g, column => {
                const columnIndex = column.charCodeAt(0) - 65 + 1; // Convert column letter to index
                const referencedCell = row.cells[columnIndex]; // Get the cell in the same row
                return referencedCell ? (parseFloat(referencedCell.textContent) || 0) : 0; // Use 0 if the cell is empty
            });

            try {
                // Evaluate the mathematical expression
                const result = eval(updatedFormula);
                const targetCell = row.cells[cell.cellIndex]; // Update the same column in each row
                targetCell.textContent = result; // Update the cell with the calculated result
                adjustFontSize(targetCell); // Adjust font size for the updated cell
            } catch (error) {
                console.error(`Error evaluating formula in row ${rowIndex}:`, error);
            }
        }
    }

    // Add a new row below the clicked row
    function addRow(rowIndex) {
        const newRow = document.createElement("tr");
        const rowNumber = rowIndex + 2; // New row number (1-based index)

        // Create the row label
        newRow.innerHTML = `<td class="row-label">${rowNumber}<div class="hover-buttons"><button class="addRow">+</button><button class="deleteRow">−</button></div></td>`;

        // Add empty cells for each column in the header row (excluding the row label)
        const columnCount = spreadsheet.rows[0].cells.length - 1; // Exclude the row label column
        for (let i = 0; i < columnCount; i++) {
            const newCell = document.createElement("td");
            newCell.contentEditable = "true";
            newRow.appendChild(newCell);
        }

        // Insert the new row manually
        const rows = Array.from(spreadsheet.tBodies[0].rows);
        if (rowIndex < rows.length - 1) {
            rows[rowIndex].after(newRow); // Use `after` to insert the new row below the clicked row
        } else {
            spreadsheet.tBodies[0].appendChild(newRow); // Append to the bottom if it's the last row
        }

        updateRowLabels();
        addResizingHandles(); // Reapply resizing handles to include the new row
    }

    // Add a new column to the right of the clicked column
    function addColumn(columnIndex) {
        const adjustedColumnIndex = columnIndex + 1; // Insert at the next index (to the right of the clicked column)

        // Create the new header cell
        const newHeader = document.createElement("th");
        const columnLabel = String.fromCharCode(64 + adjustedColumnIndex); // Generate column label (A, B, C, etc.)
        newHeader.className = "header-cell";
        newHeader.innerHTML = `${columnLabel}<div class="hover-buttons"><button class="addColumn">+</button><button class="deleteColumn">−</button></div>`;

        // Insert the new header at the correct position
        const headers = spreadsheet.rows[0];
        const headerCells = Array.from(headers.cells);
        if (adjustedColumnIndex <= headerCells.length) {
            headers.insertBefore(newHeader, headerCells[adjustedColumnIndex]); // Insert before the next column
        } else {
            headers.appendChild(newHeader); // Append to the end if it's the last column
        }

        // Add a new cell to each row at the correct position
        Array.from(spreadsheet.tBodies[0].rows).forEach((row) => {
            const newCell = document.createElement("td");
            newCell.contentEditable = "true";
            const rowCells = Array.from(row.cells);
            if (adjustedColumnIndex <= rowCells.length) {
                row.insertBefore(newCell, rowCells[adjustedColumnIndex]); // Insert before the next column
            } else {
                row.appendChild(newCell); // Append to the end if it's the last column
            }
        });

        updateColumnLabels(); // Update column labels dynamically
        addResizingHandles(); // Reapply resizing handles to include the new column
    }

    // Delete a row and update row labels
    function deleteRow(targetRow) {
        if (spreadsheet.rows.length > 2) { // Ensure at least one data row remains
            targetRow.remove();
            updateRowLabels();
            addResizingHandles(); // Reapply resizing handles after removing a row
        }
    }

    // Delete a column and update column labels
    function deleteColumn(targetColumnIndex) {
        if (spreadsheet.rows[0].cells.length > 2) { // Ensure at least one data column remains
            // Delete the header cell
            spreadsheet.rows[0].deleteCell(targetColumnIndex);

            // Delete the corresponding cell in each row
            Array.from(spreadsheet.tBodies[0].rows).forEach(row => {
                row.deleteCell(targetColumnIndex);
            });

            updateColumnLabels();
            addResizingHandles(); // Reapply resizing handles after removing a column
        }
    }

    // Update row labels
    function updateRowLabels() {
        Array.from(spreadsheet.rows).forEach((row, index) => {
            if (index === 0) return; // Skip the header row
            const rowLabel = row.cells[0];
            if (rowLabel) {
                rowLabel.innerHTML = `${index}<div class="hover-buttons"><button class="addRow">+</button><button class="deleteRow">−</button></div>`;
                rowLabel.className = "row-label";
            }
        });
    }

    // Update column labels
    function updateColumnLabels() {
        const headerRow = spreadsheet.rows[0];
        if (headerRow) {
            Array.from(headerRow.cells).forEach((cell, index) => {
                if (index > 0) { // Skip the row label column
                    cell.innerHTML = `${String.fromCharCode(64 + index)}<div class="hover-buttons"><button class="addColumn">+</button><button class="deleteColumn">−</button></div>`;
                    cell.className = "header-cell";
                }
            });
        }
    }

    // Event listener for cell blur
    spreadsheet.addEventListener("blur", handleCellBlur, true);

    // Event listener for button clicks
    document.addEventListener("click", (e) => {
        if (e.target.classList.contains("addRow")) {
            const rowIndex = Array.from(spreadsheet.tBodies[0].rows).indexOf(e.target.closest("tr"));
            addRow(rowIndex); // Pass the index of the clicked row
        }
        if (e.target.classList.contains("deleteRow")) {
            const targetRow = e.target.closest("tr");
            deleteRow(targetRow);
        }
        if (e.target.classList.contains("addColumn")) {
            const columnIndex = Array.from(spreadsheet.rows[0].cells).indexOf(e.target.closest("th")); // 1-based index
            addColumn(columnIndex); // Pass the index of the clicked column
        }
        if (e.target.classList.contains("deleteColumn")) {
            const targetColumnIndex = Array.from(spreadsheet.rows[0].cells).indexOf(e.target.closest("th"));
            deleteColumn(targetColumnIndex);
        }
    });

    // Function to get the spreadsheet title
    function getSpreadsheetTitle() {
        return spreadsheetTitle.value.trim() || "Untitled Spreadsheet";
    }

    // Function to export the spreadsheet to CSV
    function exportToCsv() {
        const rows = Array.from(spreadsheet.rows);
        const csvContent = rows.map((row, rowIndex) => {
            return Array.from(row.cells)
                .filter((cell, colIndex) => !(rowIndex === 0 && colIndex === 0)) // Skip the top-left corner cell
                .filter((cell, colIndex) => rowIndex !== 0 || colIndex > 0) // Skip row labels and column headers
                .map(cell => cell.textContent.trim()) // Get the text content of each cell
                .join(","); // Join cells with commas
        }).join("\n"); // Join rows with newlines

        const blob = new Blob([csvContent], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `${getSpreadsheetTitle()}.csv`;
        a.click();
        URL.revokeObjectURL(url);
    }

    // Import a CSV file into the spreadsheet
    function importFromCsv(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function (e) {
            const csvContent = e.target.result.split("\n").map(row => row.split(","));

            // Clear existing rows and columns
            while (spreadsheet.rows.length > 0) {
                spreadsheet.deleteRow(0);
            }

            // Rebuild the header row
            const headerRow = spreadsheet.insertRow();
            headerRow.id = "headers";
            headerRow.insertCell(0); // Add an empty cell for the row labels
            for (let i = 0; i < csvContent[0].length; i++) {
                const headerCell = document.createElement("th");
                headerCell.className = "header-cell";
                headerCell.innerHTML = `${String.fromCharCode(65 + i)}<div class="hover-buttons"><button class="addColumn">+</button><button class="deleteColumn">−</button></div>`;
                headerRow.appendChild(headerCell);
            }

            // Add rows from the CSV
            csvContent.forEach((rowValues, rowIndex) => {
                const newRow = spreadsheet.insertRow();
                const rowLabel = newRow.insertCell(0);
                rowLabel.className = "row-label";
                rowLabel.innerHTML = `${rowIndex + 1}<div class="hover-buttons"><button class="addRow">+</button><button class="deleteRow">−</button></div>`;

                rowValues.forEach(value => {
                    const newCell = newRow.insertCell();
                    newCell.contentEditable = "true";
                    newCell.textContent = value.trim();
                });
            });

            updateRowLabels();
            updateColumnLabels();
            addResizingHandles(); // Reapply resizing handles after importing
        };

        reader.readAsText(file);
    }

    // Event listener for the Export button
    exportCsvButton.addEventListener("click", exportToCsv);

    // Event listener for the Import button
    csvInput.addEventListener("change", importFromCsv);

    // Initialize a 7x7 spreadsheet
    function initializeSpreadsheet() {
        for (let i = 1; i <= 7; i++) addColumn(i - 1);
        for (let i = 1; i <= 7; i++) addRow(i - 1);

        // Adjust font size for all cells
        Array.from(spreadsheet.getElementsByTagName("td")).forEach(cell => {
            adjustFontSize(cell);
        });
    }

    initializeSpreadsheet();

    spreadsheet.addEventListener("input", (event) => {
        const cell = event.target;
        if (cell.tagName === "TD" && cell.isContentEditable) {
            adjustFontSize(cell);
        }
    });

    document.getElementById("exportCsv").addEventListener("click", exportToCsv);
    document.getElementById("csvInput").addEventListener("change", importFromCsv);

    // Function to handle paste events
    function handlePaste(event) {
        event.preventDefault(); // Prevent the default paste behavior

        const clipboardData = event.clipboardData || window.clipboardData;
        const pastedData = clipboardData.getData("text/plain"); // Get the plain text data from the clipboard

        const rows = pastedData.split("\n").map(row => row.split("\t")); // Parse rows and columns (tab-separated)

        const targetCell = event.target; // The cell where the paste occurred
        const startRowIndex = targetCell.parentElement.rowIndex; // Get the row index of the target cell
        const startColIndex = targetCell.cellIndex; // Get the column index of the target cell

        rows.forEach((rowData, rowOffset) => {
            const currentRowIndex = startRowIndex + rowOffset;
            if (currentRowIndex >= spreadsheet.rows.length) {
                addRow(spreadsheet.rows.length - 1); // Add a new row if needed
            }

            const currentRow = spreadsheet.rows[currentRowIndex];
            rowData.forEach((cellData, colOffset) => {
                const currentColIndex = startColIndex + colOffset;
                if (currentColIndex >= currentRow.cells.length) {
                    addColumn(currentRow.cells.length - 1); // Add a new column if needed
                }

                const currentCell = currentRow.cells[currentColIndex];
                currentCell.textContent = cellData.trim(); // Set the cell content
            });
        });

        updateRowLabels(); // Update row labels after pasting
        updateColumnLabels(); // Update column labels after pasting
    }

    // Add event listener for paste events
    spreadsheet.addEventListener("paste", handlePaste);

    // Get the modal, button, and close elements
    const helpModal = document.getElementById("helpModal");
    const helpButton = document.getElementById("helpButton");
    const closeButton = document.querySelector(".modal-content .close");

    // Show the modal when the Help button is clicked
    helpButton.addEventListener("click", () => {
        helpModal.style.display = "block";
    });

    // Hide the modal when the close button is clicked
    closeButton.addEventListener("click", () => {
        helpModal.style.display = "none";
    });

    // Hide the modal when clicking outside the modal content
    window.addEventListener("click", (event) => {
        if (event.target === helpModal) {
            helpModal.style.display = "none";
        }
    });

    let isResizing = false;
    let startX, startY, startWidth, startHeight, targetElement;

    // Add resizing handles to each column header
    Array.from(spreadsheet.querySelectorAll("th")).forEach((th) => {
        const resizeHandle = document.createElement("div");
        resizeHandle.className = "column-resize-handle";
        th.appendChild(resizeHandle);

        resizeHandle.addEventListener("mousedown", (e) => {
            isResizing = true;
            startX = e.clientX;
            startWidth = th.offsetWidth;
            targetElement = th;
            document.body.style.cursor = "col-resize";
        });
    });

    // Add resizing handles to each row label
    Array.from(spreadsheet.querySelectorAll(".row-label")).forEach((td) => {
        const resizeHandle = document.createElement("div");
        resizeHandle.className = "row-resize-handle";
        td.appendChild(resizeHandle);

        resizeHandle.addEventListener("mousedown", (e) => {
            isResizing = true;
            startY = e.clientY;
            startHeight = td.offsetHeight;
            targetElement = td;
            document.body.style.cursor = "row-resize";
        });
    });

    // Handle mouse movement for resizing
    document.addEventListener("mousemove", (e) => {
        if (!isResizing) return;

        const standardWidth = 100; // Default column width
        const standardHeight = 40; // Default row height
        const snapThreshold = 20; // Snap when within 10px of the standard size

        if (targetElement.tagName === "TH") {
            // Column resizing
            const deltaX = e.clientX - startX;
            let newWidth = startWidth + deltaX;

            // Snap to the standard width if close enough
            if (Math.abs(newWidth - standardWidth) <= snapThreshold) {
                newWidth = standardWidth;
            }

            targetElement.style.width = `${Math.max(newWidth, 50)}px`; // Minimum width of 50px
        } else if (targetElement.classList.contains("row-label")) {
            // Row resizing
            const deltaY = e.clientY - startY;
            let newHeight = startHeight + deltaY;

            // Snap to the standard height if close enough
            if (Math.abs(newHeight - standardHeight) <= snapThreshold) {
                newHeight = standardHeight;
            }

            targetElement.parentElement.style.height = `${Math.max(newHeight, 20)}px`; // Minimum height of 20px
        }
    });

    // Stop resizing on mouseup
    document.addEventListener("mouseup", () => {
        if (isResizing) {
            isResizing = false;
            document.body.style.cursor = "default";
        }
    });

    // Add resizing handles dynamically for all rows and columns
    function addResizingHandles() {
        // Add resizing handles to column headers
        Array.from(spreadsheet.querySelectorAll("th")).forEach((th) => {
            if (!th.querySelector(".column-resize-handle")) {
                const resizeHandle = document.createElement("div");
                resizeHandle.className = "column-resize-handle";
                th.appendChild(resizeHandle);

                // Attach the mousedown event listener for column resizing
                resizeHandle.addEventListener("mousedown", (e) => {
                    isResizing = true;
                    startX = e.clientX;
                    startWidth = th.offsetWidth;
                    targetElement = th;
                    document.body.style.cursor = "col-resize";
                });
            }
        });

        // Add resizing handles to row labels
        Array.from(spreadsheet.querySelectorAll(".row-label")).forEach((td) => {
            if (!td.querySelector(".row-resize-handle")) {
                const resizeHandle = document.createElement("div");
                resizeHandle.className = "row-resize-handle";
                td.appendChild(resizeHandle);

                // Attach the mousedown event listener for row resizing
                resizeHandle.addEventListener("mousedown", (e) => {
                    isResizing = true;
                    startY = e.clientY;
                    startHeight = td.offsetHeight;
                    targetElement = td;
                    document.body.style.cursor = "row-resize";
                });
            }
        });
    }

    // Call the function initially to add handles to existing rows and columns
    addResizingHandles();

    // Add event listeners for adding rows and columns
    document.addEventListener("click", (e) => {
        if (e.target.classList.contains("addRow")) {
            const rowIndex = Array.from(spreadsheet.tBodies[0].rows).indexOf(e.target.closest("tr"));
            addRow(rowIndex);
            addResizingHandles(); // Add resizing handles to the new row
        }
        if (e.target.classList.contains("addColumn")) {
            const columnIndex = Array.from(spreadsheet.querySelectorAll("th")).indexOf(e.target.closest("th")) - 1;
            addColumn(columnIndex);
            addResizingHandles(); // Add resizing handles to the new column
        }
    });
});