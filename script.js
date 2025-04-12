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

    // Function to evaluate formulas
    function evaluateFormula(cell, formula) {
        console.log(`Evaluating formula: ${formula} in cell:`, cell);

        try {
            // Check if the formula contains SUM, MIN, or MAX
            const match = formula.match(/(sum|min|max)\((\w+)\)/i);
            if (match) {
                console.log(`Detected aggregate function: ${match[1]} with argument: ${match[2]}`);
                const operation = match[1].toLowerCase(); // sum, min, max
                const argument = match[2]; // Row number or column letter

                if (!isNaN(argument)) {
                    console.log(`Processing row-based operation for row: ${argument}`);
                    const rowIndex = parseInt(argument, 10);
                    const row = spreadsheet.rows[rowIndex];
                    if (row) {
                        const values = Array.from(row.cells)
                            .slice(1) // Skip the row label
                            .map(cell => parseFloat(cell.textContent) || NaN); // Parse cell values
                        console.log(`Extracted values from row ${argument}:`, values);
                        const result = calculate(values, operation);
                        console.log(`Result of ${operation} for row ${argument}:`, result);
                        cell.textContent = result !== null ? result : "N/A"; // Display "N/A" if no valid values
                        return;
                    } else {
                        console.error(`Row ${argument} does not exist`);
                        cell.textContent = "Row does not exist"; // Specific error for invalid row
                        return;
                    }
                } else if (/^[A-Z]+$/i.test(argument)) {
                    console.log(`Processing column-based operation for column: ${argument}`);
                    const colIndex = argument.toUpperCase().charCodeAt(0) - 65; // Convert column letter to index
                    const values = Array.from(spreadsheet.rows)
                        .slice(1) // Skip the header row
                        .map(row => {
                            const cell = row.cells[colIndex + 1]; // Adjust for the row label column
                            return cell ? parseFloat(cell.textContent) || NaN : NaN; // Parse cell values or use NaN for empty cells
                        });
                    console.log(`Extracted values from column ${argument}:`, values);
                    const result = calculate(values, operation);
                    console.log(`Result of ${operation} for column ${argument}:`, result);
                    cell.textContent = result !== null ? result : "N/A"; // Display "N/A" if no valid values
                    return;
                } else {
                    console.error(`Invalid argument in formula: ${argument}`);
                    cell.textContent = "Invalid argument in formula"; // Specific error for invalid argument
                    return;
                }
            }

            // Handle basic arithmetic operations like =A,1+B1
            console.log(`Processing arithmetic formula: ${formula}`);
            const updatedFormula = formula.replace(/([A-Z]+)(\d+)/g, (_, column, row) => {
                const columnIndex = column.toUpperCase().charCodeAt(0) - 65 + 1; // Convert column letter to index
                const rowIndex = parseInt(row, 10); // Get the row index
                const referencedCell = spreadsheet.rows[rowIndex]?.cells[columnIndex]; // Get the referenced cell
                if (!referencedCell) {
                    throw new Error(`Cell ${column}${row} does not exist`); // Throw an error for missing cell
                }
                const cellValue = parseFloat(referencedCell.textContent.trim()); // Extract and parse the cell value
                console.log(`Extracted value for ${column}${row}:`, cellValue); // Debugging log
                if (isNaN(cellValue)) {
                    throw new Error(`Cell ${column}${row} contains invalid data`); // Throw an error for invalid data
                }
                return cellValue; // Replace the cell reference with its numeric value
            });
            console.log(`Updated formula after replacing cell references: ${updatedFormula}`);
            const result = eval(updatedFormula);
            console.log(`Result of arithmetic formula: ${result}`);
            cell.textContent = result; // Update the cell with the calculated result
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
            } else if (/^(sum|min|max)\(/i.test(formula)) {
                // If the formula starts with SUM, MIN, or MAX, call evaluateFormula
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

    // Add a new row
    function addRow(rowIndex) {
        const newRow = document.createElement("tr");
        const rowNumber = rowIndex + 1;
        newRow.innerHTML = `<td class="row-label">${rowNumber}<div class="hover-buttons"><button class="addRow">+</button><button class="deleteRow">−</button></div></td>` +
            Array.from(headers.cells).slice(1).map(() => '<td contenteditable="true"></td>').join("");

        const rows = Array.from(spreadsheet.tBodies[0].rows);
        if (rowIndex < rows.length) {
            spreadsheet.tBodies[0].insertBefore(newRow, rows[rowIndex + 1]);
        } else {
            spreadsheet.tBodies[0].appendChild(newRow);
        }

        updateRowLabels();
        addResizingHandles(); // Reapply resizing handles to include the new row
    }

    // Add a new column
    function addColumn(columnIndex) {
        const adjustedColumnIndex = columnIndex + 1; // Add 1 to insert to the right of the selected column

        const newHeader = document.createElement("th");
        newHeader.className = "header-cell";
        newHeader.innerHTML = `<div class="hover-buttons"><button class="addColumn">+</button><button class="deleteColumn">−</button></div>`;

        // Insert the new header at the correct position
        headers.insertBefore(newHeader, headers.cells[adjustedColumnIndex]);

        // Add a new cell to each row at the correct position
        Array.from(spreadsheet.tBodies[0].rows).forEach(row => {
            const newCell = document.createElement("td");
            newCell.contentEditable = "true";
            row.insertBefore(newCell, row.cells[adjustedColumnIndex]);
        });

        updateColumnLabels();
        addResizingHandles(); // Reapply resizing handles to include the new column
    }

    // Function to delete a row and update row labels
    function deleteRow(targetRow) {
        if (spreadsheet.rows.length > 2) { // Ensure at least one data row remains
            targetRow.remove();
            updateRowLabels();
            addResizingHandles(); // Reapply resizing handles after removing a row
        }
    }

    // Function to delete a column and update column labels
    function deleteColumn(targetColumnIndex) {
        if (headers.cells.length > 2) { // Ensure at least one data column remains
            // Delete the header cell
            headers.deleteCell(targetColumnIndex);

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
                if (index > 0) {
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
            addRow(rowIndex);
            addResizingHandles(); // Add resizing handles to the new row
        }
        if (e.target.classList.contains("deleteRow")) {
            const targetRow = e.target.closest("tr");
            deleteRow(targetRow);
        }
        if (e.target.classList.contains("addColumn")) {
            const columnIndex = Array.from(headers.cells).indexOf(e.target.closest("th")) - 1; // Exclude the row label column
            addColumn(columnIndex);
            addResizingHandles(); // Add resizing handles to the new column
        }
        if (e.target.classList.contains("deleteColumn")) {
            const targetColumnIndex = Array.from(headers.cells).indexOf(e.target.closest("th"));
            deleteColumn(targetColumnIndex);
        }
    });

    // Function to get the spreadsheet title
    function getSpreadsheetTitle() {
        return spreadsheetTitle.value.trim() || "Untitled Spreadsheet";
    }

    // Function to export the spreadsheet to CSV
    function exportToCsv() {
        const title = document.getElementById("spreadsheetTitle").value.trim() || "Untitled Spreadsheet"; // Get the title or use a default
        const rows = Array.from(document.querySelectorAll("#spreadsheet tbody tr"));
        const csvContent = [
            title, // Add the title as the first line
            ...rows.map(row => {
                const cells = Array.from(row.cells).slice(1); // Skip the row label
                return cells.map(cell => cell.textContent.trim()).join(","); // Get cell values
            })
        ].join("\n");

        const blob = new Blob([csvContent], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;

        // Use the title as the filename, replacing invalid characters with underscores
        const sanitizedTitle = title.replace(/[^a-zA-Z0-9-_]/g, "_");
        a.download = `${sanitizedTitle}.csv`;

        a.click();
        URL.revokeObjectURL(url);
    }

    // Function to import a CSV file into the spreadsheet
    function importFromCsv(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function (e) {
            const csvContent = e.target.result.split("\n");
            const title = csvContent.shift(); // Extract the first line as the title
            document.getElementById("spreadsheetTitle").value = title; // Set the spreadsheet title

            const rows = csvContent.map(row => row.split(","));
            while (spreadsheet.rows.length > 1) {
                spreadsheet.deleteRow(1); // Clear existing rows (except headers)
            }

            rows.forEach((rowValues, rowIndex) => {
                const newRow = spreadsheet.insertRow();
                const rowLabel = newRow.insertCell(0);
                rowLabel.className = "row-label";
                rowLabel.innerHTML = `${rowIndex + 1}<div class="hover-buttons"><button class="addRow">+</button><button class="deleteRow">−</button></div>`;

                rowValues.forEach((value, colIndex) => {
                    const newCell = newRow.insertCell(colIndex + 1);
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