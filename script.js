class ExcelSheet {
    constructor() {
        this.columnCanvas = document.getElementById("column-canvas");
        this.columnCtx = this.columnCanvas.getContext("2d");
        this.colSizeArray = Array(28).fill(65);
        this.columnCtx.fillStyle = "#ebebeb";
        this.columnCtx.fillRect(0, 0, this.columnCanvas.width, this.columnCanvas.height);

        this.rowCanvas = document.getElementById("row-canvas");
        this.rowCanvas.setAttribute('tabindex', '0');
        this.rowCtx = this.rowCanvas.getContext("2d");
        this.rowSizeArray = Array(28).fill(20);
        this.rowCtx.fillStyle = "#ebebeb";
        this.rowCtx.fillRect(0, 0, this.rowCanvas.width, this.rowCanvas.height);

        this.cellCanvas = document.getElementById("cell-canvas");
        this.cellCtx = this.cellCanvas.getContext("2d");
        this.cellCanvas.style.position = "sticky";
        this.cellCanvas.style.top = 0;
        this.cellCanvas.style.left = 0;

        this.mouseDownOnColumnXPos = 0;
        this.mouseMoveOnColumnXPos = 0;
        this.mouseMoveOnColumnYPos = 0;
        this.isMouseDownOnColumn = false;

        this.mouseDownOnRowYPos = 0;
        this.mouseMoveOnRowYPos = 0;
        this.isMouseMoveOnRow = false;
        this.mouseMoveOnRowXPos = 0;
        this.isMouseDownOnRow = false;

        this.columnCanvas.addEventListener('mousedown', this.handleColumnMouseDown.bind(this));
        this.columnCanvas.addEventListener('mousemove', this.handleColumnMouseMove.bind(this));

        this.columnCanvas.addEventListener('mouseup', (event) => {
            this.isMouseDownOnColumn = false;
            this.columnCanvas.style.cursor = "default";
            this.columnCanvas.removeEventListener('mousedown', this.handleColumnMouseDown.bind(this));
            this.columnCanvas.removeEventListener('mousemove', this.handleColumnMouseMove.bind(this));
        });

        this.rowCanvas.addEventListener('mousedown', this.handleRowMouseDown.bind(this));
        this.rowCanvas.addEventListener('mousemove', this.handleRowMouseMove.bind(this));

        this.rowCanvas.addEventListener('mouseup', (event) => {
            this.isMouseDownOnRow = false;
            this.rowCanvas.style.cursor = "default";
            this.rowCanvas.removeEventListener('mousedown', this.handleRowMouseDown.bind(this));
            this.rowCanvas.removeEventListener('mousemove', this.handleRowMouseMove.bind(this));
        });

        this.inputBox = this.createElement("input", "text", "inputBox");
        this.inputBox.setAttribute("name", "cellInput");
        document.getElementById("cell-container").appendChild(this.inputBox);
        this.inputBox.style.position = "absolute";
        this.inputBox.style.display = "none";

        this.inputBox.addEventListener('keypress', this.addData.bind(this));

        this.cellPosInfo = {
            cellColumnNo: 0,
            cellRowNo: 0,
            cellPosX: 0,
            cellPosY: 0
        }

        this.isMouseDownOnCellCanvas = false;
        this.cellCanvas.addEventListener('mousedown', this.getInputBox.bind(this));

        this.currColNo = 0;
        this.currRowNo = 0;
        this.endRowNo = 0;
        this.endColNo = 0;
        document.getElementById("cell-container").addEventListener('mousemove', this.handleMouseMoveOnCell.bind(this));
        document.getElementById("cell-container").addEventListener('mouseup', (event) => {
            if (this.isMouseDownOnCellCanvas === true) {
                this.isMouseDownOnCellCanvas = false;
            }
        });

        this.selectedCellsArray = new Array();

        this.graphContainer = this.createElement("div", "", "graphContainer");
        document.getElementById("cell-container").appendChild(this.graphContainer);
        this.exitGraphIconContainer = this.createElement("div", "", "exitGraphIconContainer");
        this.graphContainer.appendChild(this.exitGraphIconContainer);
        this.graphCanvas = this.createElement("canvas", "", "graphCanvas");
        this.graphCanvas.id = "graphCanvas";
        this.graphContainer.appendChild(this.graphCanvas);
        this.exitGraphIcon = this.createElement("img", "", "exitGraphIcon");
        this.exitGraphIcon.setAttribute("src", "exitIcon.jpg");
        this.exitGraphIconContainer.appendChild(this.exitGraphIcon);
        this.graphContainer.style.display = "none";

        document.getElementById("bar-graph-btn").addEventListener('click', this.showBarGraph.bind(this));
        document.getElementById("line-graph-btn").addEventListener('click', this.showLineGraph.bind(this));

        this.exitGraphIconContainer.addEventListener('click', () => {
            this.graphContainer.style.display = "none";
        });

        this.cellsToReplace = [];
        this.replaceButton = document.getElementById("replace-btn");
        this.replaceInputBox = document.getElementById("replace-inputBox");

        this.cellsToSelect = [];
        this.findButton = document.getElementById("find-btn");
        this.findInputBox = document.getElementById("find-inputBox");

        this.clearHighlightBtn = document.getElementById("clear-highlight-btn");
        this.clearHighlightBtn.addEventListener('click', (event) => {
            if (this.findInputBox.value == '') {
                return;
            }
            this.findInputBox.value = '';
            this.replaceInputBox.value = '';
            this.clearCellCanvas();
            this.redrawCellCanvas();
            this.renderData();
        })

        document.getElementById("find-btn-exit").addEventListener('click', (event) => {
            document.getElementById("find-input-container").style.display = "none";
        })

        document.getElementById("replace-btn-exit").addEventListener('click', (event) => {
            document.getElementById("replace-input-container").style.display = "none";
        })

        this.findButton.addEventListener('click', this.findElement.bind(this));

        document.getElementById("replace-btn-exit").addEventListener('click', (event) => {
            document.getElementById("replace-input-container").style.display = "none";
        })

        this.replaceButton.addEventListener('click', this.replaceElement.bind(this));

        this.scrollY = 0;
        this.newRowStart = 0
        this.newRowStart1 = 0;

        this.scrollX = 0;
        this.newColStart = 0;
        this.newColStart1 = 0;

        document.getElementById("outer-cell-container").addEventListener('scroll', (event) => {
            const element = document.getElementById("outer-cell-container");
            this.scrollY = element.scrollTop;
            this.handleDivHeight(this.scrollY);
            this.handleYScroll();

            this.scrollX = element.scrollLeft;
            this.handleDivWidth(this.scrollX);
            if (this.scrollX > 0) {
                this.handleXScroll();
            }

            document.getElementById("demo").innerHTML = "X-Scroll " + element.scrollLeft +
                "<br>Y-Scroll " + element.scrollTop;
        })

        this.dataArray = [];
        for (let i = 0; i < 100001; i++) {
            this.dataArray[i] = [];
            for (let j = 0; j < 13; j++) {
                this.dataArray[i][j] = "";
            }
        }

        this.form = document.getElementById('csvForm');

        this.form.addEventListener('submit', (e) => {
            e.preventDefault();
            const fileInput = document.getElementById('csvFile');
            const formData = new FormData();
            formData.append('file', fileInput.files[0]);

            fetch('https://localhost:7018/api/UserRecords/upload', {
                method: 'POST',
                body: formData
            }).then(response => {
                if (response.ok) {
                    alert('File uploaded successfully');

                }
            });
        });

        this.fetchData(0, 27);

        this.rowCanvas.addEventListener('click', () => {
            this.rowCanvas.focus();
        });

        this.rowCanvas.addEventListener('click', this.highlightRow.bind(this));
        this.isMouseClickedOnRow = false;
        this.highlightedRowNo = 0;

        this.rowCanvas.addEventListener('keydown', this.deleteRecord.bind(this));
    };
    // Outside Constructor

    fetchData(startRow = 0, rowCount = this.rowSizeArray.length) {
        fetch(`https://localhost:7018/api/UserRecords?startRow=${startRow}&rowCount=${rowCount}`)
            .then(response => response.json())
            .then(data => {
                data.forEach((record, rowIndex) => {
                    this.dataArray[startRow + rowIndex] = Object.values(record);
                });
                this.clearCellCanvas()
                this.renderData(this.startRow, 0);
                this.redrawCellCanvas();
            });
    }

    // deleteRecord(event) {
    //     if (event.key === "Delete") {
    //         console.log("Delete pressed", this.highlightedRowNo);

    //         if (this.highlightedRowNo === null) return; // Ensure a row is selected

    //         const emailId = this.dataArray[this.highlightedRowNo][0]; // Fetch EmailID from the row

    //         fetch(`https://localhost:7018/api/UserRecords/${emailId}`, {
    //             method: 'DELETE'
    //         })
    //             .then(response => {
    //                 if (response.ok) {
    //                     this.dataArray.splice(this.highlightedRowNo, 1); 
    //                     this.clearCellCanvas(); 
    //                     this.redrawCellCanvas();
    //                     this.renderData();

    //                 }
    //             })
    //             .catch(error => console.error("Error deleting record:", error));
    //     }

    //     alert('Record deleted successfully!');
    // }

    deleteRecord(event) {
        if (event.key === "Delete") {
            console.log("Delete pressed", this.highlightedRowNo);

            if (this.highlightedRowNo === null) return; // Ensure a row is selected

            const emailId = this.dataArray[this.highlightedRowNo][0]; // Fetch EmailID from the row

            fetch(`https://localhost:7018/api/UserRecords/${emailId}`, {
                method: 'DELETE'
            })
                .then(response => {
                    if (response.ok) {
                        this.dataArray.splice(this.highlightedRowNo, 1);

                        this.clearCellCanvas();
                        this.redrawCellCanvas();
                        this.renderData();

                        setTimeout(() => {
                            alert('Record deleted successfully!');
                        }, 100);
                    }
                })
                .catch(error => console.error("Error deleting record:", error));
        }
    }


    highlightRow(event) {
        this.isMouseClickedOnRow = true;
        if (this.isMouseClickedOnRow == true) {
            let cumulativeRowHeight = 0, rowNo = 0;
            for (let i = 0; i < this.rowSizeArray.length; i++) {
                cumulativeRowHeight += this.rowSizeArray[i];
                if (cumulativeRowHeight > this.mouseDownOnRowYPos) {
                    rowNo = i;
                    this.highlightedRowNo = i;
                    break;
                }
            }

            this.clearCellCanvas();
            this.cellCtx.strokeStyle = "#107c41";
            this.cellCtx.lineWidth = 4;
            this.cellCtx.strokeRect(0, cumulativeRowHeight - this.rowSizeArray[rowNo],
                this.cellCanvas.width, this.rowSizeArray[rowNo]);
            this.cellCtx.fillStyle = "#ebebeb";
            this.cellCtx.fillRect(0, cumulativeRowHeight - this.rowSizeArray[rowNo],
                this.cellCanvas.width, this.rowSizeArray[rowNo]);
            this.redrawCellCanvas();
            this.renderData();
        }
    }

    updateServerRecord(cellRowNo, cellColumnNo) {
        const updatedRecord = this.dataArray[cellRowNo];
        const emailId = updatedRecord[0];

        const record = {
            EmailID: emailId,
            Name: updatedRecord[1],
            Country: updatedRecord[2],
            State: updatedRecord[3],
            City: updatedRecord[4],
            Telephone: updatedRecord[5],
            AddressLine1: updatedRecord[6],
            AddressLine2: updatedRecord[7],
            DateOfBirth: updatedRecord[8],
            SalaryFY2019_20: updatedRecord[9],
            SalaryFY2020_21: updatedRecord[10],
            SalaryFY2021_22: updatedRecord[11],
            SalaryFY2022_23: updatedRecord[12],
            SalaryFY2023_24: updatedRecord[13]
        };

        fetch(`https://localhost:7018/api/UserRecords/${emailId}`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(record)
        })
            .then(response => {
                if (response.ok) {
                    alert('Record updated successfully!');
                } else {
                    alert('Failed to update record.');
                }
            });
    }

    addData(event) {
        if (event.key === "Enter") {
            let inputText = this.inputBox.value;
            this.setCellPosInfo(event);
            let cellRowNo = this.cellPosInfo.cellRowNo;
            let cellColumnNo = this.cellPosInfo.cellColumnNo;
            this.dataArray[cellRowNo][cellColumnNo] = inputText;
            this.updateServerRecord(cellRowNo, cellColumnNo);
        }
    }

    renderData(startRow = 0, startCol = 0) {
        let cumulativeRowHeight = -1 * this.rowSizeArray[0], cumulativeColumnWidth = -1 * this.colSizeArray[0];

        for (let i = startRow; i < this.rowSizeArray.length && i < this.dataArray.length; i++) {
            cumulativeRowHeight += this.rowSizeArray[i];

            for (let j = startCol; j < this.colSizeArray.length && j < this.dataArray[i].length; j++) {
                cumulativeColumnWidth += this.colSizeArray[j];

                this.cellCtx.save();

                this.cellCtx.beginPath();
                this.cellCtx.rect(cumulativeColumnWidth - this.scrollX, cumulativeRowHeight - this.scrollY, this.colSizeArray[j], this.rowSizeArray[i]);
                this.cellCtx.clip();

                this.cellCtx.textBaseline = "top";
                this.cellCtx.textAlign = "left";
                this.cellCtx.font = "15px Arial";
                this.cellCtx.fillStyle = "#000000";

                if (this.dataArray[i] && this.dataArray[i][j] !== undefined) {
                    this.cellCtx.fillText(this.dataArray[i][j], cumulativeColumnWidth + 4 - this.scrollX, cumulativeRowHeight + 4 - this.scrollY);
                }

                this.cellCtx.restore();
            }
            cumulativeColumnWidth = -1 * this.colSizeArray[0];
        }
    }

    handleYScroll() {

        // Vertical scrolling Row Canvas
        this.clearRowCanvas();
        this.redrawRowCanvas(this.scrollY);
        let cumulativeRowHeight1 = 0;
        for (let i = 0; i < this.rowSizeArray.length; i++) {
            cumulativeRowHeight1 += this.rowSizeArray[i];
            if (cumulativeRowHeight1 > this.scrollY) {
                this.newRowStart1 = i;
                break;
            }
        }

        for (let i = 0; i < this.newRowStart1; i++) {
            this.rowSizeArray.push(20);
        }

        this.nameRows(this.newRowStart1);

        // Vertical scrolling Cell Canvas
        this.clearCellCanvas();
        this.redrawCellCanvas(this.scrollY, this.scrollX);

        let cumulativeRowHeight = 0;
        for (let i = 0; i < this.rowSizeArray.length; i++) {
            cumulativeRowHeight += this.rowSizeArray[i];
            if (cumulativeRowHeight > this.scrollY) {
                this.newRowStart = i;
                break;
            }
        }
        this.fetchData(this.newRowStart);
        this.renderData(this.newRowStart, this.newColStart);
    }

    handleDivWidth(scrollX) {
        let cellContainer = document.getElementById("cell-container");
        let currentWidth = window.getComputedStyle(cellContainer).width;
        let numericWidth = parseInt(currentWidth);
        if (scrollX > (numericWidth - 2000)) {
            cellContainer.style.width = numericWidth + 6000 + 'px';
        }
    }

    handleXScroll() {

        // Horizontal scrolling Column Canvas
        this.clearColCanvas();
        this.redrawColCanvas(this.scrollX);
        let cumulativeColWidtht1 = 0;
        for (let i = 0; i < this.colSizeArray.length; i++) {
            cumulativeColWidtht1 += this.colSizeArray[i];
            if (cumulativeColWidtht1 > this.scrollX) {
                this.newColStart1 = i;
                break;
            }
        }

        for (let i = 0; i < this.newColStart1; i++) {
            this.colSizeArray.push(65);
        }

        this.nameColumns(this.newColStart1);

        // Horizontal scrolling Cell Canvas
        this.clearCellCanvas();
        this.redrawCellCanvas(this.scrollY, this.scrollX);

        let cumulativeColumnWidth = 0;
        for (let i = 0; i < this.colSizeArray.length; i++) {
            cumulativeColumnWidth += this.colSizeArray[i];
            if (cumulativeColumnWidth > this.scrollX) {
                this.newColStart = i;
                break;
            }
        }
        this.renderData(this.newRowStart, this.newColStart);
    }

    handleDivHeight(scrollY) {
        let cellContainer = document.getElementById("cell-container");
        let currentHeight = window.getComputedStyle(cellContainer).height;
        let numericHeight = parseInt(currentHeight);
        if (scrollY > (numericHeight - 1000)) {
            cellContainer.style.height = numericHeight + 6000 + 'px';
        }
    }

    clearCellCanvas() {
        this.cellCtx.clearRect(0, 0, 1820, 560);
    }

    redrawCellCanvas(scrollY, scrollX) {
        this.drawGridColumnLines(0, 560, scrollX);
        this.drawGridRowLines(0, 1820, scrollY);
    }

    clearRowCanvas() {
        this.rowCtx.clearRect(0, 0, this.rowCanvas.width, this.rowCanvas.height);
    }

    clearColCanvas() {
        this.columnCtx.clearRect(0, 0, this.columnCanvas.width, this.columnCanvas.height);
    }

    redrawRowCanvas(scrollY) {
        this.rowCtx.fillStyle = "#ebebeb";
        this.rowCtx.fillRect(0, 0, this.rowCanvas.width, this.rowCanvas.height);
        this.drawRowLines(0, 43, scrollY);
        this.drawRowBoundary(43, 0, 43, 560);
    }

    redrawColCanvas(scrollX) {
        this.columnCtx.fillStyle = "#ebebeb";
        this.columnCtx.fillRect(0, 0, this.columnCanvas.width, this.columnCanvas.height);
        this.drawColumnLines(10, 30, scrollX);
        this.drawColumnBoundary(0, 30, 1820, 30);
    }

    replaceElement(event) {
        document.getElementById("replace-input-container").style.display = "block";
        if (!this.replaceInputBox.value) {
            return;
        }

        let findInputValue = this.findInputBox.value;
        let replaceInputValue = this.replaceInputBox.value;
        for (let i = 0; i < this.dataArray.length; i++) {
            for (let j = 0; j < this.dataArray[i].length; j++) {
                if (this.dataArray[i][j] == findInputValue) {
                    this.dataArray[i][j] = replaceInputValue;
                    let cellNo = [i, j];
                    this.cellsToReplace.push(cellNo);
                }
            }
        }

        this.clearCellCanvas();
        this.redrawCellCanvas();
        this.highlightElement(this.cellsToReplace);
        this.renderData();
    }

    highlightElement(cellsToSelectArray) {

        this.clearCellCanvas();
        this.redrawCellCanvas();

        for (let cellNo = 0; cellNo < cellsToSelectArray.length; cellNo++) {

            let rowNo = cellsToSelectArray[cellNo][0];
            let colNo = cellsToSelectArray[cellNo][1];

            let cumulativeColumnWidth = 0, cumulativeRowHeight = 0;
            for (let i = 0; i < rowNo; i++) {
                cumulativeRowHeight += this.rowSizeArray[i];
            }

            for (let j = 0; j < colNo; j++) {
                cumulativeColumnWidth += this.colSizeArray[j];
            }

            // console.log(cumulativeRowHeight, cumulativeColumnWidth);

            this.cellCtx.fillStyle = "#c1e1d0";
            this.cellCtx.fillRect(cumulativeColumnWidth + 2, cumulativeRowHeight + 2,
                (this.colSizeArray[colNo]) - 4, (this.rowSizeArray[rowNo]) - 4
            )
        }

        this.renderData();
    }

    findElement(event) {
        document.getElementById("find-input-container").style.display = "block";
        if (!this.findInputBox.value) {
            return;
        }
        let flag = false;

        let findInputValue = this.findInputBox.value;
        for (let i = 0; i < this.dataArray.length; i++) {
            for (let j = 0; j < this.dataArray[i].length; j++) {
                if (this.dataArray[i][j] == findInputValue) {
                    flag = true;
                    // console.log(this.dataArray[i][j], findInputValue)
                    let cellIdx = [i, j];
                    this.cellsToSelect.push(cellIdx);
                }
            }
        }

        if (flag == true) {
            this.highlightElement(this.cellsToSelect);
        }
    }

    selectCellsBorder(currRowNo, currColumnNo, mouseMoveCellRow, mouseMoveCellCol) {
        let drawX = 0, drawY = 0, startX = 0, startY = 0, totalX = 0, totalY = 0;

        let startColNo = Math.min(currColumnNo, mouseMoveCellCol);
        let startRowNo = Math.min(currRowNo, mouseMoveCellRow);
        let endColNo = Math.max(currColumnNo, mouseMoveCellCol);
        let endRowNo = Math.max(currRowNo, mouseMoveCellRow);

        for (let i = 0; i <= endColNo; i++) {
            totalX += this.colSizeArray[i];
        }

        for (let i = 0; i < startColNo; i++) {
            startX += this.colSizeArray[i];
        }

        for (let i = 0; i <= endRowNo; i++) {
            totalY += this.rowSizeArray[i + 1];
        }

        for (let i = 0; i < startRowNo; i++) {
            startY += this.rowSizeArray[i];
        }

        drawX = totalX - startX;
        drawY = totalY - startY;
        let currCellXPos = this.cellPosInfo.cellPosX;
        let currCellYPos = this.cellPosInfo.cellPosY;

        this.storeSelectedCells(currRowNo, currColumnNo, mouseMoveCellRow, mouseMoveCellCol);

        this.clearColumnAndCellCanvas();
        this.clearRowAndCellCanvas();

        this.cellCtx.fillStyle = "#e7f1ec"
        this.cellCtx.fillRect(startX, startY, drawX, drawY)

        this.cellCtx.fillStyle = "white"
        this.cellCtx.fillRect(currCellXPos, currCellYPos,
            this.colSizeArray[currColumnNo], this.rowSizeArray[currRowNo]);

        this.redrawColumnAndCellCanvas();
        this.redrawRowAndCellCanvas();
        this.renderData();

        this.cellCtx.strokeStyle = "#107c41";
        this.cellCtx.lineWidth = 2;
        this.cellCtx.strokeRect(startX, startY, drawX, drawY);

        this.performCalculations();
        // this.drawCellBorder("#bcbcbc")
    }

    getInputBox(event) {
        this.isMouseDownOnCellCanvas = true;
        let inputBoxValue = this.inputBox.value;

        this.dataArray[this.cellPosInfo.cellRowNo][this.cellPosInfo.cellColumnNo] = this.inputBox.value;

        this.setCellPosInfo(event);
        let cellRowNo = this.cellPosInfo.cellRowNo;
        let cellColNo = this.cellPosInfo.cellColumnNo;
        let cellPosX = this.cellPosInfo.cellPosX;
        let cellPosY = this.cellPosInfo.cellPosY;

        this.inputBox.style.display = "flex";
        this.inputBox.style.fontSize = "15px";
        this.inputBox.style.boxSizing = "border-box";
        this.inputBox.style.left = cellPosX + "px";
        this.inputBox.style.top = cellPosY + "px";
        this.inputBox.style.width = this.colSizeArray[cellColNo] + "px";
        this.inputBox.style.height = this.rowSizeArray[cellRowNo] + "px";
        // this.inputBox.style.border = "none";
        this.inputBox.style.background = "none";
        this.inputBox.style.outline = "none";
        this.inputBox.style.focus;

        this.cellCtx.textBaseline = "top";
        this.cellCtx.textAlign = "left";
        this.cellCtx.font = "14px Arial";
        this.cellCtx.fillStyle = "#000000";

        let currInputBoxValue = this.dataArray[cellRowNo][cellColNo];
        this.dataArray[cellRowNo][cellColNo] = "";
        // this.clearColumnAndCellCanvas();
        // this.clearRowAndCellCanvas();
        this.clearCellCanvas();
        this.inputBox.value = currInputBoxValue;
        // this.redrawColumnAndCellCanvas();
        // this.redrawRowAndCellCanvas();
        this.redrawCellCanvas();
        this.renderData();
        this.drawCellBorder()
        this.dataArray[cellRowNo][cellColNo] = currInputBoxValue;
    }

    drawCellBorder() {
        let cellRowNo = this.cellPosInfo.cellRowNo;
        let cellColNo = this.cellPosInfo.cellColumnNo;
        let cellPosX = this.cellPosInfo.cellPosX;
        let cellPosY = this.cellPosInfo.cellPosY;
        this.cellCtx.strokeStyle = "#107c41";
        this.cellCtx.lineWidth = 2;
        this.cellCtx.strokeRect(cellPosX, cellPosY, this.colSizeArray[cellColNo], this.rowSizeArray[cellRowNo]);
    }

    setCellPosInfo(event) {
        let clickCoordinates = this.getClickCoordinates(event);
        let cumulativeColumnWidth = 0;
        let cumulativeRowHeight = 0;


        for (let i = 0; i < this.colSizeArray.length; i++) {
            cumulativeColumnWidth += this.colSizeArray[i];
            if (clickCoordinates.clickPosX < cumulativeColumnWidth) {
                this.cellPosInfo.cellColumnNo = i;
                this.cellPosInfo.cellPosX = cumulativeColumnWidth - this.colSizeArray[i];
                this.startX = cumulativeColumnWidth - this.colSizeArray[i];
                this.currColNo = i;
                break;
            }
        }

        for (let i = 0; i < this.rowSizeArray.length; i++) {
            cumulativeRowHeight += this.rowSizeArray[i];
            if (clickCoordinates.clickPosY < cumulativeRowHeight) {
                this.cellPosInfo.cellRowNo = i;
                this.cellPosInfo.cellPosY = cumulativeRowHeight - this.rowSizeArray[i];
                this.startY = cumulativeRowHeight - this.rowSizeArray[i];
                this.currRowNo = i;
                break;
            }
        }
    }

    getClickCoordinates(event) {
        let rect = event.target.getBoundingClientRect();
        let clickPosX = event.clientX - rect.left;
        let clickPosY = event.clientY - rect.top;

        return { clickPosX: clickPosX, clickPosY: clickPosY };
    }

    storeSelectedCells(currRowNo, currColumnNo, mouseMoveCellRow, mouseMoveCellCol) {

        let startRow = Math.min(currRowNo, mouseMoveCellRow);
        let endRow = Math.max(currRowNo, mouseMoveCellRow);
        let startCol = Math.min(currColumnNo, mouseMoveCellCol);
        let endCol = Math.max(currColumnNo, mouseMoveCellCol);
        this.selectedCellsArray = [];

        for (let i = startRow; i <= endRow; i++) {
            for (let j = startCol; j <= endCol; j++) {
                let cellPos = new Object();
                cellPos.rowNo = i;
                cellPos.colNo = j;
                this.selectedCellsArray.push(cellPos);
            }
        }

        this.endRowNo = mouseMoveCellRow;
        this.endColNo = mouseMoveCellCol;
    }

    /**
     * 
     * @param {string} tag 
     * @param {string} type 
     * @param {string} className 
     * @returns{HTMLElement}
     */
    createElement(tag, type = "", className) {
        const element = document.createElement(tag);
        element.className = className;
        if (type != "") {
            element.type = type;
        }
        return element;
    }

    handleRowMouseMove = (event) => {

        let cursorPosInfo = this.getClickCoordinates(event);
        let clickPosX = cursorPosInfo.clickPosX;
        let clickPosY = cursorPosInfo.clickPosY;
        let cumulativeRowHeight = 0;

        for (let i = 0; i < this.rowSizeArray.length; i++) {
            cumulativeRowHeight += this.rowSizeArray[i];

            if (clickPosY < (cumulativeRowHeight + 5) &&
                clickPosY > (cumulativeRowHeight - 5)) {
                this.rowCanvas.style.cursor = "row-resize";
                break;
            }

            else {
                this.rowCanvas.style.cursor = "default";
            }
        }

        if (this.isMouseDownOnRow == true) {
            const rect = event.target.getBoundingClientRect();
            this.mouseMoveOnRowYPos = event.clientY - rect.top;
            this.mouseMoveOnRowXPos = event.clientX - rect.left;
            if (this.mouseMoveOnRowXPos > this.rowCanvas.width - 4) {
                this.isMouseDownOnRow = false;
            }
            this.resizeRowHeight();
        }
    };

    handleRowMouseDown = (event) => {
        this.isMouseDownOnRow = true;
        this.rowCanvas.style.cursor = "row-resize";
        const rect = event.target.getBoundingClientRect();
        this.mouseDownOnRowYPos = event.clientY - rect.top;
    }

    handleMouseMoveOnCell(event) {
        if (this.isMouseDownOnCellCanvas == true) {
            let rect = event.target.getBoundingClientRect();
            let mouseMovePosX = event.clientX - rect.left;
            let mouseMovePosY = event.clientY - rect.top;
            let cumulativeSelectedWidth = 0;
            let cumulativeSelectedHeight = 0;
            let mouseMoveCellPosInfo = this.getMouseMoveCellPosInfo(event);
            let mouseMoveCellCol = mouseMoveCellPosInfo.cellColumnNo;
            let mouseMoveCellRow = mouseMoveCellPosInfo.cellRowNo;

            this.selectCellsBorder(this.currRowNo, this.currColNo, mouseMoveCellRow, mouseMoveCellCol);
        }
    }

    resizeRowHeight() {
        this.isMouseMoveOnRow = true;
        let cumulativeRowHeight = 0;
        let rowIndex = -1;

        for (let i = 0; i < this.rowSizeArray.length; i++) {
            cumulativeRowHeight += this.rowSizeArray[i];
            if (this.mouseDownOnRowYPos < cumulativeRowHeight) {
                rowIndex = i;
                break;
            }
            if (this.mouseDownOnRowYPos < (cumulativeRowHeight + 5) &&
                this.mouseDownOnRowYPos > (cumulativeRowHeight - 5)) {
                rowIndex = i;
                break;
            }
        }

        let extendedRowHeight = this.mouseMoveOnRowYPos - this.mouseDownOnRowYPos;

        this.rowSizeArray[rowIndex] += extendedRowHeight;

        if (this.rowSizeArray[rowIndex] < 10) {
            this.isMouseDownOnRow = false;
        }

        this.mouseDownOnRowYPos = this.mouseMoveOnRowYPos;

        this.clearRowAndCellCanvas();
        this.redrawRowAndCellCanvas();

        this.drawCellBorder();
        this.highlightElement(this.cellsToSelect);
        this.renderData();
    }

    clearRowAndCellCanvas() {
        this.cellCtx.clearRect(0, 0, this.cellCanvas.width, this.cellCanvas.height);
        this.rowCtx.clearRect(0, 0, this.rowCanvas.width, this.rowCanvas.height);
    }

    redrawRowAndCellCanvas() {
        this.rowCtx.fillStyle = "#ebebeb";
        this.rowCtx.fillRect(0, 0, this.rowCanvas.width, this.rowCanvas.height);
        this.drawRowLines(0, 43);
        this.drawGridColumnLines(0, 560);
        this.drawGridRowLines(0, 1820);
        this.drawRowBoundary(43, 0, 43, 560);
        this.nameRows();
    }

    getMouseMoveCellPosInfo(event) {
        let clickCoordinates = this.getMouseMoveCoordinates(event);
        let cumulativeColumnWidth = 0;
        let cumulativeRowHeight = 0;
        let cellColumnNo = 0, cellRowNo = 0;

        for (let i = 0; i < this.colSizeArray.length; i++) {
            cumulativeColumnWidth += this.colSizeArray[i];
            if (clickCoordinates.clickPosX < cumulativeColumnWidth) {
                cellColumnNo = i;
                break;
            }
        }

        for (let i = 0; i < this.rowSizeArray.length; i++) {
            cumulativeRowHeight += this.rowSizeArray[i];
            if (clickCoordinates.clickPosY < cumulativeRowHeight) {
                cellRowNo = i;
                break;
            }
        }

        return { cellColumnNo: cellColumnNo, cellRowNo: cellRowNo };
    }

    getMouseMoveCoordinates(event) {
        let rect = document.getElementById("cell-container").getBoundingClientRect();
        let clickPosX = event.clientX - rect.left;
        let clickPosY = event.clientY - rect.top;

        return { clickPosX: clickPosX, clickPosY: clickPosY };
    }

    handleColumnMouseDown = (event) => {
        this.isMouseDownOnColumn = true;
        this.columnCanvas.style.cursor = "col-resize";
        const rect = event.target.getBoundingClientRect();
        this.mouseDownOnColumnXPos = event.clientX - rect.left;
    };

    handleColumnMouseMove = (event) => {

        let cursorPosInfo = this.getClickCoordinates(event);
        let clickPosX = cursorPosInfo.clickPosX;
        let clickPosY = cursorPosInfo.clickPosY;
        let cumulativeColumnWidth = 0;

        for (let i = 0; i < this.colSizeArray.length; i++) {
            cumulativeColumnWidth += this.colSizeArray[i];

            if (clickPosX < (cumulativeColumnWidth + 5) &&
                clickPosX > (cumulativeColumnWidth - 5)) {
                this.columnCanvas.style.cursor = "col-resize";
                break;
            }

            else {
                this.columnCanvas.style.cursor = "default";
            }
        }

        if (this.isMouseDownOnColumn == true) {
            const rect = event.target.getBoundingClientRect();
            this.mouseMoveOnColumnXPos = event.clientX - rect.left;
            this.mouseMoveOnColumnYPos = event.clientY - rect.top;
            if (this.mouseMoveOnColumnYPos > this.columnCanvas.height - 2) {
                this.isMouseDownOnColumn = false;
            }
            this.resizeColumnWidth();
        }
    };

    destroyGraph() {
        if (this.draw) {
            this.draw.destroy()
        }
    }

    showLineGraph(event) {
        this.destroyGraph();

        this.graphContainer.style.display = "block";
        let xValues = [];
        let yValues = [];

        // Populate xValues based on selectedCellsArray
        for (let i = 0; i < this.selectedCellsArray.length; i++) {
            if (!xValues.includes(this.selectedCellsArray[i].rowNo + 1)) {
                xValues.push(this.selectedCellsArray[i].rowNo + 1);
            }
        }

        for (let i = this.currColNo; i <= this.endColNo; i++) {
            let line = i + 1
            let myObj = {
                label: "Line" + line,
                borderColor: "red",
                data: [],
                fill: false,
                tension: false
            }
            for (let j = this.currRowNo; j <= this.endRowNo; j++) {
                myObj.data.push(this.dataArray[j][i]);
            }

            yValues.push(myObj);
        }

        // Create the line chart
        this.draw = new Chart("graphCanvas", {
            type: "line",
            data: {
                labels: xValues,
                datasets: yValues  // Assign yValues to the datasets field
            },
            options: {
                responsive: false
            }
        });
    }

    showBarGraph(event) {
        this.destroyGraph()

        this.graphContainer.style.display = "block";
        let xValues = [];
        let yValues = [];

        for (let i = 0; i < this.selectedCellsArray.length; i++) {
            if (!xValues.includes(this.selectedCellsArray[i].rowNo + 1)) {
                xValues.push(this.selectedCellsArray[i].rowNo + 1);
            }
        }

        for (let i = this.currColNo; i <= this.endColNo; i++) {
            let myObj = {
                label: i + 1,
                backgroundColor: "red",
                data: []
            }
            for (let j = this.currRowNo; j <= this.endRowNo; j++) {
                myObj.data.push(this.dataArray[j][i]);
            }

            yValues.push(myObj);
        }

        this.draw = new Chart("graphCanvas", {
            type: "bar",
            data: {
                labels: xValues,
                datasets: yValues
            },

            options: {
                responsive: false
            }
        });
        // console.log(yValues);
    }

    performCalculations() {
        document.getElementById("sum").innerHTML = "";
        let sum = 0;
        for (let i = 0; i < this.selectedCellsArray.length; i++) {
            let rowNo = this.selectedCellsArray[i].rowNo;
            let colNo = this.selectedCellsArray[i].colNo;
            let cellElement = Number(this.dataArray[rowNo][colNo]);
            sum = sum + cellElement;
        }
        document.getElementById("sum").innerHTML = sum;

        document.getElementById("average").innerHTML = "";
        let average = 0;
        let sumAvg = 0;
        for (let i = 0; i < this.selectedCellsArray.length; i++) {
            let rowNo = this.selectedCellsArray[i].rowNo;
            let colNo = this.selectedCellsArray[i].colNo;
            let cellElement = Number(this.dataArray[rowNo][colNo]);
            sumAvg = sumAvg + cellElement;
        }
        document.getElementById("average").innerHTML = sumAvg / this.selectedCellsArray.length;

        document.getElementById("minValue").innerHTML = "";
        let minValue = Number.MAX_SAFE_INTEGER;
        for (let i = 0; i < this.selectedCellsArray.length; i++) {
            let rowNo = this.selectedCellsArray[i].rowNo;
            let colNo = this.selectedCellsArray[i].colNo;
            let cellElement = Number(this.dataArray[rowNo][colNo]);
            minValue = Math.min(minValue, cellElement);
        }
        document.getElementById("minValue").innerHTML = minValue;

        document.getElementById("maxValue").innerHTML = "";
        let maxValue = Number.MIN_SAFE_INTEGER;
        for (let i = 0; i < this.selectedCellsArray.length; i++) {
            let rowNo = this.selectedCellsArray[i].rowNo;
            let colNo = this.selectedCellsArray[i].colNo;
            let cellElement = Number(this.dataArray[rowNo][colNo]);
            maxValue = Math.max(maxValue, cellElement);
        }
        document.getElementById("maxValue").innerHTML = maxValue;
    }

    clearColumnAndCellCanvas() {
        this.cellCtx.clearRect(0, 0, this.cellCanvas.width, this.cellCanvas.height);
        this.columnCtx.clearRect(0, 0, this.columnCanvas.width, this.columnCanvas.height);
    }

    redrawColumnAndCellCanvas() {
        this.columnCtx.fillStyle = "#ebebeb";
        this.columnCtx.fillRect(0, 0, this.columnCanvas.width, this.columnCanvas.height);
        this.drawColumnLines(10, 30);
        this.drawGridColumnLines(0, 560);
        this.nameColumns();
        this.drawGridRowLines(0, 1820);
        this.drawColumnBoundary(0, 30, 1820, 30);
    }

    resizeColumnWidth() {
        let cumulativeColumnWidth = 0;
        let columnIndex = -1;

        for (let i = 0; i < this.colSizeArray.length; i++) {
            cumulativeColumnWidth += this.colSizeArray[i];

            if (this.mouseDownOnColumnXPos < (cumulativeColumnWidth + 5) &&
                this.mouseDownOnColumnXPos > (cumulativeColumnWidth - 5)) {
                columnIndex = i;
                break;
            }
        }

        let extendedColumnWidth = this.mouseMoveOnColumnXPos - this.mouseDownOnColumnXPos;

        this.colSizeArray[columnIndex] += extendedColumnWidth;

        if (this.colSizeArray[columnIndex] < 10) {
            this.isMouseDownOnColumn = false;
        }

        this.mouseDownOnColumnXPos = this.mouseMoveOnColumnXPos;

        this.clearColumnAndCellCanvas();
        this.redrawColumnAndCellCanvas();
        this.drawCellBorder();
        this.highlightElement(this.cellsToSelect);
        this.renderData();
        this.inputBox.style.width = this.colSizeArray[this.cellPosInfo.cellColumnNo] + "px";
    }

    drawColumnBoundary(moveToX, moveToY, lineToX, lineToY) {
        this.columnCtx.beginPath();
        this.columnCtx.moveTo(moveToX, moveToY);
        this.columnCtx.lineTo(lineToX, lineToY);
        this.columnCtx.strokeStyle = "#bcbcbc";
        this.cellCtx.lineWidth = 1;
        this.columnCtx.stroke();
    }

    drawColumnLines(moveToY, lineToY, scrollX = 0) {
        let accumulatedColumnWidth = 0;
        for (let i = 0; i < this.colSizeArray.length; i++) {
            accumulatedColumnWidth += this.colSizeArray[i];
            this.columnCtx.beginPath();
            this.columnCtx.moveTo(accumulatedColumnWidth - scrollX, moveToY);
            this.columnCtx.lineTo(accumulatedColumnWidth - scrollX, lineToY);
            this.columnCtx.strokeStyle = "#bcbcbc";
            this.cellCtx.lineWidth = 1;
            this.columnCtx.stroke();
        }
    }

    nameColumns(newColStart = 0) {
        // console.log(newColStart);
        this.columnCtx.font = "14px Arial";
        this.columnCtx.textAlign = "center";
        this.columnCtx.textBaseline = "middle";
        this.columnCtx.fillStyle = "#616161";

        let cumulativeColumnWidth = 0;
        let prevCumulativeColumnWidth = 0;

        for (let j = newColStart; j < this.colSizeArray.length; j++) {
            cumulativeColumnWidth += this.colSizeArray[j];
            prevCumulativeColumnWidth = cumulativeColumnWidth - this.colSizeArray[j];

            let columnName = this.getColumnName(j + newColStart);

            let left = prevCumulativeColumnWidth + ((cumulativeColumnWidth - prevCumulativeColumnWidth) / 2);

            this.columnCtx.fillText(columnName, left - this.scrollX, 21);
        }
    }

    getColumnName(index) {
        let columnName = '';
        while (index >= 0) {
            columnName = String.fromCharCode((index % 26) + 65) + columnName;
            index = Math.floor(index / 26) - 1;
        }
        return columnName;
    }

    drawRowBoundary(moveToX, moveToY, lineToX, lineToY) {
        this.rowCtx.beginPath();
        this.rowCtx.moveTo(moveToX, moveToY);
        this.rowCtx.lineTo(lineToX, lineToY);
        this.rowCtx.strokeStyle = "#bcbcbc";
        this.cellCtx.lineWidth = 1;
        this.rowCtx.stroke();
    }

    drawRowLines(moveToX, lineToX, scrollY = 0) {
        let accumulatedRowHeight = 0;
        for (let i = 0; i < this.rowSizeArray.length; i++) {
            accumulatedRowHeight += this.rowSizeArray[i];
            this.rowCtx.beginPath();
            this.rowCtx.moveTo(moveToX, accumulatedRowHeight - scrollY);
            this.rowCtx.lineTo(lineToX, accumulatedRowHeight - scrollY);
            this.rowCtx.strokeStyle = "#bcbcbc";
            this.cellCtx.lineWidth = 1;
            this.rowCtx.stroke();
        }
    }

    nameRows(newRowStart = 0) {
        this.rowCtx.font = "14px Arial";
        this.rowCtx.textAlign = "center";
        this.rowCtx.textBaseline = "middle";
        let rowIdx = newRowStart, prevCumulativeRowHeight = 0, cumulativeRowHeight = 0;

        while (rowIdx < this.rowSizeArray.length) {
            cumulativeRowHeight += this.rowSizeArray[rowIdx];
            prevCumulativeRowHeight = cumulativeRowHeight - this.rowSizeArray[rowIdx];
            this.rowCtx.fillStyle = "#616161";
            let top = prevCumulativeRowHeight + ((cumulativeRowHeight - prevCumulativeRowHeight) / 2)
            this.rowCtx.fillText(rowIdx + 1, 25, top - this.scrollY);
            rowIdx++;
        }
    }

    drawGridColumnLines(moveToY, lineToY, scrollX = 0) {
        let accumulatedColumnWidth = 0;
        for (let i = 0; i < this.colSizeArray.length; i++) {
            accumulatedColumnWidth += this.colSizeArray[i];
            this.cellCtx.beginPath();
            this.cellCtx.moveTo(accumulatedColumnWidth - scrollX, moveToY);
            this.cellCtx.lineTo(accumulatedColumnWidth - scrollX, lineToY);
            this.cellCtx.lineWidth = 1;
            this.cellCtx.strokeStyle = "#bcbcbc";
            this.cellCtx.stroke();
        }
    }

    drawGridRowLines(moveToX, lineToX, scrollY = 0) {
        let accumulatedRowHeight = 0;
        for (let i = 0; i < this.rowSizeArray.length; i++) {
            accumulatedRowHeight += this.rowSizeArray[i];
            this.cellCtx.beginPath();
            this.cellCtx.moveTo(moveToX, accumulatedRowHeight - scrollY);
            this.cellCtx.lineTo(lineToX, accumulatedRowHeight - scrollY);
            this.cellCtx.strokeStyle = "#bcbcbc";
            this.cellCtx.lineWidth = 1;
            this.cellCtx.stroke();
        }
    }
}

const myExcelSheet = new ExcelSheet();

// Drawing Column Canvas
myExcelSheet.drawColumnBoundary(0, 30, 1820, 30);
myExcelSheet.drawColumnLines(10, 30);
myExcelSheet.nameColumns();

// Drawing Row Canvas
myExcelSheet.drawRowBoundary(43, 0, 43, 560);
myExcelSheet.drawRowLines(0, 43);
myExcelSheet.nameRows();

// Drawing Cells Canvas
myExcelSheet.drawGridColumnLines(0, 560);
myExcelSheet.drawGridRowLines(0, 1820);

// let nums = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50];

// for(let i=0;i<100;i++){
//     for(let j=0;j<100;j++){
//         let idx = Math.floor(Math.random()*nums.length);
//         myExcelSheet.dataArray[i][j] = nums[idx];
//     }
// }

myExcelSheet.renderData();