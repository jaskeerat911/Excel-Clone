const PS = new PerfectScrollbar("#cells", {
    wheelSpeed: 2,
    wheelPropagation: true,
});

function findRowCol(ele) {
    let idArray = $(ele).attr("id").split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [rowId, colId];
}

for (let i = 1; i <= 100; i++) {
    let str = "";
    let n = i;

    while (n > 0) {
        let rem = n % 26;
        if (rem == 0) {
            str = 'Z' + str;
            n = Math.floor((n / 26)) - 1;
        } else {
            str = String.fromCharCode((rem - 1) + 65) + str;
            n = Math.floor((n / 26));
        }
    }
    $("#columns").append(`<div class="column-name">${str}</div>`);
    $("#rows").append(`<div class="row-name">${i}</div>`);
}

//perfect scrollbar
$("#cells").scroll(function () {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
})

let cellData = { "Sheet1": {} };
let selectedSheet = "Sheet1";
let totalSheets = 1;
let lastelyAddedSheetNumber = 1;
let mousemoved = false;
let startCellStored = false;
let startCell;
let endCell;
let defaultProperties = {
    "font-family": "Noto Sans",
    "font-size": 14,
    "text": "",
    "bold": false,
    "italic": false,
    "underlined": false,
    "alignment": "left",
    "bgcolor": "#fff",
    "color": "#444",
    "border": "none"
}

function loadNewSheet() {
    $("#cells").text("");
    for (let i = 1; i <= 100; i++) {
        let row = $(`<div class="cell-row"></div>`);
        for (let j = 1; j <= 100; j++){
            row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`)
        }
        $("#cells").append(row);
    }
    addEventsToCells();
    addSheetTabEventListeners();
}

loadNewSheet();

//Add Events to cells
function addEventsToCells() {
    $(".input-cell").dblclick(function () {
        $(this).attr("contenteditable", "true");
        $(this).focus();
    })
    
    $(".input-cell").blur(function () {
        $(this).attr("contenteditable", "false");
        let [rowId, colId] = findRowCol(this);
        updateCellData("text", $(this).text());
        console.log(cellData);
    })
    
    $(".input-cell").click(function (e) {
        let [rowId, colId] = findRowCol(this);
        let [topCell, bottomCell, leftCell, rightCell] = getTopBottomLeftRight(rowId, colId);
    
        if ($(this).hasClass("selected") && e.ctrlKey) {
            unselectCell(this, e, topCell, bottomCell, leftCell, rightCell);
        }
        else {
            selectCell(this, e, topCell, bottomCell, leftCell, rightCell);
        }
    })
    
    $(".input-cell").mousemove(function (e) {
        e.preventDefault();
        if (e.buttons == 1 && !e.ctrlKey) {
            $(".input-cell.selected").removeClass("selected top-selected  bottom-selected left-selected right-selected");
            mousemoved = true;
            if (!startCellStored) {
                let [rowId, colId] = findRowCol(e.target);
                startCell = { rowId: rowId, colId: colId };
                startCellStored = true;
            }
            else {
                let [rowId, colId] = findRowCol(e.target);
                endCell = { rowId: rowId, colId: colId };
                selectAllBetweenRange(startCell, endCell);
            }
        }
        else if (e.buttons == 0 && mousemoved) {
            startCellStored = false;
            mousemoved = false;
        }
    })
}
//

function getTopBottomLeftRight(rowId, colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);

    return [topCell, bottomCell, leftCell, rightCell];
}

function unselectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {
    if (e.ctrlKey && $(ele).attr("contenteditable") == "false") {
        
        if ($(ele).hasClass("top-selected")) {
            topCell.removeClass("bottom-selected");
        }
        
        if ($(ele).hasClass("bottom-selected")) {
            bottomCell.removeClass("top-selected");
        }
        
        if ($(ele).hasClass("left-selected")) {
            leftCell.removeClass("right-selected");
        }
        
        if ($(ele).hasClass("right-selected")) {
            rightCell.removeClass("left-selected");
        }
        $(ele).removeClass("selected top-selected bottom-selected left-selected right-selected");
    }
}

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell, mouseSelection) {
    if (e.ctrlKey || mouseSelection) {
        
        // top selected or not
        let topSelected;
        if (topCell) {
            topSelected = topCell.hasClass("selected");
        }
        // bottom selected or not
        let bottomSelected;
        if (bottomCell) {
            bottomSelected = bottomCell.hasClass("selected");
        }

        // left selected or not
        let leftSelected;
        if (leftCell) {
            leftSelected = leftCell.hasClass("selected");
        }
        // right selected or not
        let rightSelected;
        if (rightCell) {
            rightSelected = rightCell.hasClass("selected");
        }

        if (topSelected) {
            topCell.addClass("bottom-selected");
            $(ele).addClass("top-selected");
        }

        if (leftSelected) {
            leftCell.addClass("right-selected");
            $(ele).addClass("left-selected");
        }

        if (rightSelected) {
            rightCell.addClass("left-selected");
            $(ele).addClass("right-selected");
        }

        if (bottomSelected) {
            bottomCell.addClass("top-selected");
            $(ele).addClass("bottom-selected");
        }
    }
    else {
        $(".input-cell").removeClass("selected top-selected bottom-selected left-selected right-selected");
    }
    $(ele).addClass("selected");
    changeHeader(findRowCol(ele));
}

function changeHeader([rowId, colId]) {
    let data;
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        data = cellData[selectedSheet][rowId - 1][colId - 1];
    }
    else {
        data = defaultProperties;
    }
    $("#font-family").val(data["font-family"]);
    $("#font-family").css("font-family", data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");
    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, "underlined");
    $("#fill-color-icon").css("border-bottom", `4px solid ${data.bgcolor}`);
    $("#text-color-icon").css("border-bottom", `4px solid ${data.color}`);
    $("#border").val(data["border"]);
}

function addRemoveSelectFromFontStyle(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    }
    else {
        $(`#${property}`).removeClass("selected");
    }
}

function selectAllBetweenRange(start, end) {
    for (let i = (start.rowId < end.rowId ? start.rowId : end.rowId); i <= (start.rowId < end.rowId ? end.rowId : start.rowId); i++){
        for (let j = (start.colId < end.colId ? start.colId : end.colId); j <= (start.colId < end.colId ? end.colId : start.colId); j++){
            let [topCell, bottomCell, leftCell, rightCell] = getTopBottomLeftRight(i, j);
            selectCell($(`#row-${i}-col-${j}`), {}, topCell, bottomCell, leftCell, rightCell, true);
        }
    }
}

$(".menu-selector").change(function (e) {
    let value = $(this).val();
    let key = $(this).attr("id");
    if (key == 'font-family') {
        $("#font-family").css(key, value);
    }
    if (!isNaN(value)) {
        value = parseInt(value);
    }

    $(".input-cell.selected").css(key, value);
    updateCellData(key, value);
})

//need to change
$("#border").change(function (e) {
    let value = $(this).val();
    $(".input-cell.selected").css("border", "");
    if (value == "none") {
        $(".input-cell.selected").removeClass("border-outer border-left border-right border-top border-bottom");
    }
    else if (value == "outer") {
        $(".input-cell.selected").addClass("border-outer");
    }
    else {
        $(".input-cell.selected").css(`border-${value}`, "3px solid #444");
    }
    updateCellData("border", value);
})

$("#bold").click(function (e) {
    setFontStyle(this, "bold", "font-weight", "bold");
})

$("#italic").click(function (e) {
    setFontStyle(this, "italic","font-style", "italic");
})

$("#underlined").click(function (e) {
    setFontStyle(this, "underlined","text-decoration", "underline"); 
})

function setFontStyle(ele, property, key, value) {
    if ($(ele).hasClass("selected")) {
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(key, "");
        updateCellData(property, false);
    }
    else {
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key, value);
        updateCellData(property, true);
    }
}

$(".alignment").click(function (e) {
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    let alignment = $(this).attr("data-type");
    $(".input-cell.selected").css("text-align", alignment);
    updateCellData("alignment", alignment);
})

function updateCellData(property, value) {
    if (value != defaultProperties[property]) {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCol(data);
            if (cellData[selectedSheet][rowId - 1] == undefined) {
                cellData[selectedSheet][rowId - 1] = {};
                cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties }  //{...array/object} is sparse array or object used to make a copy of array or object
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
            }
            else {
                if (cellData[selectedSheet][rowId - 1][colId - 1] == undefined) {
                    cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties };
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
                else {
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
            }
        });
    }
    else {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCol(data);
            if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                if (JSON.stringify(cellData[selectedSheet][rowId - 1][colId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][rowId - 1][colId - 1];
                    if(Object.keys(cellData[selectedSheet][rowId - 1]).length == 0) {
                        delete cellData[selectedSheet][rowId - 1];
                    }
                }
            }
        });
    }
}

$(".color-pick").colorPick({
    'initialColor': '#TYPECOLOR',
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function() {
        if (this.color != "#TYPECOLOR") {
            if (this.element.attr("id") == "fill-color") {
                $("#fill-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("background-color", this.color);
                updateCellData("bgcolor", this.color);
            }
            else {
                $("#text-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                updateCellData("color", this.color);
            }
        }
    }
});

$("#fill-color-icon, #text-color-icon").click(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 5);
});

$(".container").click(function(e) {
    $(".sheet-options-modal").remove();
    if ($(".sheet-list-modal").hasClass("active")) {
        $(".sheet-list-modal").remove();
    }
    else {
        $(".sheet-list-modal").addClass("active");
    }
});

function selectSheet(ele) {
    $(".sheet-tab.selected").removeClass("selected");
    $(ele).addClass("selected");
    emptySheet();
    selectedSheet = $(ele).text();
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
    loadSheet();
}

function emptySheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for(let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text("");
            cell.css({
                "font-family" : "Noto Sans",
                "font-size" : 14,
                "background-color" : "#fff",
                "color": "#444",
                "font-weight" : "",
                "font-style" : "",
                "text-decoration" : "",
                "text-align": "left",
                "border": ""
            });
            cell.removeClass("border-outer border-top border-left border-bottom border-right");
        }
    }
}

function loadSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for(let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text(data[rowId][colId].text);
            cell.css({
                "font-family" : data[rowId][colId]["font-family"],
                "font-size" : data[rowId][colId]["font-size"] + "px",
                "background-color" : data[rowId][colId]["bgcolor"],
                "color": data[rowId][colId].color,
                "font-weight" : data[rowId][colId].bold ? "bold" : "",
                "font-style" : data[rowId][colId].italic ? "italic" : "",
                "text-decoration" : data[rowId][colId].underlined ? "underline" : "",
                "text-align" : data[rowId][colId].alignment 
            });
            data[rowId][colId]["border"] == "none" ? cell.css("border", "") : cell.addClass(`border-${data[rowId][colId]["border"]}`);
        }
    }
}

$(".add-sheet").click(function (e) {
    emptySheet();
    totalSheets++;
    lastelyAddedSheetNumber++;
    while (Object.keys(cellData).includes("Sheet" + lastelyAddedSheetNumber)) {
        lastelyAddedSheetNumber++;
    }
    cellData[`Sheet${lastelyAddedSheetNumber}`] = {};
    selectedSheet = `Sheet${lastelyAddedSheetNumber}`;
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(
        `<div class="sheet-tab selected">Sheet${lastelyAddedSheetNumber}</div>`
    );
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
    addSheetTabEventListeners();
    $("#row-1-col-1").click();
});

//add sheet tab event listeners
function addSheetTabEventListeners() {
    $(".sheet-tab.selected").bind("contextmenu", function (e) {
        e.preventDefault();
        $(".sheet-options-modal").remove();
        $(".sheet-list-modal").remove();
        let modal = $(
            `<div class="sheet-options-modal">
                <div class="option sheet-rename">Rename</div>
                <div class="option sheet-delete" >Delete</div>
            </div>`
        )
        $(".container").append(modal);
        $(".sheet-options-modal").css({ "bottom": 0.04 * $(".container").height(), "left": e.pageX });
        if (totalSheets == 1) {
            $(".sheet-delete").addClass("disabled");
        }
        $(".sheet-rename").click(function (e) {
            let renameModal = `<div class="sheet-modal-parent">
                                    <div class="sheet-rename-modal">
                                        <div class="sheet-modal-title">
                                            <span>Rename Sheet</span>
                                        </div>
                                        <div class="sheet-modal-input-container">
                                            <span class="sheet-modal-input-title">Rename sheet to:</span>
                                            <input class="sheet-modal-input" type="text">
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button ok-button">OK</div>
                                            <div class="button cancel-button">Cancel</div>
                                        </div>
                                    </div>
                                </div>`;
            
            $(".container").append(renameModal);
            $(".sheet-modal-input").focus();
    
            $(".cancel-button").click(function () {
                $(".sheet-modal-parent").remove();
            })
    
            $(".ok-button").click(function () {
                renameSheet();
            })
    
            $(".sheet-modal-input").keypress(function (e) {
                if (e.key == "Enter") {
                    renameSheet();
                }
            })
        });
    
        if (!$(".sheet-delete").hasClass("disabled")) {
            $(".sheet-delete").click(function (e) {
                let deleteModal = `<div class="sheet-modal-parent">
                                        <div class="sheet-delete-modal">
                                            <div class="sheet-modal-title">
                                                <span>Sheet Name</span>
                                            </div>
                                            <div class="sheet-modal-detail-container">
                                                <span class="sheet-modal-detail-title">Are you Sure?</span>
                                            </div>
                                            <div class="sheet-modal-confirmation">
                                                <div class="button delete-button">
                                                    <span class="material-icons delete-icon">delete</span>
                                                    Delete
                                                </div>
                                                <div class="button cancel-button">Cancel</div>
                                            </div>
                                        </div>
                                    </div>`;
                
                $(".container").append(deleteModal);
        
                $(".cancel-button").click(function () {
                    $(".sheet-modal-parent").remove();
                })
        
                $(".delete-button").click(function (e) {
                    deleteSheet(e);
                })
            });
        }
    
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
        }
    });

    $(".sheet-tab.selected").click(function (e) {
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
            $("#row-1-col-1").click();
        }
    });
}

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val();
    if (newSheetName && !Object.keys(cellData).includes(newSheetName)) {
        let newCellData = {};
        for (let i of Object.keys(cellData)) {
            if (i == selectedSheet) {
                newCellData[newSheetName] = cellData[i];
            }
            else {
                newCellData[i] = cellData[i];
            }
        }
        cellData = newCellData;
        selectedSheet = newSheetName;
                    
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove();
    }
    else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`
            <div class = "error"> Sheet Name is not valid or Sheet already exists </div>
        `);
    }
}

function deleteSheet(ele) {
    $(".sheet-modal-parent").remove();
    let keyArray = Object.keys(cellData);
    let selectedSheetIndex = keyArray.indexOf(selectedSheet);
    let currentSelectedSheet = $(".sheet-tab.selected");
    if (selectedSheetIndex == 0) {
        selectSheet(currentSelectedSheet.next()[0]);
    }
    else {
        selectSheet(currentSelectedSheet.prev()[0]);
    }
    delete cellData[currentSelectedSheet.text()];
    currentSelectedSheet.remove();
    totalSheets--;
}

$(".left-scroller").click(function (e) {

    let keysArray = Object.keys(cellData);  
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if (selectedSheetIndex != 0) {
        selectSheet($(".sheet-tab.selected").prev()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
})

$(".right-scroller").click(function (e) {
    let keysArray = Object.keys(cellData);
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if (selectedSheetIndex != (keysArray.length - 1)) {
        selectSheet($(".sheet-tab.selected").next()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
})

$(".sheet-menu").click(function (e) {
    $(".sheet-list-modal").remove();
    let sheetMenu = $(`<div class="sheet-list-modal">
                        </div>`);
    
    $(".container").append(sheetMenu);
    $(".sheet-list-modal").css({ "bottom": 0.04 * $(".container").height(), "left": e.pageX });
    let sheetArray = $(".sheet-tab");
    for (let i = 0; i < sheetArray.length; i++) {
        sheetMenu.append(`<div class="sheet-list" id=${i + 1}>${sheetArray[i].textContent}</div>`)
    }

    $(".sheet-list").click(function (e) {
        let sheetIndex = $(this).attr("id") - 1;
        selectSheet(sheetArray[sheetIndex]);
    })
    
});

let sheetZoom = 1;
let sheetZoomPercentage = 100;
$("#zoom-percentage").text(sheetZoomPercentage + "%");

$("#zoom-in").click(function (e) {
    if (sheetZoomPercentage <= 200) {
        sheetZoom += 0.1;
        sheetZoomPercentage += 10;
        $(".data-container").css("zoom", sheetZoom);
        $("#zoom-percentage").text(sheetZoomPercentage + "%");
    }
})

$("#zoom-out").click(function (e) {
    if (sheetZoomPercentage >= 20) {
        sheetZoom -= 0.1;
        sheetZoomPercentage -= 10;
        $(".data-container").css("zoom", sheetZoom);
        $("#zoom-percentage").text(sheetZoomPercentage + "%");
    }
})

$(".menu-bar-item").click(function () {
    $(".menu-bar-item.selected").removeClass("selected");
    $(this).addClass("selected");
    if ($(this).text() == "File") {
        let modal = $(`<div class="file-modal">
                            <div class="file-options-modal">
                                <div class="close">
                                    <div class="material-icons close-icon">arrow_circle_down</div>
                                    <div>Close</div>
                                </div>
                                <div class="new">
                                    <div class="material-icons new-icon">insert_drive_file</div>
                                    <div>New</div>
                                </div>
                                <div class="open">
                                    <div class="material-icons open-icon">folder_open</div>
                                    <div>Open</div>
                                </div>
                                <div class="save">
                                    <div class="material-icons save-icon">save</div>
                                    <div>Save</div>
                                </div>
                            </div>
                            <div class="file-recent-modal">

                            </div>
                            <div class="file-transparent-modal"></div>
                        </div>`);
        $(".container").append(modal);
        
        $(".close, .file-transparent-modal").click(function (e) {
            $(".file-modal").remove();
            let currentSelectedMenubaritem = $(".menu-bar-item.selected");
            $(".menu-bar-item.selected").removeClass("selected");
            currentSelectedMenubaritem.next().addClass("selected");
        })
    }
})

