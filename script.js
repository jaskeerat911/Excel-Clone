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

let cellData = { "Sheet1": [] };
let selectedSheet = "Sheet1";
let totalSheets = 1;
let lastelyAddedSheetNumber = 1;
let mousemoved = false;
let startCellStored = false;
let startCell;
let endCell;

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

function loadNewSheet() {
    $("#cells").text("");
    for (let i = 1; i <= 100; i++) {
        let row = $(`<div class="cell-row"></div>`);
        let rowArray = [];
        for (let j = 1; j <= 100; j++){
            row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`)
            rowArray.push(
                {
                    "font-family": "Noto Sans",
                    "font-size": "14",
                    "text": "",
                    "bold": false,
                    "italic": false,
                    "underlined": false,
                    "alignment": "left",
                    "bgcolor": "#fff",
                    "color": "#444",
                    "border": "none"
                }
            );
        }
        cellData[selectedSheet].push(rowArray);
        $("#cells").append(row);
    }
    addEventsToCells();
    addSheetTabEventListeners();
}

loadNewSheet();

//perfect scrollbar
$("#cells").scroll(function () {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
})

//Add Events to cells
function addEventsToCells() {
    $(".input-cell").dblclick(function () {
        $(this).attr("contenteditable", "true");
        $(this).focus();
    })
    
    $(".input-cell").blur(function () {
        $(this).attr("contenteditable", "false");
        let [rowId, colId] = findRowCol(this);
        cellData[selectedSheet][rowId - 1][colId - 1].text = $(this).text();
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
    let data = cellData[selectedSheet][rowId - 1][colId - 1];
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
    $(".input-cell.selected").each(function (index, data) {
        let [rowId, colId] = findRowCol(data);
        cellData[selectedSheet][rowId - 1][colId - 1][key] = value;
    })
})

$("#border").change(function (e) {
    let value = $(this).val();
    $(".input-cell.selected").css("border", "");
    if (value == "none") {
        $(".input-cell.selected").css("border", "");
    }
    else if (value == "outer") {
        $(".input-cell.selected").addClass("border-outer");
    }
    else {
        $(".input-cell.selected").css(`border-${value}`, "3px solid #444");
    }
    $(".input-cell.selected").each(function (index, data) {
        let [rowId, colId] = findRowCol(data);
        cellData[selectedSheet][rowId - 1][colId - 1].border = value;
    })
})

$("#bold").click(function (e) {
    selectFontStyle(this, "bold", "font-weight", "bold");
})

$("#italic").click(function (e) {
    selectFontStyle(this, "italic","font-style", "italic");
})

$("#underlined").click(function (e) {
    selectFontStyle(this, "underlined","text-decoration", "underline"); 
})

function selectFontStyle(ele, property, key, value) {
    if ($(ele).hasClass("selected")) {
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(key, "");
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCol(data);
            cellData[selectedSheet][rowId - 1][colId - 1][property] = false;
        })
    }
    else {
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key, value);
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCol(data);
            cellData[selectedSheet][rowId - 1][colId - 1][property] = true;
        })
    }
}

$(".alignment").click(function (e) {
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    let alignment = $(this).attr("data-type");
    $(".input-cell.selected").css("text-align", alignment);
    $(".input-cell.selected").each(function (index, data) {
        let [rowId, colId] = findRowCol(data);
        cellData[selectedSheet][rowId - 1][colId - 1].alignment = alignment;
    })
})

$(".color-pick").colorPick({
    'initialColor': '#TYPECOLOR',
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': false,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function() {
        if (this.color != "#TYPECOLOR") {
            if (this.element.attr("id") == "fill-color") {
                $("#fill-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("background-color", this.color);
                $(".input-cell.selected").each((index, data) => {
                    let [rowId, colId] = findRowCol(data);
                    cellData[selectedSheet][rowId - 1][colId - 1].bgcolor = this.color;
                })
            }
            else {
                $("#text-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                $(".input-cell.selected").each((index, data) => {
                    let [rowId, colId] = findRowCol(data);
                    cellData[selectedSheet][rowId - 1][colId - 1].color = this.color;
                })
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
    addLoader();
    $(".sheet-tab.selected").removeClass("selected");
    $(ele).addClass("selected");
    selectedSheet = $(ele).text();
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
    setTimeout(() => {
        loadSheet();
        removeLoader();
    }, 3);
}

function loadSheet() {
    $("#cells").text("");
    let data = cellData[selectedSheet];
    for (let i = 1; i <= data.length; i++){
        let row = $(`<div class="cell-row"></div>`);
        for (let j = 1; j <= data[i-1].length; j++){
            let cell = $(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false">${data[i - 1][j - 1].text}</div>`)
            cell.css({
                "font-family": data[i - 1][j - 1]["font-family"],
                "font-size": data[i - 1][j - 1]["font-size"] + "px",
                "background-color": data[i - 1][j - 1]["bg-color"],
                "color": data[i - 1][j - 1].color,
                "font-weight": data[i - 1][j - 1].bold ? "bold" : "",
                "font-style": data[i - 1][j - 1].italic ? "italic" : "",
                "text-decoration": data[i - 1][j - 1].underlined ? "underline" : "",
                "text-align": data[i - 1][j - 1].alignment
            });
            data[i - 1][j - 1]["border"] == "none" ? cell.css("border", "") : cell.addClass(`border-${data[i - 1][j - 1]["border"]}`);
            row.append(cell);
        }
        $("#cells").append(row);
    }

    addEventsToCells();
}

function addLoader() {
    $(".container").append(`<div class="sheet-modal-parent">
                                <div class="loading"></div>
                            </div>`
    )
}

function removeLoader() {
    $(".sheet-modal-parent").remove();
}

$(".add-sheet").click(function (e) {
    addLoader();
    totalSheets++;
    lastelyAddedSheetNumber++;
    while (Object.keys(cellData).includes("Sheet" + lastelyAddedSheetNumber)) {
        lastelyAddedSheetNumber++;
    }
    cellData[`Sheet${lastelyAddedSheetNumber}`] = [];
    selectedSheet = `Sheet${lastelyAddedSheetNumber}`;
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(
        `<div class="sheet-tab selected">Sheet${lastelyAddedSheetNumber}</div>`
    );
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
    
    setTimeout(() => {
        loadNewSheet();
        removeLoader();
    }, 3);
})

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
        }
    });
}

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val();
    if (newSheetName && !Object.keys(cellData).includes(newSheetName)) {
        cellData[newSheetName] = cellData[selectedSheet];
        delete cellData[selectedSheet];
        selectedSheet = newSheetName;

        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove();
    }
    else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(
            `<div class = "error">Sheet Name is not valid or Sheet alredy exists </div>`
        )
    }
}

function deleteSheet(ele) {
    $(".sheet-modal-parent").remove();
    let keyArray = Object.keys(cellData);
    let selectedSheetIndex = keyArray.indexOf(selectedSheet);
    let currentSelectedSheet = $(".sheet-tab.selected");
    delete cellData[selectedSheet];
    if (selectedSheetIndex == 0) {
        selectSheet(currentSelectedSheet.next()[0]);
        currentSelectedSheet.remove();
    }
    else {
        selectSheet(currentSelectedSheet.prev()[0]);
        currentSelectedSheet.remove();
    }
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
                            <div class="file-options-container">
                                <div id="close-file-modal">
                                    <span class="material-icons file-option-icon" id="file-close-icon">west</span>
                                    Close
                                </div>
                                <div class="file-option selected" id="file-option-home">
                                    <span class="material-icons file-option-icon">home</span>
                                    Home
                                </div>
                                <div class="file-option" id="file-option-new">
                                    <span class="material-icons file-option-icon">description</span>
                                    New
                                </div>
                                <div class="file-option" id="file-option-open">
                                    <span class="material-icons file-option-icon">folder_open</span>
                                    Open
                                </div>
                            </div>
                            <div class="file-option-preview-container">
                                <div class="home-file-option">
                                    <div class="file-option-title">Home</div>
                                    <div class="home-option-section">
                                        <div class="file-option-label">New</div>
                                        <div class="new-label-container">
                                            <div class="new-file-container">
                                                <div class="new-file">
                                                    <div class="material-icons new-file-icon">add</div>
                                                    <div class="new-file-title">New</div>
                                                </div>
                                            </div>
                                            <div class="new-file-container">
                                                <div class="new-file">
                                                    <div class="material-icons new-file-icon">add</div>
                                                    <div class="new-file-title">New</div>
                                                </div>
                                            </div>
                                            <div class="new-file-container">
                                                <div class="new-file">
                                                    <div class="material-icons new-file-icon">add</div>
                                                    <div class="new-file-title">New</div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="home-option-section">
                                        <div class="file-option-label">Recent</div>
                                        <div class="recent-label-container">
                                            <div class="recent-file">
                                                <img class="excel-file-icon"
                                                    src="https://img.icons8.com/color/35/000000/ms-excel.png" />
                                                <div class="recent-file-title-container">
                                                    <div class="recent-file-title">Book 1.xlsx</div>
                                                    <div class="recent-file-location">Personal</div>
                                                </div>
                                            </div>
                                            <div class="recent-file">
                                                <img class="excel-file-icon"
                                                    src="https://img.icons8.com/color/35/000000/ms-excel.png" />
                                                <div class="recent-file-title-container">
                                                    <div class="recent-file-title">Book 1.xlsx</div>
                                                    <div class="recent-file-location">Personal</div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>`)
        $(".container").append(modal);

        $("#close-file-modal").click(function () {
            $(".file-modal").remove();
            let currentSelectedMenubaritem = $(".menu-bar-item.selected");
            $(".menu-bar-item.selected").removeClass("selected");
            currentSelectedMenubaritem.next().addClass("selected");
        })
        
        $(".file-option").click(function () {
            $(".file-option.selected").removeClass("selected");
            $(this).addClass("selected");
        })

        $("#file-option-home").click(function (e) {
            $(".file-option-preview-container").text("");
            let optionContainer = $(`<div class="home-file-option">
                                        <div class="file-option-title">Home</div>
                                        <div class="home-option-section">
                                            <div class="file-option-label">New</div>
                                            <div class="new-label-container">
                                                <div class="new-file-container">
                                                    <div class="new-file">
                                                        <div class="material-icons new-file-icon">add</div>
                                                        <div class="new-file-title">New</div>
                                                    </div>
                                                </div>
                                                <div class="new-file-container">
                                                    <div class="new-file">
                                                        <div class="material-icons new-file-icon">add</div>
                                                        <div class="new-file-title">New</div>
                                                    </div>
                                                </div>
                                                <div class="new-file-container">
                                                    <div class="new-file">
                                                        <div class="material-icons new-file-icon">add</div>
                                                        <div class="new-file-title">New</div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="file-option-section">
                                            <div class="file-option-label">Recent</div>
                                            <div class="recent-label-container">
                                                <div class="recent-file">
                                                    <img class="excel-file-icon" src="https://img.icons8.com/color/35/000000/ms-excel.png"/>
                                                    <div class="recent-file-title-container">
                                                        <div class="recent-file-title">Book 1.xlsx</div>
                                                        <div class="recent-file-location">Personal</div>
                                                    </div>
                                                </div>
                                                <div class="recent-file">
                                                    <img class="excel-file-icon" src="https://img.icons8.com/color/35/000000/ms-excel.png"/>
                                                    <div class="recent-file-title-container">
                                                        <div class="recent-file-title">Book 1.xlsx</div>
                                                        <div class="recent-file-location">Personal</div>
                                                    </div>
                                                </div>
                                                <div class="recent-file">
                                                    <img class="excel-file-icon" src="https://img.icons8.com/color/35/000000/ms-excel.png"/>
                                                    <div class="recent-file-title-container">
                                                        <div class="recent-file-title">Book 1.xlsx</div>
                                                        <div class="recent-file-location">Personal</div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>`)
            
            $(".file-option-preview-container").append(optionContainer);
        })
        
        $("#file-option-new").click(function (e) {
            $(".file-option-preview-container").text("");
            let optionContainer = $(`<div class="new-file-option">
                                        <div class="file-option-title">New</div>
                                        <div class="new-option-section">
                                            <div class="new-file-container">
                                                <div class="new-file">
                                                    <div class="material-icons new-file-icon">add</div>
                                                    <div class="new-file-title">New</div>
                                                </div>
                                            </div>
                                            <div class="new-file-container">
                                                <div class="new-file">
                                                    <div class="material-icons new-file-icon">add</div>
                                                    <div class="new-file-title">New</div>
                                                </div>
                                            </div>
                                            <div class="new-file-container">
                                                <div class="new-file">
                                                    <div class="material-icons new-file-icon">add</div>
                                                    <div class="new-file-title">New</div>
                                                </div>
                                            </div>
                                            <div class="new-file-container">
                                                <div class="new-file">
                                                    <div class="material-icons new-file-icon">add</div>
                                                    <div class="new-file-title">New</div>
                                                </div>
                                            </div>
                                            <div class="new-file-container">
                                                <div class="new-file">
                                                    <div class="material-icons new-file-icon">add</div>
                                                    <div class="new-file-title">New</div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>`)
            
            $(".file-option-preview-container").append(optionContainer);
        })
        $("#file-option-open").click(function (e) {
            $(".file-option-preview-container").text("");
            let optionContainer = $(`<div class="open-file-option">
                                        <div class="file-option-title">Open</div>
                                        <div class="open-option-section">
                                            <div class="file-option-label">Recent</div>
                                            <div class="recent-label-container">
                                                <div class="recent-file">
                                                    <img class="excel-file-icon"
                                                        src="https://img.icons8.com/color/35/000000/ms-excel.png" />
                                                    <div class="recent-file-title-container">
                                                        <div class="recent-file-title">Book 1.xlsx</div>
                                                        <div class="recent-file-location">Personal</div>
                                                    </div>
                                                </div>
                                                <div class="recent-file">
                                                    <img class="excel-file-icon"
                                                        src="https://img.icons8.com/color/35/000000/ms-excel.png" />
                                                    <div class="recent-file-title-container">
                                                        <div class="recent-file-title">Book 1.xlsx</div>
                                                        <div class="recent-file-location">Personal</div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>`)
            
            $(".file-option-preview-container").append(optionContainer);
        })
    }
})



