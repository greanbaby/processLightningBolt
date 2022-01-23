// July 2, 2018
// Scott Gingras
const XLSX = require("xlsx");
// library for processing excel files
function openWorkbook (strFile) {
    return XLSX.readFile(strFile);
}
function openWorkbooks (files) {
        let arrWorkbookObjects = [],
            lengthFiles = files.length,
            i;
        for (i = 0; i < lengthFiles; i++) {
            arrWorkbookObjects.push(XLSX.readFile(files[i]));
        }
        return arrWorkbookObjects;
}
function numLines (wbLB) {
    let ws,
        wsName,
        numLines,
        jsonObj;
    wsName = wbLB.SheetNames[0];
    ws = wbLB.Sheets[wsName];
    jsonObj = XLSX.utils.sheet_to_json(ws);
    // numLines will be one less than the actual number of rows in sheet
    // because the header row becomes the json labels
    numLines = jsonObj.length;
    
    return numLines;
}

function numLinesTotal (files) {
    let arrWorkbookObjects = [],
        lineCount = 0,
        lengthObjects,
        i;
    arrWorkbookObjects = openWorkbooks(files);
    lengthObjects = arrWorkbookObjects.length;
    for (i = 0; i < lengthObjects; i++) {
        lineCount = lineCount + numLines(arrWorkbookObjects[i]);
    }
    
    return lineCount;
}

function workTypeByRow (ws, rowNumber) {
    let workTypeName = "";
    if (rowNumber < 1) {
        return workTypeName;
    }
    
    const range = XLSX.utils.decode_range(ws['!ref']);

    // if argument rowNumber is greater than the rows available
    // default to the last row in the worksheet
    if (rowNumber > range.e.r+1) {
        rowNumber = range.e.r+1;
    }
    
    // if there is no value in first column, return the empty string as value
    if (typeof(ws['A' + rowNumber]) === 'undefined') {
        workTypeName = "";
    } else {
        workTypeName = ws['A' + rowNumber].v;
    }
    
    return workTypeName;
}

function nonBlankWorkTypeByRow (ws, rowNumber) {
    // Sometimes the row given will point at a cell that is blank.
    // The way lightning-bolt.com works, it puts blank rows and
    // that user is expected to know that you just look up to the 
    // next row above which is not blank.
    // This is a recursive function that calls itself to keep looking
    // one row above until it finally finds a non-blank value.
    if (rowNumber < 1) {
        return null;
    }
    let cellVal = workTypeByRow(ws, rowNumber);
    if (cellVal === "") {
        cellVal = nonBlankWorkTypeByRow (ws, rowNumber-1);
    }
	
    return cellVal;
}

function getUniquePhysicians (ws) {
    let setPhysicians = new Set();
    const range = XLSX.utils.decode_range(ws['!ref']);
    
    for (let R = range.s.r+1; R <= range.e.r; ++R) {
        for (let C = range.s.c+1; C <= range.e.c; ++C) {
            let cell_address = {c:C, r:R};
            let cell_ref = XLSX.utils.encode_cell(cell_address);
            if (typeof(ws[cell_ref]) !== 'undefined') {
                setPhysicians.add(removeBadCharacters(ws[cell_ref].v));
            }
        }
    }
	
    return setPhysicians;
}

function removeBadCharacters(setVal) {
    //  □  :  \  /  ?  *  [  ]
	let validName = "";
	if (typeof(setVal) !== "string") {
		return "";
	} else {
        validName = setVal.replace("□", "")
            .replace(":", "")
            .replace("\\", "")
            .replace("/", "")
            .replace("[", "")
            .replace("]", "");
	}
	
	return validName.trim();
}

module.exports = {
    getUniquePhysicians,
    nonBlankWorkTypeByRow,
    numLines,
    numLinesTotal,
    openWorkbook,
    openWorkbooks,
    removeBadCharacters,
    workTypeByRow
};
