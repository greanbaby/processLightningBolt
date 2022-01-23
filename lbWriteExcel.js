// October 27, 2018
// Scott Gingras
const XLSX = require("xlsx");
// library for writing excel files

function writeExcelPhysicianSummary ( physiciansList ) {
    // create Excel workbook
    let wb = XLSX.utils.book_new();
    wb.Props = {
        Title: "The Associate Clinic Lightning-Bolt.com Supply",
        Author: "Scott Gingras"
    };
    /*
    physician
    physiciansList[physician].datesWorkingUnique object is
        datesWorkingUnique:
        { '6 ∕ 25': { workTypes: [ 'Until 3 In Clinic', 'ECG' ], dayOfWeek: 'Monday', timeslotsWorked: 0 },
        '8 ∕ 3': { workTypes: [ 'Endoscopy', 'Not Returning', 'ECG' ], dayOfWeek: 'Friday', timeslotsWorked: 18 } }
    */
    for (let physician in physiciansList) {
        if (physiciansList.hasOwnProperty(physician)) {
            // create an array of arrays for all unique dates in physiciansList
            wb.SheetNames.push(physician);
            let x = 0;
            let ws_data = [[]];
            const objDatesWorkingUnique = physiciansList[physician].datesWorkingUnique;
            for (let date in objDatesWorkingUnique) {
                if (objDatesWorkingUnique.hasOwnProperty(date)) {
                    ws_data[x] = [objDatesWorkingUnique[date].timeslotsWorked,
                        objDatesWorkingUnique[date].dayOfWeek,
                        date,
                        objDatesWorkingUnique[date].workTypes.join()
                    ];
                    x++;
                }
            }
            // convert ws_data array of arrays to ws and add this physician worksheet into the workbook wb
            const ws = XLSX.utils.aoa_to_sheet(ws_data);
            wb.Sheets[physician] = ws;
        }
    }
    const saveFileName = "../output/supplyCalculator " + generateUniqueFileName();
    write_opts = {};
    XLSX.writeFile(wb, saveFileName, write_opts);
}

function generateUniqueFileName() {
    const timestamp = new Date();
	let timeval = timestamp.toDateString() + 
	    " " + timestamp.getHours() + 
		"-" + timestamp.getMinutes() + 
		"-" + timestamp.getSeconds();
        
    return timeval + ".xlsx";
}

function testWriteExcel() {
    // create Excel workbook
    let wb = XLSX.utils.book_new();    
    // set workbook properties
    wb.Props = {
        Title: "Test SG",
        Author: "Scott Gingras"
    };
    // now we have the Excel workbook object, assign new sheet name to SheetNames array
    wb.SheetNames.push("TestSG");
    // use Array of Arrays option to generate worksheet data
    let ws_data = [['A1', 'A2'], ['B1', 'B2', 'B3']];
    ws_data[2] = ['C1', 'C2'];
    ws_data[3] = ['D1', 'D2', 'D3'];
    ws_data.push(['E1', 'E2']);
    // create worksheet from Array of Arrays using aoa_to_sheet()
    let ws = XLSX.utils.aoa_to_sheet(ws_data);
    // assign the sheet object to the workbook Sheets array
    wb.Sheets["TestSG"] = ws;
    // get unique filename to save under
    const saveFileName = "../output/testWriteExcel " + generateUniqueFileName();
    write_opts = {};
    XLSX.writeFile(wb, saveFileName, write_opts);
}

module.exports = {
    testWriteExcel,
    writeExcelPhysicianSummary
};
