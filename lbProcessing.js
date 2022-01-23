'use strict';
// July 17, 2018
// Scott Gingras
const fs = require("fs");
const XLSX = require("xlsx");
const lb = require("./lb");
const lbWriteExcel = require("./lbWriteExcel");
const lbSchedules = "../lbSchedules";

// INSTRUCTIONS:
// Open lightning-bolt.com file(s) to analyze how much physicians are going to work
// Using Default View looking at monthly schedule:
// Step 1) download 3 months of exported schedules as "xls" files
// Step 2) inside each file, copy the word "Assignment" so it is on each row with date headers
// Step 3) delete column A and columns H:I (sat:sun)
// Step 4) save each week as separate xlsx file (keep date headers for each file)
// Step 5) update "const scheduleFiles" to point to these files :
const scheduleFiles = [

 
  lbSchedules + "/dec9-dec13.xlsx",
  lbSchedules + "/dec16-dec20.xlsx",
  lbSchedules + "/dec23-dec27.xlsx",
  lbSchedules + "/dec30-jan3.xlsx"
  
  
    ];

try {
  // lbProcessing.js
  // calculates the timeslots worked for all physicians
  // based on exported lightning-bolt.com schedule files
  // process all workbooks specified in scheduleFiles array
  const wbObjList = lb.openWorkbooks(scheduleFiles);

	logMessage("Started " + (new Date()));
  logMessage(scheduleFiles);
    
  // prepare an object with properties for all physicians found in all scheduleFiles
  let physiciansList = getPhysiciansFromAllWorkbooks(wbObjList);

  // process each workbook to get dates working for each physician
  for (let i = 0; i < wbObjList.length; i++) {
      const ws = wbObjList[i].Sheets[wbObjList[i].SheetNames[0]];
      processScheduleWorksheet(ws, physiciansList);
  }
  
  analyzePhysiciansList(physiciansList);
    
	logMessage("Finished " + (new Date()));
} catch(err) {
    logMessage(err);
}
function logMessage (msg) {
	const timestamp = new Date();
	const timeval = timestamp.toDateString() + 
	    " " + timestamp.getHours() + 
		":" + timestamp.getMinutes() + 
		":" + timestamp.getSeconds();
	fs.appendFile('lbProcessingLog.txt', timeval + " " + msg + "\r\n", (err) => {
        if (err) throw err;
	});
    console.log(msg);
}
function getPhysiciansFromAllWorkbooks (wbObjList) {
  // (wbObjList) is an array of XLSX workbook objects exported from lightning-bolt.com
  try {
      let setPhysicianNames = new Set();
      let physiciansList = {};
      for (let i = 0; i < wbObjList.length; i++) {
        const ws = wbObjList[i].Sheets[wbObjList[i].SheetNames[0]];
        setPhysicianNames = union(setPhysicianNames,lb.getUniquePhysicians(ws));
      }
      for (let physicianName of setPhysicianNames) {
        // create a blank object for each physicianName in the set
        physiciansList[physicianName] = { "datesWorking" : [] };
      }
      return physiciansList;
  } catch(err) {
      logMessage("getPhysiciansFromAllWorkbooks ERROR: " + err);
  }
}
function union(setA, setB) {
    let _union = new Set(setA);
    for (let elem of setB) {
        _union.add(elem);
    }
    return _union;
}
function processScheduleWorksheet (scheduleWS, physiciansList) {
  // (scheduleWS) is a single XLSX worksheet object
	// (physiciansList) : for every name found in the schedule 
	// it becomes a property in physiciansList object
	// physiciansList[physicianName] = { "datesWorking" : [] }
	// e.g. physiciansList = {
  // "Dr. Scrimshaw" : { datesWorking:[] },
	// "Dr. Parker" : { datesWorking:[] }
	// }
	// WHERE datesWorking example:
	// { "dayOfWeek" : "Monday",
	// "dateWorking" : "6 / 25",
	// "workType" : "AllDayInClinic" }
    try {
        const range = XLSX.utils.decode_range(scheduleWS['!ref']);
        // process columns 1 thru 5 (Monday thru Friday)
        for (let i = 1; i < 6; i++) {
            processColumn(i,scheduleWS,range,physiciansList);
        }
    } catch(err) {
        logMessage("processScheduleWorksheet ERROR: " + err);
    }
}
function processColumn (column, ws, range,physiciansList) {
  // (column) is the integer representing the column to process
  // (ws) is the schedule worksheet containing 6 columns
  // (range) is the range object, range.e.r is last row
	// (physiciansList) is : physiciansList[physicianName] = { "datesWorking" : [] }
  let physicianName;
  try {
      const dayOfWeek = getDayOfWeek(column);
      const dateVal = getDateVal(column, ws);
      // loop through all of the rows for this column
      for (let i = 1; i < range.e.r+1; i++) {
        if (typeof(ws[XLSX.utils.encode_cell({c:column, r:i})]) !== 'undefined') {
          // get physician name from cell
          physicianName = ws[XLSX.utils.encode_cell({c:column, r:i})].v;
          
          // remove the "bad characters" that are in some cell values
          physicianName = lb.removeBadCharacters(physicianName);
          
          // get work type using lb.nonBlankWorkTypeByRow(ws,i+1)
          // to recursively look for the next non-blank value above
          let workingObj = generateWorkingObject(dayOfWeek, dateVal, ws, i);

          // add object containing worktype into the datesWorking property
          let currentDatesWorking = physiciansList[physicianName].datesWorking;
          currentDatesWorking.push(workingObj);
          }
      }
    } catch(err) {
        if (err instanceof TypeError) {
            logMessage("processColumn ERROR: " + err + " " + physicianName);
        } else {
            logMessage("processColumn ERROR: " + err);
        }
    }
}
function generateWorkingObject (dayOfWeek, dateVal, ws, i) {
  try {
      let workingObj = {
        "dayOfWeek" : dayOfWeek,
        "dateWorking" : dateVal,
        "workType" : lb.nonBlankWorkTypeByRow(ws,i+1)
      };
    return workingObj;
  } catch(err) {
    logMessage("generateWorkingObject ERROR: " + err);
  }
}
function getDateVal (column,ws) {
    // (column) is the column to get value from and (ws) is the worksheet
    try {
        return ws[XLSX.utils.encode_cell({c:column, r:0})].v;
    } catch(err) {
        logMessage("getDateVal ERROR: " + err);
    }
}
function getDayOfWeek (column) {
  // (column) is the integer to be converted into the Day of Week it represents in lightning-bolt.com
  let dayOfWeek;
    switch(column) {
    case 1:
      dayOfWeek = "Monday";
      break;
    case 2:
      dayOfWeek = "Tuesday";
      break;
    case 3:
      dayOfWeek = "Wednesday";
      break;
    case 4:
      dayOfWeek = "Thursday";
      break;
    case 5:
      dayOfWeek = "Friday";
      break;
  }
  return dayOfWeek;
}
// ***************************************************************************
// ** process each physician name found in the lb schedules exported
function analyzePhysiciansList (physiciansList) {
    // get actual timeslots working for each physician by going through each 
    // date the individual physician is working and then looking at the list of
    // work types they are working for that specific day of the week and
    // calculating the # of timeslots worked that day by the time started in morning,
    // amount of time taken for lunch, amount of clinic time spent on surgery days, etc.
    // at this point physiciansList is now an object populated with all physicians and their days worked

    let totalTimeslotsAllPhysicians = 0;  // running total for entire clinic

    logMessage("Successfully read lightning-bolt.com Excel exports...now analyze physicians list");
	
    for (let physician in physiciansList) {
      if (physiciansList.hasOwnProperty(physician)) {	
        const currentDatesWorking = physiciansList[physician].datesWorking;
        const currentUniqueDates = getUniqueDateValues(currentDatesWorking);

        // calculate the timeslots worked based on the array of workTypes
        processDates(physician,currentUniqueDates);
          
        // now at this point currentUniqueDates will have the time worked per date;
        // add up the totals for individual physician
        const totalTimeslots = sumTimeslotsWorked(currentUniqueDates);
        logMessage(physician + " TOTAL: " + totalTimeslots);

        // add up the running total for the entire clinic
        totalTimeslotsAllPhysicians = totalTimeslotsAllPhysicians + totalTimeslots;

        // save currentUniqueDates into physician object
        physiciansList[physician].datesWorkingUnique = currentUniqueDates;
      }
    }
    writeSummary(physiciansList);
    logMessage("TOTAL FOR CLINIC: " + totalTimeslotsAllPhysicians);
}
function writeSummary (physiciansList) {
    // write out the details for entire clinic contained in physiciansList object
    /*
    physician
    physiciansList[physician].datesWorkingUnique object is
        datesWorkingUnique:
        { '6 ∕ 25': { workTypes: [ 'Until 3 In Clinic', 'ECG' ], dayOfWeek: 'Monday', timeslotsWorked: 0 },
        '8 ∕ 3': { workTypes: [ 'Endoscopy', 'Not Returning', 'ECG' ], dayOfWeek: 'Friday', timeslotsWorked: 18 } }
    */
    
	/*
    // text file written using logMessage
    for (let physician in physiciansList) {
        if (physiciansList.hasOwnProperty(physician)) {
            const objDatesWorkingUnique = physiciansList[physician].datesWorkingUnique;
            for (let date in objDatesWorkingUnique) {
                if (objDatesWorkingUnique.hasOwnProperty(date)) {
                    logMessage( objDatesWorkingUnique[date].timeslotsWorked + " timeslots " + 
                        physician + " " + 
                        objDatesWorkingUnique[date].dayOfWeek + " " + date +
                        " " + objDatesWorkingUnique[date].workTypes.join() );
                }
            }
        }
    }
    */
	
    // Excel file written using lbWriteExcel.writeExcelPhysicianSummary
  try {
    lbWriteExcel.writeExcelPhysicianSummary(physiciansList);
  } catch(err) {
    logMessage("writeSummary " + err);
  }
    
}
function sumTimeslotsWorked(currentUniqueDates) {
	// return total of all timeslotsWorked properties of each uniqueDate object 
	let totalSum = 0;
	try {
		for (let uniqueDate in currentUniqueDates) {
      if (currentUniqueDates.hasOwnProperty(uniqueDate)) {
        totalSum = totalSum + currentUniqueDates[uniqueDate].timeslotsWorked;
      }
    }		
		return totalSum;
	} catch(err) {
		logMessage("sumTimeslotsWorked ERROR: " + err);
	}
}
function getUniqueDateValues (currentDatesWorking) {
    // currentDatesWorking will have multiple "dateWorking" values that are duplicated,
    // each one having a different workType
    // create an object with each dateWorking contained only once, and each workType
    // attached in an array to the unique dateWorking
    let setUniqueDates = new Set();
    let myUniqueDates = {};
    try {
        const currentDatesWorkingLength = currentDatesWorking.length;
        for (let i = 0; i < currentDatesWorkingLength; i++) {
            setUniqueDates.add(currentDatesWorking[i].dateWorking);
        }
        for (let dateWorkingValue of setUniqueDates) {
            myUniqueDates[dateWorkingValue] = { "workTypes" : [] };
        }
        for (let j = 0; j < currentDatesWorkingLength; j++) {
            const currentDateWorking = currentDatesWorking[j].dateWorking;
            const currentWorkType = currentDatesWorking[j].workType;
            const currentDayOfWeek = currentDatesWorking[j].dayOfWeek;

            let currentWorkTypes = myUniqueDates[currentDateWorking].workTypes;
            currentWorkTypes.push(currentWorkType);

            myUniqueDates[currentDateWorking].dayOfWeek =
                currentDayOfWeek;
        }
        return myUniqueDates;
    } catch(err) {
        logMessage("getUniqueDateValues ERROR: " + err);
    }
}
function processDates (physician, currentUniqueDates) {
    // using the physician and dayOfWeek;
    // calculate the timeslots worked based on the array of workTypes
    try {
        for (let uniqueDate in currentUniqueDates) {
            if (currentUniqueDates.hasOwnProperty(uniqueDate)) {
                const timeslotsWorkedVal = fncTimeslotsWorked(physician,
                    currentUniqueDates[uniqueDate].dayOfWeek,
                    currentUniqueDates[uniqueDate].workTypes);
                currentUniqueDates[uniqueDate].timeslotsWorked = timeslotsWorkedVal;
            }
        }
    } catch(err) {
        logMessage("processDates ERROR: " + err);
    }
}
// ***************************************************************************
// DETAILS FOR EACH INDIVIDUAL PHYSICIAN
// ON EACH DIFFERENT DAY OF WEEK
function fncTimeslotsWorked (physician, dayOfWeek, workTypes) {
    // calculate timeslots worked based on physician rules for this day of the week
    // depending on the total list of workTypes listed
    let timeslotsWorked = 0;
    
	// "Clinic Call"
	if (workTypes.indexOf("Clinic Call") !== -1) {
		if (workTypes.indexOf("ROC AM") !== -1) {
			timeslotsWorked = 0;
		}
		if (physician === "Dr.Hosford" || physician === "Dr. Duke") {
			timeslotsWorked = 0;
		} else if (physician === "Dr. Ruttle") {
			timeslotsWorked = 13;
		} else if (physician === "Dr. Myhr") {
			timeslotsWorked = 7;
		} else if (physician === "Dr. Rommens") {
			timeslotsWorked = 12;
		}
		return timeslotsWorked;
	}
	
    // "ROC AM"
    if (workTypes.indexOf("ROC AM") !== -1) {
        if (physician === "Dr. Parker") {

			//Parker
			if (dayOfWeek === "Friday") {
                timeslotsWorked = 6;
            } else {
                timeslotsWorked = 10;
            }
			
        } else if (physician === "Dr. Gelber") {

			//Gelber
            if (dayOfWeek==="Friday") {
                timeslotsWorked = 5;
            } else {
                timeslotsWorked = 9;
            }
			
        } else {
			
			//All EXCEPT Gelber or Parker
            timeslotsWorked = 3;
        }
		
		// "Night clinic"
		if (workTypes.indexOf("Night clinic") !== -1) {
			if (physician==="Dr. Gelber") {
				timeslotsWorked = timeslotsWorked + 12;
			} else {
				timeslotsWorked = timeslotsWorked + 10;
			}
		}
        return timeslotsWorked;
    }
    
    // "Brocket"
    if (workTypes.indexOf("Brocket") !== -1) {
        if (workTypes.indexOf("Not Returning") !== -1) {
            timeslotsWorked = 16;
        } else if (workTypes.indexOf("Until Noon In Clinic") !== -1) {
            timeslotsWorked = 18;
        } else {
			if (physician === "Dr. Myhr" || physician === "Dr. Rommens") {
				timeslotsWorked = 20;
			} else {
				timeslotsWorked = 22;
			}
        }
        return timeslotsWorked;
    }
    
    // "AllDayInClinic"
    if (workTypes.indexOf("AllDayInClinic") !== -1) {
		
        if (dayOfWeek === "Friday") {
			// Friday
			if (physician === "Dr. Parker") {
				timeslotsWorked = 16;
			} else if (physician === "Dr.Hosford") {
				timeslotsWorked = 12;
			} else {
				timeslotsWorked = 18;
			}
			
		} else if (dayOfWeek === "Tuesday") {
			// Tuesday
			if (physician === "Dr. Myhr" || physician === "Dr.Hosford" || physician === "Dr. Ruttle") {
				timeslotsWorked = 22;
			} else if (physician === "Dr. Gelber") {
				timeslotsWorked = 26;
			} else {
				timeslotsWorked = 24;
			}
			
        } else {
			// Mon; Wed; Thurs
            if (physician === "Dr. Gelber") {
				
                if (workTypes.indexOf("Practice MTG") !== -1) {
                    timeslotsWorked = 22;
                } else {
                    timeslotsWorked = 28;
                }
				
			} else if (physician === "Dr. Myhr" || physician === "Dr.Hosford" || physician === "Dr. Ruttle") {
				
				if (workTypes.indexOf("Practice MTG") !== -1) {
                    timeslotsWorked = 20;
                } else {
                    timeslotsWorked = 24;
                }

			} else {
				
                if (workTypes.indexOf("Practice MTG") !== -1) {
                    timeslotsWorked = 20;
                } else{
                    timeslotsWorked = 26;
                }
				
            }
        }
        return timeslotsWorked;
    }
    
    // "Until 3 In Clinic"
    if (workTypes.indexOf("Until 3 In Clinic") !== -1) {
        if (workTypes.indexOf("Start 1 In Clinic") !== -1) {
            timeslotsWorked = 8;
        } else {
            if (dayOfWeek === "Friday") {
                timeslotsWorked = 12;
            } else {
                if (physician === "Dr. Gelber") {
                    timeslotsWorked = 20;
                } else {
                    timeslotsWorked = 18;
                }
            }
        }
        return timeslotsWorked;
    }
    
    // "Until Noon In Clinic"
    if (workTypes.indexOf("Until Noon In Clinic") !== -1) {
        if (dayOfWeek === "Friday") {
            timeslotsWorked = 6;
        } else {
            if (physician === "Dr. Parker" || physician === "Dr. Gelber") {
                timeslotsWorked = 14;
            } else {
                timeslotsWorked = 12;
            }
        }
        return timeslotsWorked;
    }
    
    // "Start 1 In Clinic"
    if (workTypes.indexOf("Start 1 In Clinic") !== -1) {
        if (dayOfWeek === "Friday") {
			if (physician === "Dr.Hosford") {
				timeslotsWorked = 8;
			} else {
            timeslotsWorked = 10;
			}
        } else {
            if (workTypes.indexOf("Practice MTG") !== -1) {
                    timeslotsWorked = 6;
                } else {
					
					if (physician === "Dr. Myhr" || physician === "Dr.Hosford" || physician === "Dr. Ruttle") {
						timeslotsWorked = 10;
					} else {
						timeslotsWorked = 12;
					}
					
                }
        }
        return timeslotsWorked;
    }
    
    // "Start 3 In Clinic"
    if (workTypes.indexOf("Start 3 In Clinic") !== -1) {
        if (dayOfWeek === "Friday") {
            timeslotsWorked = 6;
        } else {
            if (workTypes.indexOf("Practice MTG") !== -1) {
                timeslotsWorked = 2;
            } else{
				
				if (physician === "Dr. Myhr" || physician === "Dr.Hosford" || physician === "Dr. Ruttle") {
					timeslotsWorked = 6;
				} else {
					timeslotsWorked = 8;
				}
				
            }
        }
        return timeslotsWorked;
    }
    
    // default return
    return timeslotsWorked;
    
}
