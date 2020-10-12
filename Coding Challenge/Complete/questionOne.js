
if (typeof require !== 'undefined') XLSX = require('xlsx');
const WorkBookUtils = require('./WorkBookUtils');
var Utils = WorkBookUtils();

//error reports
console.log("Invalid Email Addresses")
InvalidEntry("I", "@.+(\.com|\.edu|\.net|\.org)$", 2);
console.log("")
console.log("Invalid Phone Numbers")
InvalidEntry("H", "\\d\\d\\d-\\d\\d\\d-\\d\\d\\d\\d", 2);



//1
withinADay();

//2
console.log(Utils.getHeaderColumn(getSheet(),"asdf"));

//3

function withinADay() {
    //checks to see if the orderDate and ShipDate are less than 24 hours apart
    let dateData = getSheetData(2, ["A", "L", "M"]);

    //make excel sheet


    //go though data
    for (let i = 0; i < dateData.length; i++) {

        //I cant find a date parser!!!!!
        //so i guess im doing it by hand!
        let orderDate = (dateData[i])["L"].match("(\\d\\dT)")[0];
        orderDate = orderDate.substring(0, orderDate.length - 1);
        let ordertime = (dateData[i])["L"].match("\\d\\d:\\d\\d:\\d\\d")[0];

        let sendDate = (dateData[i])["M"].match("(\\d\\dT)")[0];
        sendDate = sendDate.substring(0, sendDate.length - 1);
        let sendTime = (dateData[i])["M"].match("\\d\\d:\\d\\d:\\d\\d")[0];

        if (parseInt(orderDate) + 1 === parseInt(sendDate)) {
            //next day

            let orderTimeSplit = ordertime.split(":");
            let sendTimeSplit = sendTime.split(":");
            console.log(dateData[i]["A"] + " " + parseInt(orderDate) + " " + parseInt(sendDate));
            console.log(dateData[i]["A"] + " " + orderTimeSplit[0]);

            //checks the hours:minutes:seconds to see if they are less than 24 hours
            if (parseInt(sendTimeSplit[0]) < parseInt(orderTimeSplit[0])) {
                console.log("Hour: " + dateData[i]["A"] + " " + ordertime + " " + sendTime);
            } else if (parseInt(sendTimeSplit[0]) === parseInt(orderTimeSplit[0])) {
                if (parseInt(sendTimeSplit[1]) < parseInt(orderTimeSplit[1])) {
                    console.log("Minute: " + dateData[i]["A"] + " " + ordertime + " " + sendTime);
                } else if (parseInt(sendTimeSplit[1]) === parseInt(orderTimeSplit[1])) {
                    if (parseInt(sendTimeSplit[2]) <= parseInt(orderTimeSplit[2])) {
                        console.log("Second: " + dateData[i]["A"] + " " + ordertime + " " + sendTime);
                    }
                }
            }

        } else if (parseInt(orderDate) === parseInt(sendDate)) {
            //same day

            console.log("same Day: " + dateData[i]["A"] + " " + parseInt(orderDate) + " " + parseInt(sendDate));
        }
    }
}

function getSheetData(startRow, columns) {
    //returns an array of objects that holds the data rom the columns
    //colums should be an array of column you want data from

    let result = [];

    let worksheet = getSheet();

    let rowCounter = startRow;
    let cellAdress = "A" + rowCounter;
    let cellValue = Utils.getCellValue(cellAdress, worksheet);

    while (cellValue !== undefined) {

        result.push(Utils.getRowData(worksheet, rowCounter, columns));

        rowCounter++;
        cellAdress = "A" + rowCounter;
        cellValue = Utils.getCellValue(cellAdress, worksheet);
    }

    return result;
}





/*
function getHeaderWidth(worksheet){
    //return the number of columns in the header
    //assume: no empty spaces in the header
    //assume: header cells are not merged
    //assume: header is on row 1

    let result = 0;
    let cell = getCellValue("A1",worksheet);

    while(cell !== undefined){

        let address = String.fromCharCode("A".charCodeAt(0) + result)  + "1";
        console.log(address);
        result++;
        cell = getCellValue( address, worksheet);
    }

    return result;
}
*/

function getSheet() {
    //gets the sheet with the data on it

    let workbook = XLSX.readFile('data/shippingDetails.xlsx');

    let first_sheet_name = workbook.SheetNames[0];
    return workbook.Sheets[first_sheet_name];
}

function writeReport(worksheet, startRow, columns, func, path){
    //iterates over each row in the worksheet gives it to the function (func)
    //if func returns true, it coppies the row into a workbook and saves it
    //columns (parameter) is the title of the columns wanted for func

    let workBook = Utils.makeWorkBook();
    let writeSheet = Utils.makeWorkSheet(workBook, "Report");


    Utils.saveWorkBook(workBook);
}

function InvalidEntry(CellRow, exp, startingRow) {
    //checks to see if the regular expression matches the cell
    //if it does not, it will print it out

    let worksheet = getSheet();

    let rowCounter = startingRow;
    let cellAdress = CellRow + rowCounter;
    let cellValue = Utils.getCellValue(cellAdress, worksheet);

    const regex = RegExp(exp);

    while (cellValue !== undefined) {

        if (!regex.test(cellValue)) {

            let orderNum = Utils.getCellValue("A" + rowCounter, worksheet);
            console.log(orderNum + " " + cellValue);
        }

        rowCounter++;
        cellValue = Utils.getCellValue(CellRow + rowCounter, worksheet);
    }
}
