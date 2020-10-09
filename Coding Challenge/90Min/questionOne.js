
if (typeof require !== 'undefined') XLSX = require('xlsx');

//error reports
console.log("Invalid Email Addresses")
InvalidEntry("I","@.+(\.com|\.edu|\.net|\.org)$",2);
console.log("")
console.log("Invalid Phone Numbers")
InvalidEntry("H","\\d\\d\\d-\\d\\d\\d-\\d\\d\\d\\d",2);

//1
withinADay();

//2

//3

function withinADay(){
    //checks to see if the orderDate and ShipDate are less than 24 hours apart
    let dateData = getSheetData(2,["A","L","M"]);

    //make excel sheet

    //go though data
    for(let i = 0; i < dateData.length; i++){

        //I cant find a date parser!!!!!
        //so i guess im doing it by hand!
        let orderDate = (dateData[i])["L"].match("(\\d\\dT)")[0];
        orderDate = orderDate.substring(0,orderDate.length - 1);
        let ordertime = (dateData[i])["L"].match("\\d\\d:\\d\\d:\\d\\d")[0];

        let sendDate = (dateData[i])["M"].match("(\\d\\dT)")[0];
        sendDate = sendDate.substring(0,sendDate.length - 1);
        let sendTIme = (dateData[i])["M"].match("\\d\\d:\\d\\d:\\d\\d")[0];

        console.log(dateData[i]["A"] + " " + parseInt(orderDate) + " " + parseInt(sendDate));
        if(parseInt(sendDate) + 1 >= parseInt(orderDate)){
            console.log(dateData[i]["A"]);
        }
    }
}

function getSheetData(startRow, columns){
    //returns an array of objects that holds the data rom the columns
    //colums should be an array of column you want data from

    let result = [];

    let worksheet = getSheet();

    let rowCounter = startRow;
    let cellAdress = "A" + rowCounter;
    let cellValue = getCellValue(cellAdress,worksheet);

    while(cellValue !== undefined){

        result.push(getRowData(worksheet,rowCounter,columns));

        rowCounter++;
        cellAdress = "A" + rowCounter;
        cellValue = getCellValue(cellAdress,worksheet);
    }

    return result;
}

function getRowData(worksheet, row, columns){
    //gets the data from a sigle row
    result = {};
    for(let i = 0; i < columns.length; i++){
        result[columns[i]] = getCellValue(columns[i] + row, worksheet);
    }
    return result;
}

function getSheet() {
    //gets the sheet with the data on it

    let workbook = XLSX.readFile('data/shippingDetails.xlsx');

    let first_sheet_name = workbook.SheetNames[0];
    return workbook.Sheets[first_sheet_name];
}

function getCellValue(cellAdress, worksheet) {
    //get the value of a given cell
    let desired_cell = worksheet[cellAdress];
    return (desired_cell ? desired_cell.v : undefined);
}

function InvalidEntry(CellRow, exp, startingRow){
    //checks to see if the regular expression matches the cell
    //if it does not, it will print it out
    
    let worksheet = getSheet();

    let rowCounter = startingRow;
    let cellAdress = CellRow + rowCounter;
    let cellValue = getCellValue(cellAdress,worksheet);

    const regex = RegExp(exp);

    while (cellValue !== undefined) {

        if (!regex.test(cellValue)) {

            let orderNum = getCellValue("A" + rowCounter, worksheet);
            console.log(orderNum + " " + cellValue);
        }

        rowCounter++;
        cellValue = getCellValue(CellRow + rowCounter, worksheet);
    }
}
