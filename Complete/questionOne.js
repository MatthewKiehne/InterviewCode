
if (typeof require !== 'undefined') XLSX = require('xlsx');
const WorkBookUtils = require('./WorkBookUtils'),
    fs = require('fs');

const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
});

var Utils = WorkBookUtils();

//error reports
console.log("Invalid Email Addresses")
InvalidEntry("I", "@.+(\.com|\.edu|\.net|\.org)$", 2);
console.log("")
console.log("Invalid Phone Numbers")
InvalidEntry("H", "\\d{3}-\\d{3}-\\d{4}", 2);


//1
var tempSheet = getSheet();
writeExcelReport(tempSheet, 2,
    ["OrderNum", "OrderDate", "ShipDate"],
    withinADay, "reports/LessThanADay.xlsx");


//2
console.log("");
readline.question('Please Enter a minimum amount: ', amount => {

    let func = costsMoreThan(amount);

    writeExcelReport(tempSheet, 2,
        ["OrderNum", "OrderTotal"],
        func, "reports/CostsMoreThan.xlsx");

    readline.close();
});

//3
var today = new Date();
var dd = String(today.getDate()).padStart(2, '0');
var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
var yyyy = today.getFullYear();
let fileName = mm + "" + dd + "" + yyyy + "-ordertotal.txt";

// Write data in 'Output.txt' . 
fs.writeFile('reports/' + fileName, getTotalValue(tempSheet), (err) => {

    // In case of a error throw err. 
    if (err) throw err;
});

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

function writeExcelReport(worksheet, startRow, columns, func, path) {
    //iterates over each row in the worksheet and gives it to the function (func)
    //if func returns true, it coppies the row into a workbook and saves it
    //columns (parameter) is the title of the columns wanted for func

    //holds the error data
    let errorData = [];
    errorData.push(columns);

    let rowCounter = startRow;
    cellValue = Utils.getCellValue("A" + rowCounter, worksheet);

    //gets the columns for the tiles passed in
    let headerColumns = {};
    for (let i = 0; i < columns.length; i++) {
        headerColumns[columns[i]] = Utils.getTitleColumn(worksheet, columns[i]);
    }

    //goes through the data on the page
    while (cellValue !== undefined) {

        let rowObj = {};
        for (let i = 0; i < columns.length; i++) {
            rowObj[columns[i]] = Utils.getCellValue(headerColumns[columns[i]] + rowCounter, worksheet);
        }

        if (func(rowObj)) {

            let rowData = [];
            for (let i = 0; i < columns.length; i++) {
                rowData.push(rowObj[columns[i]]);
            }
            errorData.push(rowData);
        }

        rowCounter++;
        cellValue = Utils.getCellValue("A" + rowCounter, worksheet);
    }

    //saves the workbook
    let workBook = Utils.makeWorkBook();
    let errorSheet = Utils.generateSheet(errorData);
    Utils.addWorkSheet(workBook, errorSheet, "Report");
    Utils.saveWorkBook(workBook, path);
}

function getTotalValue(worksheet) {

    let total = 0;

    let rowCounter = 2;
    cellValue = Utils.getCellValue("A" + rowCounter, worksheet);

    //gets the columns for the tiles passed in
    let column = Utils.getTitleColumn(worksheet, "OrderTotal");

    //goes through the data on the page
    while (cellValue !== undefined) {


        total += parseFloat(Utils.getCellValue(column + rowCounter, worksheet));

        rowCounter++;
        cellValue = Utils.getCellValue("A" + rowCounter, worksheet);
    }

    let split = total.toString().split(".");

    return "Orders: " + (rowCounter - 2) + " Total: $" + addComma(split[0]) + "." + split[1];
}

function commafy(nStr) {
    return nStr.toString().replace("\\B(?=(\\d{3})+(?!\\d))", ',');
}

function addComma(num) {
    if (num === null) return;

    return (
        num
            .toString() // transform the number to string
            .split("") // transform the string to array with every digit becoming an element in the array
            .reverse() // reverse the array so that we can start process the number from the least digit
            .map((digit, index) =>
                index != 0 && index % 3 === 0 ? `${digit},` : digit
            ) // map every digit from the array.
            // If the index is a multiple of 3 and it's not the least digit,
            // that is the place we insert the comma behind.
            .reverse() // reverse back the array so that the digits are sorted in correctly display order
            .join("")
    ); // transform the array back to the string
}

function costsMoreThan(amount) {
    //returns a funciton that remembers the min needed to return true

    let min = amount;

    return function not(data) {
        //checks to see if the cost is over a certain amount
        //requires: "OrderNum" and "OrderTotal"

        return parseFloat(data.OrderTotal) > min;
    }
}

function withinADay(data) {
    //checks to see if the orderDate and ShipDate are less than 24 hours apart
    //requires: "OrderNum", "OrderDate", "ShipDate" in data

    let result = false;

    //parsing date and time for send and order date/time
    let orderDate = (data.OrderDate + "").match("(\\d\\dT)")[0];
    orderDate = orderDate.substring(0, orderDate.length - 1);
    let ordertime = data.OrderDate.match("\\d\\d:\\d\\d:\\d\\d")[0];

    let sendDate = data.ShipDate.match("(\\d\\dT)")[0];
    sendDate = sendDate.substring(0, sendDate.length - 1);
    let sendTime = data.ShipDate.match("\\d\\d:\\d\\d:\\d\\d")[0];

    if (parseInt(orderDate) + 1 === parseInt(sendDate)) {
        //next day

        let orderTimeSplit = ordertime.split(":");
        let sendTimeSplit = sendTime.split(":");

        //checks the hours:minutes:seconds to see if they are less than 24 hours
        if (parseInt(sendTimeSplit[0]) < parseInt(orderTimeSplit[0])) {
            result = true;
        } else if (parseInt(sendTimeSplit[0]) === parseInt(orderTimeSplit[0])) {
            if (parseInt(sendTimeSplit[1]) < parseInt(orderTimeSplit[1])) {
                result = true;
            } else if (parseInt(sendTimeSplit[1]) === parseInt(orderTimeSplit[1])) {
                if (parseInt(sendTimeSplit[2]) <= parseInt(orderTimeSplit[2])) {
                    result = true;
                }
            }
        }

    } else if (parseInt(orderDate) === parseInt(sendDate)) {
        //same day
        result = true;
    }
    return result;
}

function getSheet() {
    //gets the sheet with the data on it

    let workbook = XLSX.readFile('data/shippingDetails.xlsx');

    let first_sheet_name = workbook.SheetNames[0];
    return workbook.Sheets[first_sheet_name];
}