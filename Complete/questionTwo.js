if (typeof require !== 'undefined') XLSX = require('xlsx');
const WorkBookUtils = require('./WorkBookUtils'),
    fs = require('fs');

const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
});

var wbUtils = WorkBookUtils();

main();

function main() {
    //the main loop for the code

    console.log("Please type the number that corresponds with the report wanted:");
    console.log("1. availability \n2. multiple variants \n3. price");

    readline.question('> ', option => {

        switch (option) {
            case "1":
                console.log("availability");
                availability();
                break;
            case "2":
                console.log("multiple variants");
                multiple();
                break;
            case "3":
                console.log("price");
                price();
                break;
            default:
            // does nothing
        }
    });
}

function availability() {
    // generates a worksheet that shows the avaible and unavialble products

    //loads json
    let products = getJsonData();

    //makes arrays of prduct availability
    let available = [];
    let disabled = [];
    products.data.forEach((pdc) => {
        if (pdc.availability === "available") {
            available.push(pdc.sku);

        } else if (pdc.availability === "disabled") {
            disabled.push(pdc.sku);
        }
    });

    //makes data to be saved in workbook
    let workSheetData = [];
    workSheetData.push(["Available", "Unavailable"]);

    //generates data that data array
    let availableLength = available.length;
    let disabledLength = disabled.length;
    for (let i = 0; i < Math.min(availableLength, disabledLength); i++) {
        workSheetData.push([available.shift(), disabled.shift()]);
    }

    //puts the reaminder of the data in the data array
    while (available.length !== 0) {
        workSheetData.push([available.shift(), ""]);
    }
    while (disabled.length !== 0) {
        workSheetData.push(["", disabled.shift()]);
    }

    //saves the data into a workbook
    saveToWorkbook(workSheetData,"reports/availability.xlsx" )   
}

function multiple() {
    //generates a report that lists all the projucts that have multiple variants

    //loads json
    let products = getJsonData();

    //makes arrays of prduct availability
    let workSheetData = [];
    workSheetData.push(["Multiple Variants"]);

    products.data.forEach((pdc) => {
        if (pdc.variants.length !== 0) {
            console.log(pdc.sku + " " + pdc.variants.length);
            workSheetData.push([pdc.sku]);
        }
    });

    //saves the data into a workbook
    saveToWorkbook(workSheetData, "reports/Variants.xlsx");
}

function price() {
    //generates a workbook that contais the sku of all products that are ...
    //either less than or greater than the number the user passed in
    //user also decides if they want greater than or less than

    //ask the user for greater than or less than
    console.log("Please type the number that corresponds with the locic wanted:");
    console.log("1. > (default)(greater than) \n2. < (less than)");

    let func = (a, b) => { return a > b; };

    readline.question('> ', option => {

        if(option === "2"){
            func = (a, b) => { return a < b; };
        }

        //asks the user for an amount
        console.log("Please type the an amount:");

        readline.question('> ', num => {

            //loads json
            let products = getJsonData();

            //makes arrays of prduct that meets requirements 
            let workSheetData = [];
            workSheetData.push(["Filtered by Price"]);

            products.data.forEach((pdc) => {

                if (func(parseFloat(pdc.price), parseFloat(num))) {

                    workSheetData.push([pdc.sku]);
                }
            });

            //saves data to excel
            saveToWorkbook(workSheetData, "reports/price.xlsx")
            readline.close();
        });
    });
}

function getJsonData(){
    //loads json for this program

    let rawdata = fs.readFileSync('./data/products.json');
    return JSON.parse(rawdata);
}

function saveToWorkbook(data, path) {
    //saves the data into a workbook

    let workBook = wbUtils.makeWorkBook();
    let workSheet = wbUtils.generateSheet(workSheetData);
    wbUtils.addWorkSheet(workBook, workSheet);
    wbUtils.saveWorkBook(workBook, path);
}