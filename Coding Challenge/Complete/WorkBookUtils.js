if (typeof require !== 'undefined') XLSX = require('xlsx');

module.exports = function Untils(){
    return {
        getTitleColumn: function (worksheet, title){
            //returns the column letter that the header is on
            //if it does not exits, return undefined
            //assume: header is on row 1 and there are no empty spaces
        
            let result = 0;
            let cell = this.getCellValue("A1",worksheet);
        
            while(cell !== title && cell !== undefined){
        
                let address = String.fromCharCode("A".charCodeAt(0) + result)  + "1";
                console.log(address);
                result++;
                cell = this.getCellValue( address, worksheet);
            }
        
            return cell === undefined ? 
                undefined : String.fromCharCode("A".charCodeAt(0) + result - 1);
        },

        getCellValue : function (cellAdress, worksheet) {
            //get the value of a given cell
            let desired_cell = worksheet[cellAdress];
            return (desired_cell ? desired_cell.v : undefined);
        }, 

        getRowData : function(worksheet, row, columns) {
            //returns an object with data from columns in a given row
        
            result = {};
            for (let i = 0; i < columns.length; i++) {
                result[columns[i]] = this.getCellValue(columns[i] + row, worksheet);
            }
            return result;
        },

        copyRow : function(worksheet, row, endLetter) {
            //returns an array of values up to the endLetter of a row
            //assume: endLetter is a single char
        
            let result = [];
            let ending = endLetter.toUppderCase();
        
            for (let i = "A".charCodeAt(0); i < ending.charCodeAt(0); i++) {
                result.push(this.getCellValue[String.fromCharCode(i) + row], worksheet);
            }
        
            return result;
        },

        makeWorkBook : function() {
            /* create a new blank workbook */
            return XLSX.utils.book_new();
        },
        
        makeWorkSheet : function (workbook, sheetName) {
            //makes a worksheet and adds it to the workbook
        
            let worksheet = XLSX.utils.aoa_to_sheet([]);
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
            return worksheet;
        },
        
        saveWorkBook : function(workBook, path) {
            //just save the workbook to the path
            XLSX.writeFile(workBook, path);
        }
    }
}