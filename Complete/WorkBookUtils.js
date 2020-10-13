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
                result++;
                cell = this.getCellValue( address, worksheet);
            }

            if(cell === undefined){
                return undefined;
            } else if(result === 0 ){
                return String.fromCharCode("A".charCodeAt(0) + result);
            } else {
                return String.fromCharCode("A".charCodeAt(0) + result - 1); 
            }
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

        makeWorkBook : function() {
            /* create a new blank workbook */
            return XLSX.utils.book_new();
        },
        
        generateSheet : function (data) {
            //makes a worksheet and adds it to the workbook
            return XLSX.utils.aoa_to_sheet(data);;
        },

        addWorkSheet : function (workBook, workSheet, workSheetName){
            XLSX.utils.book_append_sheet(workBook, workSheet, workSheetName);
        },
        
        saveWorkBook : function(workBook, path) {
            //just save the workbook to the path
            XLSX.writeFile(workBook, path);
        }
    }
}