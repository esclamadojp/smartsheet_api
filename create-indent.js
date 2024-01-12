const sheet_group = process.argv.slice(2);

console.log('Starting...Grouping by ' + sheet_group);


// If not found, install SDK package with command line: npm install smartsheet
var client = require('smartsheet');

// The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
var columnMap = {};

// Helper function to find cell in a row
function getCellByColumnName(row, columnName) {
    var columnId = columnMap[columnName];
    return row.cells.find(function(c) {
        return (c.columnId == columnId);
    });
}

function indentCells(grouped, sheetId){
   
    for (const groupName in grouped) {
        if (grouped.hasOwnProperty(groupName)) {

            grouped[groupName].shift(); 
            
            console.log(`Group: ${groupName}`);
            console.log(grouped[groupName]);
            
            var updateIndentedRowArgs = {
                body: grouped[groupName],
                sheetId: sheetId
            };

            ss.sheets.updateRow(updateIndentedRowArgs)
                .then(function(result) {
                    console.log("Updated succeded");
                })
                    .catch(function(error) {
                    console.log(error);
                });

            
            
            
        }
      }
}

// Initialize client SDK. Uses API token from environment variable "SMARTSHEET_ACCESS_TOKEN"
var ss = client.createClient({ logLevel: 'info', accessToken: 'Q3fjIhpkqVBy1xAaig1UciOZpgKldQdNoBIhQ', });

var options = {
    path: "data.xlsx",
    fileName: "data.xlsx",
    queryParameters: {
        sheetName: "MySheetBy"+sheet_group,
        headerRowIndex: 0
    }
};

ss.sheets.importXlsxSheet(options)
    .then(function(result) {
        console.log("Created sheet '" + result.result.id + "' from excel file");
        
        //sort group
        ss.sheets.getSheet({ id: result.result.id })
        .then(function(sheet) {
            console.log("Loaded " + sheet.rows.length + " rows from sheet '" + sheet.name + "'");

            // Build column map for later reference - converts column name to column id
            sheet.columns.forEach(function(column) {
                columnMap[column.title] = column.id;
            });

            // sort by group
            var sortBody = {
                sortCriteria: [
                    {
                    columnId: columnMap[sheet_group],
                    direction: "ASCENDING"
                    }
                ]
                };

            ss.sheets.sortRowsInSheet({sheetId: sheet.id, body: sortBody})
            .then((result) => {
                console.log("sort succeded");
            })
            .catch((error) => {
                console.log(error);
            });
        })
        .catch(function(error) {
            console.log(error);
        });

        // group and indent
        ss.sheets.getSheet({ id: result.result.id })
        .then(function(sheet) {

            const grouped = {};

            sheet.rows.forEach(function(row) {
                // console.log("ROW: "+row.id);
                var groupCell = getCellByColumnName(row, sheet_group);
                if (!grouped[groupCell.displayValue]) {
                    grouped[groupCell.displayValue] = [];
                }
                
                if (grouped[groupCell.displayValue].length == 0)
                    grouped[groupCell.displayValue].push({
                        id: row.id
                    });
                else
                {
                    grouped[groupCell.displayValue].push({
                        id: row.id,
                        indent: 1
                    });
                }

            });

            indentCells(grouped, sheet.id);
           
            console.log("Indent done...");
        })
        .catch(function(error) {
            console.log(error);
        });


    })
    .catch(function(error) {
        console.log(error);
    });


