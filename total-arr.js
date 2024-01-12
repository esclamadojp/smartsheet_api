const ssid = process.argv.slice(2);
const sheet_id = ssid[0];
const sheet_group = ssid[1];

console.log('Starting');

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

// Initialize client SDK. Uses API token from environment variable "SMARTSHEET_ACCESS_TOKEN"
var ss = client.createClient({ logLevel: 'info', accessToken: 'Q3fjIhpkqVBy1xAaig1UciOZpgKldQdNoBIhQ', });

ss.sheets.getSheet({ id: sheet_id })
        .then(function(sheet) {

            sheet.columns.forEach(function(column) {
                columnMap[column.title] = column.id;
            });

            const grouped = {};
            const totalArr = {};

            sheet.rows.forEach(function(row) {

                var groupCell = getCellByColumnName(row, sheet_group);
                var arrCell = getCellByColumnName(row, 'arr');

                if (!grouped[groupCell.displayValue]) {
                    grouped[groupCell.displayValue] = [];
                    totalArr[groupCell.displayValue] =  0;
                }
                
                grouped[groupCell.displayValue].push({
                    id: row.id
                });

                totalArr[groupCell.displayValue] +=  parseFloat(arrCell.displayValue);
            });

            console.log(grouped);

            for (const groupName in grouped) {
              if (grouped.hasOwnProperty(groupName)) {
                const lastId = grouped[groupName][grouped[groupName].length - 1].id;
                console.log("Last ID: "+ lastId);
                var row = [
                    {
                        "parentId": lastId, "toBottom": true,
                        "cells": [
                        {
                            "columnId": columnMap["id"],
                            "value": "Total ARR"
                        },
                        {
                            "columnId": columnMap["arr"],
                            "value":totalArr[groupName]
                        }
                      ]
                    }
                  ];
                  
                  // Set options
                var options = {
                    sheetId: sheet.id,
                    body: row
                };
                
                // // Add rows to sheet
                ss.sheets.addRows(options)
                    .then(function(newRows) {
                        console.log(newRows);
                    })
                    .catch(function(error) {
                        console.log(error);
                    });
              }
            }

                
            console.log("Done");
        })
        .catch(function(error) {
            console.log(error);
        });


