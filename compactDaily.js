const readXlsxFile = require('read-excel-file/node');
const writeXlsxFile = require('write-excel-file/node');
let data = []
let columnTitle = []
let dataexcl = [];    
let schema = [];
let dataDate = [];
readXlsxFile('./input.xlsx').then((rows) => {
    console.log(rows);
    var sortedArray = rows.sort(function(a, b) { 
        return a[1] - a[1];
    })

    rows.forEach(function(element, index) {
        sortedArray[index][1] = element[1].toLowerCase();
        sortedArray[index][5] = String(element[5]);
    })
    sortedArray.forEach((element, index) => {
        if(index !== 0){
            let date = element[2].toLocaleDateString()
            let email = element[1];
            data[date] = {...data[date], [email]: element[5]}
            if(dataDate.indexOf(date) == -1){
                dataDate.push(date);
            }
            columnTitle.push(element[1])
        }
    });
    schema.push('Tanggal/nama')
    columnTitle.forEach((element, index) => {
        if(schema[schema.length - 1] != element){
            schema.push(element)
        }
    })
    const createCsvWriter = require('csv-writer').createArrayCsvWriter;
    const csvWriter = createCsvWriter({
        header: schema,
        path: './output.csv'
    });
    let records = [];
    dataDate.forEach(date => {
        let recordRow = [date]
        Object.entries(data[date]).forEach(entry => {
            const [key, value] = entry;
            recordRow[schema.indexOf(key)] = value
          });
        records.push(recordRow)
    })
    csvWriter.writeRecords(records)       // returns a promise
        .then(() => {
            console.log('...Done');
        });

})