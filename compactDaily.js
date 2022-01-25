const readXlsxFile = require('read-excel-file/node');
const writeXlsxFile = require('write-excel-file/node');
let data = []
let columnTitle = []
let schema = [];
let dataDate = [];

let file_name;
const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
})

readline.question(`Input your file name (without .xlsx): `, name => {
    file_name= name

    readXlsxFile('./'+file_name+'.xlsx').then((rows) => {
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
        let records = [];
        dataDate.forEach(date => {
            let recordRow = [date]
            Object.entries(data[date]).forEach(entry => {
                const [key, value] = entry;
                recordRow[schema.indexOf(key)] = value
            });
            records.push(recordRow)
        })
        let formatToExcel = []
        let columntemp  = []
        schema.forEach(column => {

            columntemp.push({value: column})
        })
        formatToExcel.push(columntemp)
        records.forEach((row, index) => {
            let temp = []
            for(let i = 0; i < row.length; i++){
                temp.push({value: row[i]})
            }
            formatToExcel.push(temp)
        })

        console.log(formatToExcel)
        writeXlsxFile(formatToExcel, {
            filePath: './'+file_name+'_out.xlsx'
        }).then(res => {
            console.log(res)
        })
        // csvWriter.writeRecords(records)       // returns a promise
        //     .then(() => {
        //         console.log('...Done');
        //     });

    })
    readline.close()
})
