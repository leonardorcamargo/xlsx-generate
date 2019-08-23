const officegen = require('officegen')
const fs = require('fs')

// Create an empty Excel object:
let xlsx = officegen('xlsx')

// Officegen calling this function after finishing to generate the xlsx document:
xlsx.on('finalize', function(written) {
  console.log(
    'Finish to create a Microsoft Excel document.'
  )
})

// Officegen calling this function to report errors:
xlsx.on('error', function(err) {
  console.log(err)
})

let sheet = xlsx.makeNewSheet()
let sheet2 = xlsx.makeNewSheet()
sheet.name = 'Officegen Excel'
sheet2.name = 'Teste'

// Add data using setCell:

sheet.setCell('E7', 42)
sheet.setCell('I1', -3)
sheet.setCell('I2', 3.141592653589)
sheet.setCell('G102', 'Hello World!')

// The direct option - two-dimensional array:

sheet.data[0] = []
sheet.data[0][0] = 1
sheet.data[1] = []
sheet.data[1][3] = 'some'
sheet.data[1][4] = 'data'
sheet.data[1][5] = 'goes'
sheet.data[1][6] = 'here'
sheet.data[2] = []
sheet.data[2][5] = 'more text'
sheet.data[2][6] = 900
sheet.data[6] = []
sheet.data[6][2] = 1972

// Let's generate the Excel document into a file:

let out = fs.createWriteStream('example.xlsx')

out.on('error', function(err) {
  console.log(err)
})

// Async call to generate the output file:
xlsx.generate(out)