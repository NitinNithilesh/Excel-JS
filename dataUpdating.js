//Read a file
let Excel = require('exceljs')
let workbook = new Excel.Workbook();
workbook.xlsx.readFile("test.xlsx").then(function () {

	//Get sheet by Name
	let worksheet = workbook.getWorksheet('Sheet1');

	// Initialize the row
	let row;

	//Update a cell in for loop
	for (let i = 2; i <= 285; i++) {
		row = worksheet.getRow(i);
		let originalAmount = row.getCell(6);
		let gstPercent = row.getCell(9);
		let gstAmount, netAmount;
		gstAmount = (Number(originalAmount) * Number(gstPercent)) / 100;
		netAmount = Number(originalAmount) + gstAmount;
		row.getCell(11).value = netAmount;
	}

	// Commit the row after updating values in it
	row.commit();

	//Save the workbook
	return workbook.xlsx.writeFile("test.xlsx");

});
