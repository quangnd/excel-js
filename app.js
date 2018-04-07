const Excel = require('exceljs');

const workbook = new Excel.Workbook();
const fileName = "test.xlsx"

console.log(`Processing ${fileName}...`)
workbook.xlsx.readFile(fileName)
	.then(function () {
		const worksheet = workbook.getWorksheet('Profile and Delivery');
		const deliveryCol = worksheet.getColumn('E');
		const variableCol = worksheet.getColumn('B');
		let cellValues = [];

		worksheet.eachRow(function (row, rowNumber) {
			let variableColValue = row.getCell(2).value;
			let profileTagColValue = row.getCell(3).value;
			if (variableColValue === 'Conscientiousness' && profileTagColValue === `"Hard Worker"`) {
				cellValues.push(row.getCell(5).value)
			}
			//add more if
		});
		console.log(cellValues)
	});

