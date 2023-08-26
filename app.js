const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

const app = express();
const port = 3000;

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.render('index', { message: '' });
});

app.post('/grades', async (req, res) => {
    const sittingNum = parseInt(req.body.sitting_num);
    const college = req.body.college;

    if (!isNaN(sittingNum)) {
        const excelPath = path.join(__dirname, 'grades', `${college}.xlsx`);

        try {
            const workbook = xlsx.readFile(excelPath);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            const excelData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
            const secondColumnValues = excelData.slice(6).map(row => String(row[1]));

            const matchingRowIndex = secondColumnValues.findIndex(value => value.includes(sittingNum.toString()));

            if (matchingRowIndex !== -1) {
                const matchingRow = excelData[matchingRowIndex + 6]; // Add 6 to get the original row index
                res.render('college', { excelData: excelData, row: matchingRow, collegeName: college });
                return;
            }
        } catch (error) {
            console.error('Error reading Excel file:', error);
        }

        res.render('index', { message: 'Sitting number or college not found' });
    } else {
        res.render('index', { message: 'Sitting number must be a number' });
    }
});

app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});
