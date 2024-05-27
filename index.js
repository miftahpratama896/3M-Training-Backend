const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const sql = require("mssql");

// Database configuration
const config = {
  user: "sa",
  password: "Pai2015",
  server: "172.16.200.28",
  database: "JX2MES",
  options: {
    encrypt: false, // for Azure
    requestTimeout: 60000,
  },
};

const app = express();
const filePath = path.join(__dirname, 'data', '3M_Training.xlsx'); // Use relative path

app.use(express.json());
app.use(cors()); // Enable CORS

// Helper function to format dates
const formatDate = (date) => {
  // Add your date formatting logic here
  return date;
};

// Function to read and process Excel data
const fetchData = (filePath, sheetName) => {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { raw: false });
  return jsonData;
};

app.get('/api/employee', async (req, res) => {
  try {
    // Connect to the database
    await sql.connect(config);

    // Get the filter ID from the query parameters
    const filterID = req.query.id;

    // Build the SQL query dynamically based on the filter ID
    let query = `SELECT 
        [ID],
        MAX([WORK_DATE]) AS [WORK_DATE],
        MAX([NAME]) AS [NAME],
        MAX([SEXX]) AS [SEXX],
        MAX([ENDT]) AS [ENDT],
        MAX([DEPT]) AS [DEPT],
        MAX([MID DEPT]) AS [MID DEPT],
        MAX([SUB DEPT]) AS [SUB DEPT]
      FROM [JX2MES].[dbo].[TR_PRODUCT_PERSONNEL_DETIAL_LIST]`;

    if (filterID) {
      query += ` WHERE [ID] = ${filterID}`;
    }

    query += ` GROUP BY [ID]`;

    // Execute the SQL query
    const result = await sql.query(query);

    // Send the data as JSON response
    res.json(result.recordset);
  } catch (err) {
    // Handle errors
    console.error("Error occurred:", err);
    res.status(500).send("Internal Server Error");
  } 
  
});


// Endpoint to get 'Training Input' data
app.get('/api/training', (req, res) => {
  const sheetName = 'Training Input';
  const jsonData = fetchData(filePath, sheetName);
  const filteredData = jsonData.map(row => ({
    ID: row.ID,
    FACTORY: row.FACTORY,
    LINE: row.LINE,
    DATE: formatDate(row.DATE),
    'TRAINING PLAN': row['TRAINING PLAN'],
    'TRAINING ACTUAL': Number(row['TRAINING ACTUAL']) || 0
  }));
  res.json(filteredData);
});

app.get('/api/training-process', (req, res) => {
  const sheetName = 'Training Process';
  const jsonData = fetchData(filePath, sheetName);
  const filteredData = jsonData.map(row => ({
    TONGUE: row.TONGUE,
    VAMP: row.VAMP,
    'U-THROAT': row['U-THROAT'],
    QUARTER: row.QUARTER,
    COLLAR: row.COLLAR,
    'EYESTAY/EYELET': row['EYESTAY/EYELET'],
    'FOXING/MUDGUARD': row['FOXING/MUDGUARD'],
    SWOOSH: row.SWOOSH,
    'BACKTAB/BACKSTAY/PULL TAB/BOOTIE': row['BACKTAB/BACKSTAY/PULL TAB/BOOTIE'],
    COLLAR: row.COLLAR,
    HEEL: row.HEEL,
    'COUNTER/HAMMERING/TRIMMING/PUNCHING/HABONG/CLEAN/LACING/SIZE LABEL/MARKING': row['COUNTER/HAMMERING/TRIMMING/PUNCHING/HABONG/CLEAN/LACING/SIZE LABEL/MARKING'],
    'TONAL/ZIG ZAG': row['TONAL/ZIG ZAG'],
    TIP: row.TIP
  }));
  res.json(filteredData);
});

// Endpoint to get 'Validation - Matrix Input' data
app.get('/api/validation', (req, res) => {
  const sheetName = 'Validation - Matrix Input';
  const jsonData = fetchData(filePath, sheetName);
  const filteredData = jsonData.map(row => ({
    ID: row.ID,
    FACTORY: row.FACTORY,
    LINE: row.LINE,
    DATE: formatDate(row.DATE),
    NIK: row.NIK,
    NAME: row.NAME,
    'MAIN PROCESS': row['MAIN PROCESS'],
    PROCESS: row.PROCESS,
    'VALIDATION STATUS': row['VALIDATION STATUS'],
    SCORE: row.SCORE
  }));
  res.json(filteredData);
});

// Endpoint to save data to 'Training Input - Line' sheet
app.post('/saveData', (req, res) => {
  try {
    const data = req.body.data;

    // Read existing workbook
    const workbook = XLSX.readFile(filePath);
    const sheetNameTrainingInput = 'Training Input';
    const sheetNameTrainingInputLine = 'Training Input - Line';

    // Load sheets
    const worksheetTrainingInput = workbook.Sheets[sheetNameTrainingInput];
    let worksheetTrainingInputLine = workbook.Sheets[sheetNameTrainingInputLine];

    if (!worksheetTrainingInputLine) {
      // Create the worksheet if it doesn't exist
      worksheetTrainingInputLine = {};
      workbook.Sheets[sheetNameTrainingInputLine] = worksheetTrainingInputLine;
      workbook.SheetNames.push(sheetNameTrainingInputLine);
    }

    // Find the next available row in 'Training Input - Line' sheet
    let rowIndex = 5;
    while (worksheetTrainingInputLine[`A${rowIndex}`] && worksheetTrainingInputLine[`A${rowIndex}`].v !== undefined) {
      rowIndex++;
    }

    // Write data to 'Training Input - Line' sheet
    data.forEach((row) => {
      worksheetTrainingInputLine[`A${rowIndex}`] = { t: 'n', v: row.ID };
      worksheetTrainingInputLine[`B${rowIndex}`] = { t: 's', v: row.FACTORY };
      worksheetTrainingInputLine[`C${rowIndex}`] = { t: 'n', v: row.LINE };
      worksheetTrainingInputLine[`D${rowIndex}`] = { t: 's', v: row.DATE };
      worksheetTrainingInputLine[`E${rowIndex}`] = { t: 'n', v: Number(row['ACTUAL TRAINING']) };
      rowIndex++;
    });

    // Update the range to include new data in 'Training Input - Line' sheet
    if (!worksheetTrainingInputLine['!ref']) {
      worksheetTrainingInputLine['!ref'] = `A5:E${rowIndex - 1}`;
    } else {
      const ref = XLSX.utils.decode_range(worksheetTrainingInputLine['!ref']);
      ref.e.r = rowIndex - 1;
      worksheetTrainingInputLine['!ref'] = XLSX.utils.encode_range(ref);
    }

    // Update 'TRAINING ACTUAL' in 'Training Input' sheet
    const jsonDataTrainingInput = XLSX.utils.sheet_to_json(worksheetTrainingInput, { raw: false });

    data.forEach((row) => {
      const matchingRow = jsonDataTrainingInput.find(r => r.LINE === row.LINE && r.FACTORY === row.FACTORY);
      if (matchingRow) {
        matchingRow['TRAINING ACTUAL'] = (Number(matchingRow['TRAINING ACTUAL']) || 0) + Number(row['ACTUAL TRAINING']);
        matchingRow['TRAINING PLAN'] = row.MP; // Update TRAINING PLAN with MP
      }
    });

    // Convert updated JSON data back to worksheet
    const newWorksheetTrainingInput = XLSX.utils.json_to_sheet(jsonDataTrainingInput, { header: Object.keys(jsonDataTrainingInput[0]) });
    workbook.Sheets[sheetNameTrainingInput] = newWorksheetTrainingInput;

    // Write back to the file
    XLSX.writeFile(workbook, filePath);

    res.status(200).json({ message: 'Data saved successfully.' });
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Internal server error.' });
  }
});

const { v4: uuidv4 } = require('uuid'); // Import UUID library

app.post('/saveDataValidation', (req, res) => {
  try {
    const data = req.body.data;

    // Read existing workbook
    const workbook = XLSX.readFile(filePath);
    const sheetName = 'Validation - Matrix Input';

    // Load sheet
    let worksheet = workbook.Sheets[sheetName];

    if (!worksheet) {
      // Create the worksheet if it doesn't exist
      worksheet = {};
      workbook.Sheets[sheetName] = worksheet;
      workbook.SheetNames.push(sheetName);
    }

    // Find the next available row in 'Validation - Matrix Input' sheet
    let rowIndex = 5;
    while (worksheet[`A${rowIndex}`] && worksheet[`A${rowIndex}`].v !== undefined) {
      rowIndex++;
    }

    // Read existing data in column NIK
    const existingNIKs = new Set();
    let existingRowIndex = 5;
    while (worksheet[`E${existingRowIndex}`] && worksheet[`E${existingRowIndex}`].v !== undefined) {
      existingNIKs.add(worksheet[`E${existingRowIndex}`].v.toString());
      existingRowIndex++;
    }

    // Check if any of the new NIKs already exist in the table
    const newDataNIKs = data.map(row => row.NIK.toString());
    const duplicateNIK = newDataNIKs.find(nik => existingNIKs.has(nik));

    if (duplicateNIK) {
      return res.status(400).json({ error: 'DATA SUDAH ADA' });
    }

    // Write data to 'Validation - Matrix Input' sheet
    data.forEach((row) => {
      const id = uuidv4(); // Generate unique ID
      const nik = row.NIK.toString(); // Convert NIK to string
      worksheet[`A${rowIndex}`] = { t: 's', v: id }; // Use the generated ID
      worksheet[`B${rowIndex}`] = { t: 's', v: row.FACTORY };
      worksheet[`C${rowIndex}`] = { t: 'n', v: row.LINE };
      worksheet[`D${rowIndex}`] = { t: 's', v: row.DATE };
      worksheet[`E${rowIndex}`] = { t: 'n', v: nik };
      worksheet[`F${rowIndex}`] = { t: 's', v: row.NAME };
      worksheet[`G${rowIndex}`] = { t: 's', v: row['MAIN PROCESS'] };
      worksheet[`H${rowIndex}`] = { t: 's', v: row.PROCESS };
      worksheet[`I${rowIndex}`] = { t: 's', v: row['VALIDATION STATUS'] };
      worksheet[`J${rowIndex}`] = { t: 'n', v: row.SCORE };
      rowIndex++;
    });

    // Update the range to include new data in 'Validation - Matrix Input' sheet
    if (!worksheet['!ref']) {
      worksheet['!ref'] = `A5:J${rowIndex - 1}`;
    } else {
      const ref = XLSX.utils.decode_range(worksheet['!ref']);
      ref.e.r = rowIndex - 1;
      worksheet['!ref'] = XLSX.utils.encode_range(ref);
    }

    // Write back to the file
    XLSX.writeFile(workbook, filePath);

    res.status(200).json({ message: 'Data saved successfully to Validation - Matrix Input sheet.' });
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Internal server error.' });
  }
});

// Endpoint to export data to Excel
app.get('/api/exportExcel', (req, res) => {
  try {
    // Fetch data from each sheet
    const sheetNames = ['Training Input', 'Training Input - Line', 'Validation - Matrix Input', 'Training Process']; // List of sheet names
    const workSheets = sheetNames.map(sheetName => ({
      name: sheetName,
      data: fetchData(filePath, sheetName)
    }));

    // Create a new workbook
    const workbook = XLSX.utils.book_new();

    // Add each worksheet to the workbook
    workSheets.forEach(workSheet => {
      const worksheet = XLSX.utils.json_to_sheet(workSheet.data);
      XLSX.utils.book_append_sheet(workbook, worksheet, workSheet.name);
    });

    // Write the workbook to a file
    const excelFilePath = path.join(__dirname, 'data', `training_data_${Date.now()}.xlsx`);
    XLSX.writeFile(workbook, excelFilePath);

    res.download(excelFilePath); // Respond with the file for download
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Internal server error.' });
  }
});


app.listen(1000, () => {
  console.log('Server is running on port 1000');
});
