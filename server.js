const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const port = 3000;

// Middleware
app.use(express.json());
app.use(cors()); // Allow cross-origin requests from front-end
app.use(express.static(__dirname)); // Serve static files from the current directory

// Excel file path
const excelFilePath = path.join(__dirname, 'kissan_data.xlsx');

// Function to update or create Excel file
function updateExcel(data) {
    let workbook;
    let worksheet;

    // Check if file exists
    if (fs.existsSync(excelFilePath)) {
        workbook = XLSX.readFile(excelFilePath);
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
    } else {
        // Create new workbook and worksheet if file doesn't exist
        workbook = XLSX.utils.book_new();
        worksheet = XLSX.utils.json_to_sheet([]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Farmers');
    }

    // Convert worksheet to JSON to append new data
    let existingData = XLSX.utils.sheet_to_json(worksheet);
    existingData.push(data);

    // Update worksheet with new data
    const newWorksheet = XLSX.utils.json_to_sheet(existingData);
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

    // Write back to file
    XLSX.writeFile(workbook, excelFilePath);
}

// Handle form submission
app.post('/submit', (req, res) => {
    const formData = req.body;

    // Flatten arrays (e.g., challenges, priceInfo) into comma-separated strings
    formData.challenges = formData.challenges ? formData.challenges.join(', ') : '';
    formData.priceInfo = formData.priceInfo ? formData.priceInfo.join(', ') : '';

    // Add timestamp
    formData.timestamp = new Date().toISOString();

    try {
        // Update Excel file
        updateExcel(formData);
        res.status(200).send('Data saved successfully');
    } catch (error) {
        console.error('Error saving data:', error);
        res.status(500).send('Error saving data');
    }
});

// Start server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});