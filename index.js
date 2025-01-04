const express = require('express');
const xlsx = require('xlsx');
const cors = require('cors');
const bodyParser = require('body-parser');
const fs = require('fs');
const app = express();

app.use(cors());
app.use(bodyParser.json()); // For parsing JSON request bodies

const filePath = 'data.xlsx'; // Path to the Excel file
const sheetName = 'Sheet1'; // Specify your sheet name

// Helper function to save workbook
const saveWorkbook = (workbook) => {
    xlsx.writeFile(workbook, filePath);
};

app.get('/data', (req, res) => {
    try {
        const workbook = xlsx.readFile('data.xlsx'); // Check if file exists
        const sheetName = workbook.SheetNames[0]; // Ensure sheet exists
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]); // Ensure data format is correct
        res.json(data);
    } catch (error) {
        console.error("Error reading Excel file:", error);
        res.status(500).json({ error: "Internal Server Error" });
    }
});


// CREATE: Add new data
app.post('/data', (req, res) => {
    try {
        const newRow = req.body; // New row data from request
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet);

        data.push(newRow); // Add new row
        const updatedSheet = xlsx.utils.json_to_sheet(data);
        workbook.Sheets[sheetName] = updatedSheet;
        saveWorkbook(workbook);

        res.status(201).json({ message: 'Data added successfully', data: newRow });
    } catch (error) {
        res.status(500).json({ error: 'Error adding data' });
    }
});

// UPDATE: Modify existing data
app.put('/data/:id', (req, res) => {
    try {
        const id = parseInt(req.params.id, 10); // Row ID to update
        const updatedData = req.body; // Updated data from request
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet);

        if (id < 0 || id >= data.length) {
            return res.status(404).json({ error: 'Row not found' });
        }

        data[id] = { ...data[id], ...updatedData }; // Update row data
        const updatedSheet = xlsx.utils.json_to_sheet(data);
        workbook.Sheets[sheetName] = updatedSheet;
        saveWorkbook(workbook);

        res.json({ message: 'Data updated successfully', data: data[id] });
    } catch (error) {
        res.status(500).json({ error: 'Error updating data' });
    }
});

// DELETE: Remove data
app.delete('/data/:id', (req, res) => {
    try {
        const id = parseInt(req.params.id, 10); // Row ID to delete
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet);

        if (id < 0 || id >= data.length) {
            return res.status(404).json({ error: 'Row not found' });
        }

        const deletedData = data.splice(id, 1); // Remove row
        const updatedSheet = xlsx.utils.json_to_sheet(data);
        workbook.Sheets[sheetName] = updatedSheet;
        saveWorkbook(workbook);

        res.json({ message: 'Data deleted successfully', data: deletedData });
    } catch (error) {
        res.status(500).json({ error: 'Error deleting data' });
    }
});

// Start the server
app.listen(5000, () => {
    console.log('Server running on http://localhost:5000');
});
