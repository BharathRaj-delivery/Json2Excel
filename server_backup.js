const express = require('express');
const multer = require('multer');
const fs = require('fs');
const JSONStream = require('JSONStream');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = 3001;

app.get('/favicon.ico', (req, res) => res.status(204).end());

// Set up file storage
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public')); // Serve frontend files
app.use(express.json());

var outputFileName='';

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// API: Upload JSON and Excel format files
app.post('/upload', upload.fields([{ name: 'jsonFile' }, { name: 'excelFile' }]), async (req, res) => {
    try {
        console.log("🚀 Files uploaded, processing started...");

        // Get file paths
        const jsonFilePath = req.files['jsonFile'][0].path;
        const excelFilePath = req.files['excelFile'][0].path;
const jsonFileName = req.files['jsonFile'][0].originalname; // Get the original JSON file name
outputFileName = jsonFileName.replace(path.extname(jsonFileName), '.xlsx'); // Change extension to .xlsx
const outputFilePath = path.join(__dirname, 'public', outputFileName); // Set output file path

        // Read Excel format file
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const worksheet = workbook.worksheets[0];

        // Extract output headers and JSON property names
        let outputKeys = worksheet.getRow(1).values.slice(1).map(val => val?.toString().trim()).filter(Boolean);
        let jsonKeys = worksheet.getRow(2).values.slice(1).map(val => val?.toString().trim().replace(/\[|\]/g, '')).filter(Boolean);

        if (outputKeys.length !== jsonKeys.length) {
            throw new Error("Mismatch between header keys and JSON property names.");
        }

        console.log("✅ Extracted Headers & JSON Mapping...");

        // Create output workbook
        const outputWorkbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: outputFilePath });
        const outputWorksheet = outputWorkbook.addWorksheet('Applicants');

        // Write headers
        outputWorksheet.addRow(outputKeys).commit();

        let rowCount = 0;
        let totalRows = 0;

        // First, count total rows efficiently
        await new Promise((resolve, reject) => {
            const countStream = fs.createReadStream(jsonFilePath, 'utf-8')
                .pipe(JSONStream.parse('*'));
            countStream.on('data', () => totalRows++);
            countStream.on('end', resolve);
            countStream.on('error', reject);
        });

        console.log(`📊 Total rows to process: ${totalRows}`);

        // Process JSON data with batch writing
        let batch = [];
        const batchSize = 100; // Process in batches of 100
        let progress = 0; // Track progress

        await new Promise((resolve, reject) => {
            const jsonStream = fs.createReadStream(jsonFilePath, 'utf-8')
                .pipe(JSONStream.parse('*'));

            jsonStream.on('data', async (data) => {
                if (typeof data !== 'object' || data === null) {
                    console.error("❌ Invalid JSON data format:", data);
                    return;
                }

                // Extract row data
                let rowData = outputKeys.map((_, i) => data[jsonKeys[i]]?.toString() ?? '');
                batch.push(rowData);
                rowCount++;

                // Process batch when queue reaches batchSize
                if (batch.length >= batchSize) {
                    batch.forEach(row => outputWorksheet.addRow(row).commit());
                    batch = []; // Clear batch after processing
                }

                // Update progress
                progress = ((rowCount / totalRows) * 100).toFixed(2);
                console.log(`✅ Row ${rowCount}/${totalRows} written (${progress}% completed)...`);
            });

            jsonStream.on('end', async () => {
                try {
                    // Process remaining rows if batch is not empty
                    if (batch.length > 0) {
                        batch.forEach(row => outputWorksheet.addRow(row).commit());
                        console.log(`✅ Processed remaining ${batch.length} rows.`);
                    }

                    // Ensure all worksheet writes are completed before closing
                    await outputWorksheet.commit();
                    console.log("✅ Worksheet committed successfully!");

                    // Finalize workbook
                    await outputWorkbook.commit();
                    outputWorkbook.stream.end(); // 🔥 Forcefully close the stream
                    console.log("🎉 Workbook finalized and closed!");

                    // Delete uploaded files (JSON & Excel Format) after processing
                    fs.unlinkSync(jsonFilePath);
                    fs.unlinkSync(excelFilePath);
                    console.log("🗑️ Deleted temporary JSON & Excel format files!");

                    // Send success response
                    res.json({ success: true, downloadUrl: '/download', progress });

                    resolve();
                } catch (error) {
                    console.error("❌ Error finalizing Excel file:", error);
                    res.status(500).json({ error: "Error finalizing Excel file." });
                    reject(error);
                }
            });

            jsonStream.on('error', (err) => {
                console.error("❌ Error reading JSON file:", err);
                res.status(500).json({ error: "Error processing JSON file." });
                reject(err);
            });
        });

    } catch (error) {
        console.error("❌ Error:", error);
        res.status(500).json({ error: "Internal server error." });
    }
});

// API: Download the processed file and delete it afterward
app.get('/download', (req, res) => {
    const filePath = path.join(__dirname, 'public', outputFileName);

    if (fs.existsSync(filePath)) {
        res.download(filePath, (err) => {
            if (!err) {
                // Delete applicants.xlsx after download
                fs.unlinkSync(filePath);
                console.log("🗑️ Deleted applicants.xlsx after download!");
            } else {
                console.error("❌ Error downloading file:", err);
            }
        });
    } else {
        res.status(404).json({ error: "❌ File not found. Please upload and process data first." });
    }
});

app.listen(port, () => {
    console.log(`🚀 Server running at http://localhost:${port}`);
});
