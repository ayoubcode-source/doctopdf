const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const libre = require('libreoffice-convert');

const app = express();
const port = process.env.PORT || 3000; // Use Railway or other platform's dynamic port

// Middleware to parse form data
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Serve static files (e.g., CSS, JS)
app.use(express.static(path.join(__dirname, '../public'))); // Adjust path for static files

// Serve the HTML form
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../public/index.html')); // Path to your static HTML file
});

// Handle form submission and PDF generation
app.post('/generate-pdf', (req, res) => {
    console.log('Received a POST request to /generate-pdf');
    const templatePath = path.join(__dirname, '../public/اشهاد الشهود.docx'); // Adjust path to your Word template

    // Read the DOCX template
    fs.readFile(templatePath, 'binary', (err, content) => {
        if (err) {
            console.error('Error reading template:', err);
            return res.status(500).send('Error reading template: ' + err.message);
        }

        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '[[', end: ']]' } // Custom delimiters for placeholders
        });

        // Set data to replace placeholders in the document
        doc.setData({
            witness1_name: req.body.witness1_name || '',
            witness1_id: req.body.witness1_id || '',
            witness1_address: req.body.witness1_address || '',
            witness2_name: req.body.witness2_name || '',
            witness2_id: req.body.witness2_id || '',
            witness2_address: req.body.witness2_address || '',
            declaration_content: req.body.declaration_content || '',
            date: req.body.date || '',
        });

        try {
            doc.render();
        } catch (error) {
            console.error('Error rendering document:', error);
            return res.status(500).send('Error rendering document: ' + error.message);
        }

        // Generate the DOCX as a buffer
        const docxBuffer = doc.getZip().generate({ type: 'nodebuffer' });

        // Convert the DOCX buffer to PDF
        libre.convert(docxBuffer, '.pdf', undefined, (err, pdfBuffer) => {
            if (err) {
                console.error('Error generating PDF:', err);
                return res.status(500).send('Error generating PDF: ' + err.message);
            }

            // Set the headers to display the PDF in the browser
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'inline; filename=document.pdf'); // 'inline' shows the PDF in the browser

            // Send the PDF buffer to the client
            res.send(pdfBuffer);
        });
    });
});

// Start the Express server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
