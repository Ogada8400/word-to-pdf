// Import required libraries
const express = require('express');
const bodyParser = require('body-parser');
const app = express();
const path = require('path');
const multer = require('multer');
const docxToPdf = require('docx-pdf');
const fs = require('fs-extra');
const archiver = require('archiver');
const AdmZip = require('adm-zip');
const swaggerUi = require('swagger-ui-express');
const swaggerSpec = require('./swagger');
// Serve Swagger UI
app.use('/api-docs', swaggerUi.serve, swaggerUi.setup(swaggerSpec));

// Define the server port
const port = 3000;

// Serve static files from the 'uploads' directory
app.use(express.static('uploads'));

// Configure multer to handle file uploads with specified storage options
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  },
});
const upload = multer({ storage: storage });

// Use body-parser to parse incoming request data
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

/**
 * @swagger
 * /:
 *   get:
 *     summary: Home Page
 *     description: Serve the home page of the web app.
 *     responses:
 *       '200':
 *         description: Successful response. Returns the home page.
 */

// Define a route for the home page
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});

// Define a route for handling the conversion of DOCX files to PDF
/**
 * @swagger
 * /docxtopdf:
 *   post:
 *     summary: Convert DOCX files to PDF
 *     description: Upload and convert DOCX files to PDF format.
 *     consumes:
 *       - multipart/form-data
 *     parameters:
 *       - in: formData
 *         name: files
 *         description: DOCX file(s) to convert (supports single or multiple files).
 *         required: true
 *         type: array
 *         items:
 *           type: files
 *     responses:
 *       '200':
 *         description: Successful conversion. Returns the ZIP file containing converted PDFs.
 *       '400':
 *         description: Bad request. No files were uploaded or invalid file format.
 *       '500':
 *         description: Internal server error.
 */

// Define a route for handling the conversion of DOCX files to PDF
app.post('/docxtopdf', upload.array('files'), async (req, res) => {
  try {
    // Get the uploaded files from the request
    const files = req.files;

    // Check if files were uploaded
    if (!files || files.length === 0) {
      return res.status(400).send('No files were uploaded.');
    }

    // Determine whether a single ZIP file or individual files were uploaded
    if (files.length === 1 && files[0].mimetype === 'application/zip') {
      await processZipFile(files[0].path, res);
    } else {
      await processIndividualFiles(files, res);
    }
  } catch (error) {
    console.error('Error processing files:', error);
    res.status(500).send('Internal Server Error');
  }
});

// Function to handle the processing of individual DOCX files
async function processIndividualFiles(files, res) {
  try {
    // Set up options for creating a ZIP file to store converted PDFs
    const zipOutputPath = path.join(__dirname, 'uploads', 'converted_files.zip');
    const zipOutput = archiver('zip', { zlib: { level: 9 } });
    const zipOutputStream = fs.createWriteStream(zipOutputPath);

    // Pipe the ZIP output stream
    zipOutput.pipe(zipOutputStream);

    // Counter to track the number of files processed
    let filesProcessed = 0;

    // Event listener for when ZIP output stream finishes
    zipOutputStream.on('finish', () => {
      // Send the ZIP file for download
      res.download(zipOutputPath, 'converted_files.zip', (err) => {
        if (err) {
          console.error('Error sending ZIP file for download:', err);
        }

        // Delete the created ZIP file after download
        cleanupFiles(zipOutputPath);
      });
    });

    // Convert each individual DOCX file and append to the ZIP file
    await Promise.all(files.map(async (file) => {
      if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        try {
          await convertAndAppendToZip(file.path, zipOutput, files.length);
          filesProcessed++;

          // If all files are processed, finalize the ZIP
          if (filesProcessed === files.length) {
            zipOutput.finalize();
          }
        } catch (error) {
          console.error('Error converting and appending to ZIP:', error);
          res.status(500).send('Internal Server Error');
        }
      }
    }));
  } catch (error) {
    console.error('Error processing individual files:', error);
    res.status(500).send('Internal Server Error');
  }
}

// Function to handle the processing of ZIP files
async function processZipFile(zipFilePath, res) {
  try {
    // Create a directory to extract files from the ZIP archive
    const extractPath = path.join(__dirname, 'uploads', 'zip');
    fs.ensureDirSync(extractPath);

    // Extract files from the ZIP archive to the specified directory
    const zip = new AdmZip(zipFilePath);
    zip.extractAllTo(extractPath, true);

    // Get a list of folders extracted from the ZIP archive
    const extractedFolders = fs.readdirSync(extractPath);

    // Log the extracted folders for debugging purposes
    console.log('Extracted Folders:', extractedFolders);

    // Array to store paths of DOCX files found in the ZIP archive
    const docxFiles = [];

    // Iterate through each extracted folder
    extractedFolders.forEach((folder) => {
      const folderPath = path.join(extractPath, folder);

      // Get a list of files within each folder
      const filesInFolder = fs.readdirSync(folderPath);

      // Log the files in each folder for debugging purposes
      console.log(`Files in ${folder}:`, filesInFolder);

      // Filter and map files to get paths of DOCX files
      const docxFilesInFolder = filesInFolder
        .filter((file) => file.endsWith('.docx'))
        .map((file) => path.join(folderPath, file));

      // Concatenate the DOCX files from the current folder to the main array
      docxFiles.push(...docxFilesInFolder);
    });

    // Log all DOCX files found in the ZIP archive for debugging purposes
    console.log('All DOCX Files:', docxFiles);

    // Respond with an error if no DOCX files are found in the ZIP archive
    if (docxFiles.length === 0) {
      return res.status(400).send('No DOCX files found in the ZIP archive.');
    }

    // Set up options for creating a ZIP file to store converted PDFs
    const zipOutputPath = path.join(__dirname, 'uploads', 'converted_files.zip');
    const zipOutput = archiver('zip', { zlib: { level: 9 } });
    const zipOutputStream = fs.createWriteStream(zipOutputPath);

    // Pipe the ZIP output stream
    zipOutput.pipe(zipOutputStream);

    // Convert each DOCX file and append to the ZIP file
    await Promise.all(
      docxFiles.map((docxFilePath) =>
        convertAndAppendToZip(docxFilePath, zipOutput, docxFiles.length)
      )
    );

    // Event listener for when ZIP output stream finishes
    zipOutputStream.on('finish', () => {
      // Send the ZIP file for download
      res.download(zipOutputPath, 'converted_files.zip', (err) => {
        if (err) {
          console.error('Error sending ZIP file for download:', err);
        }

        // Delete the created ZIP file after download
        cleanupFiles(zipOutputPath);
      });
    });

    // Finalize the ZIP output stream
    zipOutput.finalize();
  } catch (error) {
    console.error('Error processing ZIP file:', error);
    res.status(500).send('Internal Server Error');
  }
}

// Function to convert a DOCX file to PDF and append it to the ZIP
async function convertAndAppendToZip(filePath, zip, totalFiles) {
  try {
    // Generate the output path for the converted PDF
    const outputFilePath = path.join(
      __dirname,
      'uploads',
      Date.now() + '-' + path.basename(filePath, path.extname(filePath)) + '.pdf'
    );

    // Convert the DOCX file to PDF
    await new Promise((resolve, reject) => {
      docxToPdf(filePath, outputFilePath, (err) => {
        if (err) {
          console.error('Error converting DOCX to PDF:', err);
          reject(err);
          // Clean up temporary files in case of an error
          cleanupFiles(outputFilePath);
        } else {
          // Check if the file exists before appending it to the ZIP
          if (fs.existsSync(outputFilePath)) {
            zip.append(fs.createReadStream(outputFilePath), {
              name: path.basename(filePath, path.extname(filePath)) + '.pdf',
            });
            resolve();
          } else {
            console.error(`File not found: ${outputFilePath}`);
            reject(new Error('Converted PDF file not found.'));
          }
        }

        // If all files are processed, call the callback function
        if (--totalFiles === 0) {
          resolve();
        }
      });
    });
  } catch (error) {
    console.error('Error in convertAndAppendToZip:', error);
    throw error; // Re-throw the error for centralized error handling
  }
}

// Function to clean up files (delete)
function cleanupFiles(...filePaths) {
  filePaths.forEach((filePath) => {
    fs.unlink(filePath, (err) => {
      if (err) {
        console.error('Error deleting file:', err);
      }
    });
  });
}

// Start the server and listen on the specified port
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
