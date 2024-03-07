const swaggerJsdoc = require('swagger-jsdoc');
const { serveWithOptions } = require('swagger-ui-express');

const options = {
  definition: {
    openapi: '3.0.0',
    info: {
      title: 'Word to pdf Web App API Documentation',
      version: '1.0.0',
      description: 'The DOCX to PDF Conversion Web App allows users to convert DOCX files to PDF format. It supports both individual DOCX files and ZIP archives containing multiple DOCX files. The converted PDF files are then bundled into a ZIP file for easy download..',
    },
  },
  apis: ['./index.js'], // Path to the file containing API routes
};

const swaggerSpec = swaggerJsdoc(options);

module.exports = swaggerSpec;
