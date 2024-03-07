# Word to PDF Converter

This project is a simple Word to PDF converter.

## Description

The Word to PDF Converter is a command-line tool that converts Word documents (.docx) to PDF format. It utilizes a custom script to handle the conversion process efficiently.

## Installation

1. Clone the repository:
git clone https://github.com/Ogada8400/word-to-pdf.git


2. Install dependencies:
npm install


3. Run the server:
nodemon index.js


## Usage

1. Open your web browser and navigate to `http://localhost:3000`.
2. Upload one or more DOCX files using the provided form.
3. Click the "Convert to PDF" button to initiate the conversion process.
4. Once the conversion is complete, download the generated ZIP file containing the converted PDF files.

## API Documentation

The API documentation is available at `http://localhost:3000/api-docs`, where you can explore the available endpoints and test them using Swagger UI.

## Dependencies

- [Express](https://www.npmjs.com/package/express): Web framework for Node.js
- [Multer](https://www.npmjs.com/package/multer): Middleware for handling file uploads
- [Docx to Pdf](https://www.npmjs.com/package/docx-pdf): Library for converting DOCX files to PDF format
- [Swagger UI Express](https://www.npmjs.com/package/swagger-ui-express): Middleware for serving Swagger UI
- [Adm-Zip](https://www.npmjs.com/package/adm-zip): Library for working with ZIP archives
- [Body Parser](https://www.npmjs.com/package/body-parser): Middleware for parsing request bodies

## Authors

- Ian Ogada - [GitHub](https://github.com/Ogada8400)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
