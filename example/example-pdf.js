const fs = require('fs');
const path = require('path');

// Import the specific functions we need
const { generateContainer, convertDocxToPdf } = require('../dist/html-to-docx.umd');

// create the output path to root
const docxOutputPath = path.join(__dirname, '../example.docx');
const pdfOutputPath = path.join(__dirname, '../example.pdf');
const htmlFilePath = path.join(__dirname, 'example.html');

(async () => {
  try {
    console.log('Testing PDF conversion functionality...');

    // Check if HTML file exists
    if (!fs.existsSync(htmlFilePath)) {
      throw new Error(`HTML file not found: ${htmlFilePath}`);
    }

    const htmlString = fs.readFileSync(htmlFilePath, 'utf-8');
    console.log('HTML content loaded successfully');

    // First generate DOCX using the named export
    console.log('1. Generating DOCX from HTML...');
    const docxBuffer = await generateContainer(htmlString, null, {
      footer: true,
      pageNumber: true,
      preprocessing: { skipHTMLMinify: false },
    });

    fs.writeFileSync(docxOutputPath, docxBuffer);
    console.log('✓ DOCX file created:', docxOutputPath);

    console.log('3. Converting DOCX to PDF...');
    // Now convert to PDF using the native method
    const pdfBuffer = await convertDocxToPdf(docxBuffer);

    fs.writeFileSync(pdfOutputPath, pdfBuffer);
    console.log('✓ PDF file created:', pdfOutputPath);
    console.log('✅ PDF conversion test completed successfully!');
  } catch (error) {
    console.error('❌ Error:', error.message);

    // Provide helpful error messages
    if (error.message.includes('PDF conversion failed')) {
      console.log('\nTroubleshooting tips:');
      console.log('1. Make sure LibreOffice is installed');
      console.log('2. Ensure soffice command is available in your PATH');
      console.log('3. Try running: soffice --headless --convert-to pdf test.docx');
    }

    console.error('Stack:', error.stack);
  }
})();
