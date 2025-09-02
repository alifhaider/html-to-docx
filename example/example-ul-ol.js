const fs = require('fs');
const { HTMLtoDOCX } = require('../dist/html-to-docx.umd');

const outputPath = './example-ul-ol.docx';

const htmlString = fs.readFileSync('./example/ul-ol.html', 'utf-8');

(async () => {
  const fileBuffer = await HTMLtoDOCX(htmlString, null, {
    table: {
      row: { cantSplit: true },
      addSpacingAfter: true,
    },
    footer: true,
    pageNumber: true,
    preprocessing: { skipHTMLMinify: false },
  });

  fs.writeFile(outputPath, fileBuffer, (error) => {
    if (error) {
      console.log('Docx file creation failed');
      return;
    }
    console.log('Docx file created successfully');
  });
})();
