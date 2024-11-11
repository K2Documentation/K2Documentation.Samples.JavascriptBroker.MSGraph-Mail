const fs = require('fs');
const path = require('path');

// Specify the directory containing the output files
const outputDir = path.resolve(__dirname, 'dist');

// Function to remove newlines from a file
function removeNewlines(filePath) {
    const content = fs.readFileSync(filePath, 'utf8');
    const singleLineContent = content
        .replace(/\n+/g, '')
        .replace(/\t+/g, '')
        .replace(/\s/g, ''); // Remove all newlines, tabs, spaces
    fs.writeFileSync(filePath, singleLineContent, 'utf8');
}

// Process each file in the output directory
function processDirectory(directory) {
    fs.readdirSync(directory).forEach(file => {
        const fullPath = path.join(directory, file);
        if (fs.statSync(fullPath).isDirectory()) {
            processDirectory(fullPath); // Recursively process subdirectories
        } else if (file.endsWith('.min.js')) { // Only look at .min files
            removeNewlines(fullPath);
        }
    });
}

processDirectory(outputDir);
console.log('Post-processing complete: All files are now single-line.');
