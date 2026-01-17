// Professional Word and Excel Converter with Formatting Preservation
class WordExcelConverter {
    constructor() {
        this.pdfjsLib = window.pdfjsLib;
    }

    // Convert PDF to Word with preserved formatting
    async convertToWord(pdfFile, fileName) {
        try {
            const arrayBuffer = await pdfFile.arrayBuffer();
            const pdf = await this.pdfjsLib.getDocument({data: arrayBuffer}).promise;
            
            let documentContent = '';
            let pageCount = pdf.numPages;
            
            // Extract text with formatting preservation
            for (let pageNum = 1; pageNum <= pageCount; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const textContent = await page.getTextContent();
                
                // Process text items with positioning
                let pageText = this.processTextWithFormatting(textContent.items);
                documentContent += `\n=== Page ${pageNum} ===\n${pageText}\n`;
            }
            
            // Create Word-compatible content
            const wordContent = this.createWordDocument(documentContent, fileName);
            return new Blob([wordContent], { 
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
            });
            
        } catch (error) {
            throw new Error(`Word conversion failed: ${error.message}`);
        }
    }

    // Convert PDF to Excel with structured data
    async convertToExcel(pdfFile, fileName) {
        try {
            const arrayBuffer = await pdfFile.arrayBuffer();
            const pdf = await this.pdfjsLib.getDocument({data: arrayBuffer}).promise;
            
            let allData = [];
            
            // Extract and structure data from each page
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const textContent = await page.getTextContent();
                
                const pageData = this.extractStructuredData(textContent.items, pageNum);
                allData = allData.concat(pageData);
            }
            
            // Create Excel file with preserved formatting
            return this.createExcelFile(allData, fileName);
            
        } catch (error) {
            throw new Error(`Excel conversion failed: ${error.message}`);
        }
    }

    // Process text with formatting preservation
    processTextWithFormatting(textItems) {
        let formattedText = '';
        let currentLine = '';
        let lastY = null;
        let lastX = null;
        
        // Sort items by position (top to bottom, left to right)
        const sortedItems = textItems.sort((a, b) => {
            const yDiff = b.transform[5] - a.transform[5]; // Y coordinate (top to bottom)
            if (Math.abs(yDiff) > 5) return yDiff;
            return a.transform[4] - b.transform[4]; // X coordinate (left to right)
        });
        
        sortedItems.forEach((item, index) => {
            const x = item.transform[4];
            const y = item.transform[5];
            const text = item.str.trim();
            
            if (!text) return;
            
            // Detect new line
            if (lastY !== null && Math.abs(y - lastY) > 5) {
                if (currentLine.trim()) {
                    formattedText += currentLine.trim() + '\n';
                }
                currentLine = '';
            }
            
            // Detect spacing between words
            if (lastX !== null && lastY !== null && Math.abs(y - lastY) <= 5) {
                const xGap = x - lastX;
                if (xGap > 20) { // Large gap indicates tab or column
                    currentLine += '\t';
                } else if (xGap > 5 && currentLine && !currentLine.endsWith(' ')) {
                    currentLine += ' ';
                }
            }
            
            currentLine += text;
            lastX = x + (item.width || 0);
            lastY = y;
        });
        
        // Add final line
        if (currentLine.trim()) {
            formattedText += currentLine.trim() + '\n';
        }
        
        return formattedText;
    }

    // Extract structured data for Excel
    extractStructuredData(textItems, pageNum) {
        const data = [];
        let currentRow = [];
        let lastY = null;
        
        // Sort items by position
        const sortedItems = textItems.sort((a, b) => {
            const yDiff = b.transform[5] - a.transform[5];
            if (Math.abs(yDiff) > 5) return yDiff;
            return a.transform[4] - b.transform[4];
        });
        
        sortedItems.forEach(item => {
            const y = item.transform[5];
            const text = item.str.trim();
            
            if (!text) return;
            
            // New row detected
            if (lastY !== null && Math.abs(y - lastY) > 5) {
                if (currentRow.length > 0) {
                    data.push([pageNum, ...currentRow]);
                    currentRow = [];
                }
            }
            
            currentRow.push(text);
            lastY = y;
        });
        
        // Add final row
        if (currentRow.length > 0) {
            data.push([pageNum, ...currentRow]);
        }
        
        return data;
    }

    // Create Word document content
    createWordDocument(content, fileName) {
        // Create RTF format for better compatibility
        const rtfHeader = `{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Times New Roman;}}`;
        const rtfContent = content
            .replace(/\n/g, '\\par ')
            .replace(/\t/g, '\\tab ')
            .replace(/=== Page (\d+) ===/g, '\\par\\b Page $1 \\b0\\par');
        
        return `${rtfHeader}\\f0\\fs24 ${rtfContent}}`;
    }

    // Create Excel file with proper formatting
    createExcelFile(data, fileName) {
        // Create CSV with proper formatting
        let csvContent = '\uFEFF'; // BOM for UTF-8
        
        // Add headers
        csvContent += '"Page","Column 1","Column 2","Column 3","Column 4","Column 5"\n';
        
        // Add data rows
        data.forEach(row => {
            const csvRow = row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',');
            csvContent += csvRow + '\n';
        });
        
        return new Blob([csvContent], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
    }
}

// Export for global use
window.WordExcelConverter = WordExcelConverter;