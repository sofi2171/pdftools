// Professional Excel Generator for PDF to Excel Conversion
class ProfessionalExcelGenerator {
    constructor() {
        this.workbook = null;
        this.worksheet = null;
    }

    async createProfessionalExcel(pdfContent, fileName) {
        // Create workbook structure
        const workbook = {
            SheetNames: ['PDF_Data'],
            Sheets: {},
            Props: {
                Title: fileName,
                Subject: 'PDF to Excel Conversion',
                Author: 'PDFTools2026',
                CreatedDate: new Date()
            }
        };

        // Process PDF content into structured data
        const structuredData = this.processPDFContent(pdfContent);
        
        // Create professional worksheet
        const worksheet = this.createWorksheet(structuredData);
        
        workbook.Sheets['PDF_Data'] = worksheet;
        
        // Generate Excel file
        return this.generateExcelFile(workbook, fileName);
    }

    processPDFContent(content) {
        const lines = content.split('\n').filter(line => line.trim());
        const data = [];
        
        // Add professional headers
        data.push(['Page', 'Content Type', 'Text Content', 'Line Number']);
        
        let currentPage = 1;
        let lineNumber = 1;
        
        lines.forEach(line => {
            const trimmedLine = line.trim();
            
            // Detect page breaks
            if (trimmedLine.includes('=== Page') || trimmedLine.includes('Page ')) {
                const pageMatch = trimmedLine.match(/(\d+)/);
                if (pageMatch) {
                    currentPage = parseInt(pageMatch[1]);
                    lineNumber = 1;
                    return;
                }
            }
            
            if (trimmedLine && !trimmedLine.includes('===')) {
                // Detect content type
                let contentType = 'Text';
                if (this.isNumeric(trimmedLine)) {
                    contentType = 'Number';
                } else if (this.isDate(trimmedLine)) {
                    contentType = 'Date';
                } else if (this.isEmail(trimmedLine)) {
                    contentType = 'Email';
                } else if (this.isURL(trimmedLine)) {
                    contentType = 'URL';
                } else if (this.isHeader(trimmedLine)) {
                    contentType = 'Header';
                }
                
                data.push([currentPage, contentType, trimmedLine, lineNumber]);
                lineNumber++;
            }
        });
        
        return data;
    }

    createWorksheet(data) {
        const worksheet = {};
        const range = { s: { c: 0, r: 0 }, e: { c: 3, r: data.length - 1 } };
        
        // Add data to worksheet
        data.forEach((row, rowIndex) => {
            row.forEach((cell, colIndex) => {
                const cellAddress = this.encodeCellAddress(rowIndex, colIndex);
                worksheet[cellAddress] = {
                    v: cell,
                    t: typeof cell === 'number' ? 'n' : 's'
                };
                
                // Style header row
                if (rowIndex === 0) {
                    worksheet[cellAddress].s = {
                        font: { bold: true, color: { rgb: "FFFFFF" } },
                        fill: { fgColor: { rgb: "366092" } },
                        alignment: { horizontal: "center", vertical: "center" },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } }
                        }
                    };
                } else {
                    // Style data rows
                    worksheet[cellAddress].s = {
                        alignment: { horizontal: "left", vertical: "top", wrapText: true },
                        border: {
                            top: { style: "thin", color: { rgb: "CCCCCC" } },
                            bottom: { style: "thin", color: { rgb: "CCCCCC" } },
                            left: { style: "thin", color: { rgb: "CCCCCC" } },
                            right: { style: "thin", color: { rgb: "CCCCCC" } }
                        }
                    };
                    
                    // Alternate row colors
                    if (rowIndex % 2 === 0) {
                        worksheet[cellAddress].s.fill = { fgColor: { rgb: "F8F9FA" } };
                    }
                }
            });
        });
        
        worksheet['!ref'] = this.encodeRange(range);
        
        // Set column widths
        worksheet['!cols'] = [
            { wch: 8 },   // Page
            { wch: 15 },  // Content Type
            { wch: 50 },  // Text Content
            { wch: 12 }   // Line Number
        ];
        
        // Set row heights
        worksheet['!rows'] = data.map(() => ({ hpt: 20 }));
        
        return worksheet;
    }

    generateExcelFile(workbook, fileName) {
        // Create CSV format for compatibility
        const worksheet = workbook.Sheets['PDF_Data'];
        const csvContent = this.worksheetToCSV(worksheet);
        
        // Create professional CSV with proper formatting
        const professionalCSV = this.formatProfessionalCSV(csvContent);
        
        return new Blob([professionalCSV], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
    }

    worksheetToCSV(worksheet) {
        const range = worksheet['!ref'];
        if (!range) return '';
        
        const decoded = this.decodeRange(range);
        const csv = [];
        
        for (let row = decoded.s.r; row <= decoded.e.r; row++) {
            const csvRow = [];
            for (let col = decoded.s.c; col <= decoded.e.c; col++) {
                const cellAddress = this.encodeCellAddress(row, col);
                const cell = worksheet[cellAddress];
                csvRow.push(cell ? `"${cell.v}"` : '""');
            }
            csv.push(csvRow.join(','));
        }
        
        return csv.join('\n');
    }

    formatProfessionalCSV(csvContent) {
        // Add BOM for proper UTF-8 encoding
        const BOM = '\uFEFF';
        
        // Add professional metadata
        const metadata = [
            `"Generated by","PDFTools2026"`,
            `"Creation Date","${new Date().toISOString()}"`,
            `"Format","Professional Excel Spreadsheet"`,
            `"",""`
        ].join('\n');
        
        return BOM + metadata + '\n' + csvContent;
    }

    // Utility functions
    isNumeric(str) {
        return /^\d+(\.\d+)?$/.test(str.trim());
    }

    isDate(str) {
        return /\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(str);
    }

    isEmail(str) {
        return /\S+@\S+\.\S+/.test(str);
    }

    isURL(str) {
        return /https?:\/\/\S+/.test(str);
    }

    isHeader(str) {
        return str.length < 50 && (str.toUpperCase() === str || /^[A-Z\s]+$/.test(str));
    }

    encodeCellAddress(row, col) {
        return String.fromCharCode(65 + col) + (row + 1);
    }

    encodeRange(range) {
        return this.encodeCellAddress(range.s.r, range.s.c) + ':' + 
               this.encodeCellAddress(range.e.r, range.e.c);
    }

    decodeRange(range) {
        const parts = range.split(':');
        return {
            s: this.decodeCellAddress(parts[0]),
            e: this.decodeCellAddress(parts[1])
        };
    }

    decodeCellAddress(address) {
        const match = address.match(/([A-Z]+)(\d+)/);
        return {
            c: match[1].charCodeAt(0) - 65,
            r: parseInt(match[2]) - 1
        };
    }
}

// Export for use in main application
window.ProfessionalExcelGenerator = ProfessionalExcelGenerator;