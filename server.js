const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const app = express();

// Enable CORS for frontend
app.use(cors());
app.use(express.json());
app.use(express.static('public')); // Serve static files

// Configure multer for file uploads
const upload = multer({ 
  dest: 'uploads/',
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['.xlsx', '.xls'];
    const fileExt = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(fileExt)) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files are allowed'), false);
    }
  }
});

// Fiber color coding standards
const BUFFER_COLORS = ['BL', 'OR', 'GR', 'BR', 'SL', 'WH', 'RD', 'BK', 'YL', 'VT', 'RS', 'AQ'];
const FIBER_COLORS = ['bl', 'or', 'gr', 'br', 'sl', 'wh', 'rd', 'bk', 'yl', 'vt', 'rs', 'aq'];

class SpliceSheetGenerator {
  constructor() {
    this.fiberColorIndex = 0;
    this.bufferTubeIndex = 0;
  }

  // Calculate buffer tube and fiber numbers based on position
  calculateFiberPosition(fiberNumber, fibersPerTube = 12) {
    const bufferTubeNumber = Math.floor((fiberNumber - 1) / fibersPerTube) + 1;
    const fiberInTube = ((fiberNumber - 1) % fibersPerTube) + 1;
    
    return {
      bufferTube: bufferTubeNumber,
      bufferColor: BUFFER_COLORS[(bufferTubeNumber - 1) % BUFFER_COLORS.length],
      fiberNumber: fiberInTube,
      fiberColor: FIBER_COLORS[(fiberInTube - 1) % FIBER_COLORS.length]
    };
  }

  // Generate splice sheet data
  generateSpliceSheet(inputData) {
    const {
      ports = 96,
      mainCableName = 'FDH108_144F_1-96',
      cables = [],
      addresses = []
    } = inputData;

    const spliceData = [];
    let addressIndex = 0;
    
    // Generate header row
    const headers = this.generateHeaders(cables);
    spliceData.push(headers);

    // Generate data rows
    for (let port = 1; port <= ports; port++) {
      const row = [port, mainCableName];
      
      // Add cable information for each defined cable
      cables.forEach((cable) => {
        if (port <= cable.fiberCount) {
          const fiberPos = this.calculateFiberPosition(port);
          row.push(
            port, // Port number for this cable
            cable.name,
            fiberPos.bufferTube,
            fiberPos.bufferColor,
            fiberPos.fiberNumber,
            fiberPos.fiberColor
          );
        } else {
          // Add empty values for unused cable positions
          row.push('', cable.name, '', '', '', '');
        }
      });

      // Add MST and address information
      if (addressIndex < addresses.length) {
        const addressInfo = addresses[addressIndex];
        row.push(addressInfo.mst || '', addressInfo.address || '');
        
        // Add sheet and terminal info if available
        if (addressInfo.sheet) {
          row.push(`SHEET # ${addressInfo.sheet}`);
        }
        if (addressInfo.terminal) {
          row.push(addressInfo.terminal);
        }
        
        // Move to next address based on pattern (every 4 ports)
        if (port % 4 === 0 && addressIndex < addresses.length - 1) {
          addressIndex++;
        }
      } else {
        // Mark as unused if no address data
        row.push('', 'Unused');
      }

      spliceData.push(row);
    }

    return spliceData;
  }

  // Generate dynamic headers based on cable configuration
  generateHeaders(cables) {
    const headers = ['Port #', 'Main Cable'];
    
    cables.forEach(cable => {
      headers.push('Port #', 'Cable', 'B#', '(B)', 'F#', '(F)');
    });
    
    headers.push('MST', 'Address');
    return headers;
  }

  // Parse input Excel file
  parseInputFile(filePath) {
    try {
      const workbook = XLSX.readFile(filePath);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      // Parse the input structure
      const headers = data[0] || [];
      const inputData = this.parseInputStructure(headers, data.slice(1));
      
      return inputData;
    } catch (error) {
      throw new Error(`Error parsing input file: ${error.message}`);
    }
  }

  // Parse input structure to understand cable configuration
  parseInputStructure(headers, data) {
    const cables = [];
    const addresses = [];
    
    // Extract cable information from headers
    headers.forEach((header, index) => {
      if (header && typeof header === 'string' && header.toLowerCase().includes('cable')) {
        const cableMatch = header.match(/(\d+)F/i);
        const fiberCount = cableMatch ? parseInt(cableMatch[1]) : 144;
        cables.push({
          name: header,
          fiberCount: fiberCount
        });
      }
    });

    // Parse address data if provided in subsequent rows
    data.forEach((row, index) => {
      if (row && row.length > 0) {
        const addressData = {
          mst: row[4] || `MST_F${1000 + index}ECOATSAVE.21082${index % 10}`,
          address: row[5] || 'Unused',
          sheet: row[6] ? parseInt(row[6]) : Math.floor(index / 3) + 10,
          terminal: row[7] || (index < 2 ? `T${index + 1}` : undefined)
        };
        addresses.push(addressData);
      }
    });

    return {
      ports: data.length > 0 ? Math.max(96, data.length) : 96,
      mainCableName: 'FDH108_144F_1-96',
      cables: cables.length > 0 ? cables : [
        { name: '144F(1)', fiberCount: 144 },
        { name: '144F(2)', fiberCount: 144 },
        { name: '48F(3)', fiberCount: 48 },
        { name: '48F(4)', fiberCount: 48 }
      ],
      addresses: addresses.length > 0 ? addresses : this.generateSampleAddresses()
    };
  }

  // Generate sample addresses for demonstration
  generateSampleAddresses() {
    const sampleAddresses = [
      '2101 MARENGO LK RD', '1245 E COATS AVE', '821 E COATS AVE', '801 E COATS AVE',
      '976 E COATS AVE', '764 E COATS AVE', '268 HIGHLAND CIR', '608 7TH AVE',
      '702 7TH AVE', '302 TUCKER ST', '207 TUCKER ST', '205 TUCKER ST',
      '101 E COATS AVE', '108 E COATS AVE'
    ];
    
    return sampleAddresses.map((addr, index) => ({
      mst: `MST_F${1000 + index}ECOATSAVE.21082${index % 10}`,
      address: addr,
      sheet: Math.floor(index / 3) + 10,
      terminal: index < 2 ? `T${index + 1}` : undefined
    }));
  }

  // Export to Excel format
  exportToExcel(spliceData, filename = 'splice_sheet.xlsx') {
    const worksheet = XLSX.utils.aoa_to_sheet(spliceData);
    const workbook = XLSX.utils.book_new();
    
    // Set column widths
    worksheet['!cols'] = [
      { width: 8 },   // Port #
      { width: 20 },  // Main Cable
      { width: 8 },   // Port #
      { width: 15 },  // Cable
      { width: 6 },   // B#
      { width: 6 },   // (B)
      { width: 6 },   // F#
      { width: 6 },   // (F)
      { width: 8 },   // Port #
      { width: 15 },  // Cable
      { width: 6 },   // B#
      { width: 6 },   // (B)
      { width: 6 },   // F#
      { width: 6 },   // (F)
      { width: 30 },  // MST
      { width: 25 },  // Address
    ];
    
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Splice Sheet');
    
    // Ensure output directory exists
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    const outputPath = path.join(outputDir, filename);
    XLSX.writeFile(workbook, outputPath);
    
    return outputPath;
  }
}

// API Routes

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', message: 'Splice Sheet Generator API is running' });
});

// Serve the frontend
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Generate splice sheet from uploaded Excel file
app.post('/generate-splice-sheet', upload.single('inputFile'), (req, res) => {
  try {
    const generator = new SpliceSheetGenerator();
    
    let inputData;
    
    if (req.file) {
      // Parse uploaded file
      inputData = generator.parseInputFile(req.file.path);
      
      // Clean up uploaded file
      fs.unlink(req.file.path, (err) => {
        if (err) console.error('Error deleting uploaded file:', err);
      });
    } else {
      // Use default sample data
      inputData = {
        ports: 96,
        mainCableName: 'FDH108_144F_1-96',
        cables: [
          { name: '144F(1)', fiberCount: 144 },
          { name: '144F(2)', fiberCount: 144 }
        ],
        addresses: generator.generateSampleAddresses()
      };
    }
    
    // Generate splice sheet
    const spliceData = generator.generateSpliceSheet(inputData);
    
    // Export to Excel
    const filename = `splice_sheet_${Date.now()}.xlsx`;
    const filePath = generator.exportToExcel(spliceData, filename);
    
    res.json({
      success: true,
      message: 'Splice sheet generated successfully',
      filename: filename,
      rowCount: spliceData.length - 1, // Excluding header
      downloadUrl: `/download/${filename}`,
      preview: spliceData.slice(0, 10) // Return first 10 rows as preview
    });
    
  } catch (error) {
    console.error('Error generating splice sheet:', error);
    res.status(500).json({
      success: false,
      message: 'Error generating splice sheet',
      error: error.message
    });
  }
});

// Generate splice sheet with custom parameters
app.post('/generate-custom-splice-sheet', (req, res) => {
  try {
    const generator = new SpliceSheetGenerator();
    const {
      ports = 96,
      mainCableName = 'FDH108_144F_1-96',
      cables = [
        { name: '144F(1)', fiberCount: 144 },
        { name: '144F(2)', fiberCount: 144 }
      ],
      addresses = []
    } = req.body;
    
    const inputData = {
      ports,
      mainCableName,
      cables,
      addresses: addresses.length > 0 ? addresses : generator.generateSampleAddresses()
    };
    
    const spliceData = generator.generateSpliceSheet(inputData);
    
    // Export to Excel
    const filename = `custom_splice_sheet_${Date.now()}.xlsx`;
    const filePath = generator.exportToExcel(spliceData, filename);
    
    res.json({
      success: true,
      message: 'Custom splice sheet generated successfully',
      filename: filename,
      downloadUrl: `/download/${filename}`,
      preview: spliceData.slice(0, 10),
      summary: {
        totalPorts: ports,
        cables: cables.length,
        addresses: inputData.addresses.length
      }
    });
    
  } catch (error) {
    console.error('Error generating custom splice sheet:', error);
    res.status(500).json({
      success: false,
      message: 'Error generating custom splice sheet',
      error: error.message
    });
  }
});

// Download generated file
app.get('/download/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'output', filename);
  
  if (fs.existsSync(filePath)) {
    res.download(filePath, filename, (err) => {
      if (err) {
        console.error('Error downloading file:', err);
        res.status(500).json({ error: 'Error downloading file' });
      }
    });
  } else {
    res.status(404).json({ error: 'File not found' });
  }
});

// Get fiber color standards
app.get('/fiber-standards', (req, res) => {
  res.json({
    bufferColors: BUFFER_COLORS,
    fiberColors: FIBER_COLORS,
    standardFibersPerTube: 12,
    colorCodingStandard: 'TIA-598-C'
  });
});

// Error handling middleware
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File too large' });
    }
  }
  
  res.status(500).json({ error: error.message });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Splice Sheet Generator API running on port ${PORT}`);
  console.log(`Frontend available at: http://localhost:${PORT}`);
  console.log(`Health check: http://localhost:${PORT}/health`);
  
  // Create necessary directories
  const dirs = ['uploads', 'output', 'public'];
  dirs.forEach(dir => {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
  });
});

module.exports = app;