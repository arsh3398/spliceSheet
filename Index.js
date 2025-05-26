const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const app = express();

// Configure multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Fiber color coding standards
const BUFFER_COLORS = ['BL', 'OR', 'GR', 'BR', 'SL', 'WH', 'RD', 'BK', 'YL', 'VT', 'RS', 'AQ'];
const FIBER_COLORS = ['bl', 'or', 'gr', 'br', 'sl', 'wh', 'rd', 'bk', 'yl', 'vt', 'rs', 'aq'];

// Sample data structure for demonstration
const SAMPLE_INPUT_DATA = [
  {
    ports: 96,
    cables: [
      { name: '144F(1)', fiberCount: 144 },
      { name: '144F(2)', fiberCount: 144 },
      { name: '48F(3)', fiberCount: 48 },
      { name: '48F(4)', fiberCount: 48 }
    ],
    mstPrefix: 'MST_F',
    addresses: [
      { mst: 'MST_F1245ECOATSAVE.210819', address: '2101 MARENGO LK RD', sheet: 14, terminal: 'T1' },
      { mst: 'MST_F1245ECOATSAVE.210819', address: '1245 E COATS AVE', sheet: 14 },
      { mst: 'MST_F1245ECOATSAVE.210819', address: 'VAC', sheet: 15 },
      { mst: 'MST_F976ECOATSAVE.210820', address: '821 E COATS AVE', sheet: 16, terminal: 'T2' },
      // Add more addresses as needed
    ]
  }
];

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
      cables.forEach((cable, cableIndex) => {
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
          // Add null values for unused cable positions
          row.push(port, cable.name, null, null, null, null);
        }
      });

      // Add MST and address information
      if (addressIndex < addresses.length) {
        const addressInfo = addresses[addressIndex];
        row.push(addressInfo.mst, addressInfo.address);
        
        // Add sheet and terminal info if available
        if (addressInfo.sheet) {
          row.push(`SHEET # ${addressInfo.sheet}`);
        }
        if (addressInfo.terminal) {
          row.push(addressInfo.terminal);
        }
        
        // Move to next address based on some logic (every few ports or specific pattern)
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
    const headers = ['Port #', 'Cable Name'];
    
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
      const headers = data[0];
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
      if (header && header.toLowerCase().includes('cable')) {
        // Extract cable info (this would need to be customized based on actual input format)
        const cableMatch = header.match(/Cable\s*(\d+)/i);
        if (cableMatch) {
          cables.push({
            name: `${header}`,
            fiberCount: 144 // Default, could be parsed from data
          });
        }
      }
    });

    // Parse address data if provided in subsequent rows
    data.forEach(row => {
      if (row && row.length > 0) {
        // Parse address information based on input structure
        const addressData = {
          mst: row[4] || '',
          address: row[5] || 'Unused',
          sheet: Math.floor(Math.random() * 20) + 1, // Random for demo
          terminal: row.length > 6 ? row[6] : undefined
        };
        addresses.push(addressData);
      }
    });

    return {
      ports: 360, // Default or parsed from input
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
    
    // Add some styling and formatting
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    
    // Set column widths
    worksheet['!cols'] = [
      { width: 8 },   // Port #
      { width: 20 },  // Cable Name
      { width: 8 },   // Port #
      { width: 12 },  // Cable
      { width: 6 },   // B#
      { width: 6 },   // (B)
      { width: 6 },   // F#
      { width: 6 },   // (F)
      { width: 12 },  // Cable
      { width: 6 },   // B#
      { width: 6 },   // (B)
      { width: 6 },   // F#
      { width: 6 },   // (F)
      { width: 30 },  // MST
      { width: 25 },  // Address
    ];
    
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Splice Sheet');
    XLSX.writeFile(workbook, filename);
    
    return filename;
  }
}

// API Routes
app.use(express.json());

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', message: 'Splice Sheet Generator API is running' });
});

// Generate splice sheet from uploaded Excel file
app.post('/generate-splice-sheet', upload.single('inputFile'), (req, res) => {
  try {
    const generator = new SpliceSheetGenerator();
    
    let inputData;
    
    if (req.file) {
      // Parse uploaded file
      inputData = generator.parseInputFile(req.file.path);
    } else {
      // Use sample data or request body
      inputData = req.body.inputData || SAMPLE_INPUT_DATA[0];
    }
    
    // Generate splice sheet
    const spliceData = generator.generateSpliceSheet(inputData);
    
    // Export to Excel
    const filename = `splice_sheet_${Date.now()}.xlsx`;
    generator.exportToExcel(spliceData, filename);
    
    res.json({
      success: true,
      message: 'Splice sheet generated successfully',
      filename: filename,
      rowCount: spliceData.length - 1, // Excluding header
      data: spliceData.slice(0, 10) // Return first 10 rows as preview
    });
    
  } catch (error) {
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
    
    res.json({
      success: true,
      message: 'Custom splice sheet generated successfully',
      data: spliceData,
      summary: {
        totalPorts: ports,
        cables: cables.length,
        addresses: inputData.addresses.length
      }
    });
    
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error generating custom splice sheet',
      error: error.message
    });
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Splice Sheet Generator API running on port ${PORT}`);
  console.log(`Health check: http://localhost:${PORT}/health`);
});

module.exports = app;