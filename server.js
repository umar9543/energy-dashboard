const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
app.use(cors());

const EXCEL_FILE_PATH = path.join(__dirname, '260719 PST_Electricity_Dashboard_2026.xlsx');
const CARBON_FILE_PATH = path.join(__dirname, 'STARTEAM_PST_Electricity_Carbon_Footprint.xlsx');

// Load and parse Excel file once at startup (in a real app, you might want to cache or reload this periodically)
let sourceData = [];
let carbonData = [];

try {
    const workbook = xlsx.readFile(EXCEL_FILE_PATH);
    
    // Read Source Data sheet
    const sourceSheet = workbook.Sheets['Source Data'];
    
    // The data actually starts at row 5 (index 4), we need to parse it carefully
    const rawData = xlsx.utils.sheet_to_json(sourceSheet, { header: 1 });
    
    // Extract headers from the 5th row
    const headers = rawData[4];
    
    // Map the rest of the rows into objects
    for (let i = 5; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0 || !row[0] || String(row[0]).toLowerCase().includes('total')) {
            continue; // Skip empty rows and total rows
        }
        
        let rowObj = {};
        headers.forEach((header, index) => {
            if (header) {
                rowObj[header] = row[index] || 0;
            }
        });
        sourceData.push(rowObj);
    }
    console.log(`Loaded ${sourceData.length} records from Excel.`);
} catch (error) {
    console.error("Error loading Excel file:", error);
}

try {
    const carbonWorkbook = xlsx.readFile(CARBON_FILE_PATH);
    const carbonSheet = carbonWorkbook.Sheets['Carbon Footprint'];
    
    // The headers for carbon footprint are on row 4 (index 3)
    const carbonRawData = xlsx.utils.sheet_to_json(carbonSheet, { header: 1 });
    const carbonHeaders = carbonRawData[3];
    
    for (let i = 4; i < carbonRawData.length; i++) {
        const row = carbonRawData[i];
        if (!row || row.length === 0 || !row[0] || String(row[0]).toLowerCase().includes('total')) {
            continue; // Skip empty rows and total rows
        }
        
        let rowObj = {};
        carbonHeaders.forEach((header, index) => {
            if (header) {
                rowObj[header] = row[index] || 0;
            }
        });
        carbonData.push(rowObj);
    }
    console.log(`Loaded ${carbonData.length} records from Carbon Footprint Excel.`);
} catch (error) {
    console.error("Error loading Carbon Footprint Excel file:", error);
}

app.get('/api/dashboard', (req, res) => {
    const { year = 'All', month = 'All', region = 'All', priority = 'All', department = 'All' } = req.query;

    let filteredData = [...sourceData];
    let filteredCarbonData = [...carbonData];

    // Filter by Region (Data is only for China)
    if (region !== 'All' && region.toLowerCase() !== 'china') {
        filteredData = []; // No data for other regions like Thailand
        filteredCarbonData = [];
    }

    // Filter by Year (Data is only for 2026)
    if (year !== 'All' && year !== '2026') {
        filteredData = [];
        filteredCarbonData = [];
    }

    // Filter by Department (Maps to 'Process Step')
    if (department !== 'All') {
        const deptLower = String(department).toLowerCase();
        filteredData = filteredData.filter(row => 
            row['Process Step'] && 
            String(row['Process Step']).toLowerCase().includes(deptLower)
        );
        filteredCarbonData = filteredCarbonData.filter(row => 
            row['Process Step'] && 
            String(row['Process Step']).toLowerCase().includes(deptLower)
        );
    }
    
    // Priority filter (Not clearly mapped to a column, ignoring for now as it's just 'High'/'Low' consumption)
    // You could potentially filter by row.Total > threshold if needed.

    // Calculate Totals
    let totalConsumption = 0;
    let totalRecords = filteredData.length;
    let totalCarbonEmission = 0;
    
    let totalConsumptionByMonth = {
        Jan: 0, Feb: 0, Mar: 0, Apr: 0, May: 0, Jun: 0,
        Jul: 0, Aug: 0, Sep: 0, Oct: 0, Nov: 0, Dec: 0
    };

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    filteredData.forEach(row => {
        if (month !== 'All') {
            // If a specific month is selected, only add that month's consumption
            if (months.includes(month)) {
                totalConsumption += Number(row[month]) || 0;
                totalConsumptionByMonth[month] += Number(row[month]) || 0;
            }
        } else {
            // If all months, add up all months
            months.forEach(m => {
                totalConsumption += Number(row[m]) || 0;
                totalConsumptionByMonth[m] += Number(row[m]) || 0;
            });
        }
    });

    filteredCarbonData.forEach(row => {
        if (month !== 'All') {
            if (months.includes(month)) {
                totalCarbonEmission += Number(row[month]) || 0;
            }
        } else {
            months.forEach(m => {
                totalCarbonEmission += Number(row[m]) || 0;
            });
        }
    });

    const response = {
        mainResult: {
            totalConsumption: totalConsumption,
            totalCost: 0,
            totalCarbonEmission: totalCarbonEmission,
            totalRecords: totalRecords,
            avgEfficiencyRating: 0,
            utilizationRate: 0
        },
        totalConsumptionByMonth: totalConsumptionByMonth,
        totalCostByFacility: {
            // Leaving empty or dummy as cost is 0
        },
        totalConsumptionByEnergy: {
            Electricity: totalConsumption,
            Diesel: 0,
            "Natural Gas": 0,
            Solar: 0,
            Wind: 0
        }
    };

    res.json(response);
});

app.get('/api/processes', (req, res) => {
    const processSet = new Set();
    sourceData.forEach(row => {
        if (row['Process Step']) {
            // Remove leading numbers and hyphens (e.g. "1 - Board Cut" -> "Board Cut")
            const cleanName = String(row['Process Step']).replace(/^\d+\s*-\s*/, '').trim();
            processSet.add(cleanName);
        }
    });
    const processes = Array.from(processSet).sort();
    res.json(processes);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
