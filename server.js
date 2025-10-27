const express = require('express');
const http = require('http');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

const PORT = 3000;
const DB_FILE = 'database.json';
const EXCHANGE_RATE = 4100; // 1 EUR = 4100 UGX

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({ 
    storage: storage,
    limits: {
        fileSize: 10 * 1024 * 1024, // 10MB limit
    },
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
            file.mimetype === 'application/vnd.ms-excel') {
            cb(null, true);
        } else {
            cb(new Error('Only Excel files are allowed'), false);
        }
    }
});

// Middleware
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

// WebSocket connection for real-time updates
wss.on('connection', (ws) => {
    console.log('New client connected');
    
    // Send current database to new client
    const database = loadDatabase();
    ws.send(JSON.stringify({
        type: 'database_loaded',
        data: database,
        message: 'Database loaded successfully'
    }));

    ws.on('close', () => {
        console.log('Client disconnected');
    });
});

// Broadcast to all connected clients
function broadcast(message) {
    wss.clients.forEach(client => {
        if (client.readyState === WebSocket.OPEN) {
            client.send(JSON.stringify(message));
        }
    });
}

// Database functions
function loadDatabase() {
    try {
        if (fs.existsSync(DB_FILE)) {
            const data = fs.readFileSync(DB_FILE, 'utf8');
            return JSON.parse(data);
        } else {
            const emptyDatabase = createEmptyDatabase();
            saveDatabase(emptyDatabase);
            return emptyDatabase;
        }
    } catch (error) {
        console.error('Error loading database:', error);
        return createEmptyDatabase();
    }
}

function saveDatabase(data) {
    try {
        data.last_updated = new Date().toISOString();
        // Recalculate all metadata before saving
        data = recalculateAllMetadata(data);
        fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2));
        return true;
    } catch (error) {
        console.error('Error saving database:', error);
        return false;
    }
}

function createEmptyDatabase() {
    return {
        sponsorship_programs: {
            "CH": {
                "program_name": "CH FINANCIAL ANALYSIS REPORT TERM III 2025",
                "students": [],
                "metadata": {
                    "total_students": 0,
                    "sponsorship_types": {},
                    "monthly_costs_ugx": 0,
                    "monthly_costs_eur": 0
                }
            },
            "YSP": {
                "program_name": "YSP FINANCIAL ANALYSIS REPORT Term III 2025",
                "students": [],
                "metadata": {
                    "total_students": 0,
                    "sponsorship_types": {},
                    "monthly_costs_ugx": 0,
                    "monthly_costs_eur": 0
                }
            },
            "ICCSP": {
                "program_name": "ICCSP FINANCIAL ANALYSIS REPORT TERM III 2025",
                "students": [],
                "metadata": {
                    "total_students": 0,
                    "sponsorship_types": {},
                    "monthly_costs_ugx": 0,
                    "monthly_costs_eur": 0
                }
            },
            "OTM_GA": {
                "program_name": "OTM-GA FINANCIAL ANALYSIS REPORT TERM III 2025",
                "students": [],
                "metadata": {
                    "total_students": 0,
                    "sponsorship_types": {},
                    "monthly_costs_ugx": 0,
                    "monthly_costs_eur": 0
                }
            }
        },
        sponsorship_registry: {
            "CH": { "students": [], "metadata": { "total_students": 0, "active_students": 0, "total_monthly_funding": 0 } },
            "YSP": { "students": [], "metadata": { "total_students": 0, "active_students": 0, "total_monthly_funding": 0 } },
            "ICCSP": { "students": [], "metadata": { "total_students": 0, "active_students": 0, "total_monthly_funding": 0 } },
            "OTM_GA": { "students": [], "metadata": { "total_students": 0, "active_students": 0, "total_monthly_funding": 0 } }
        },
        financial_review: {},
        daily_expenses: [],
        events: [],
        system_settings: {
            organization: {
                name: "Sponsorship Pro",
                currency: "EUR",
                logo: null
            },
            forex: {
                manual_rate: EXCHANGE_RATE,
                auto_update: false
            },
            notifications: {
                email: true,
                event_reminders: true,
                low_balance: true
            }
        },
        metadata: {
            source_files: [],
            extraction_date: new Date().toISOString().split('T')[0],
            currency_conversion_rate: { euro_to_ugx: EXCHANGE_RATE },
            programs_summary: {
                total_students_across_all_programs: 0,
                total_active_sponsorships: 0,
                total_monthly_funding_euros: 0,
                total_monthly_funding_ugx: 0
            }
        },
        last_updated: new Date().toISOString()
    };
}

// Enhanced financial calculations
function recalculateAllMetadata(database) {
    let totalStudents = 0;
    let totalActiveSponsorships = 0;
    let totalMonthlyFundingEUR = 0;
    
    // Calculate program-level metadata
    Object.entries(database.sponsorship_programs).forEach(([programName, program]) => {
        if (program.students) {
            // Update program metadata
            program.metadata = program.metadata || {};
            program.metadata.total_students = program.students.length;
            totalStudents += program.students.length;
            
            // Calculate sponsorship types distribution
            program.metadata.sponsorship_types = {};
            program.students.forEach(student => {
                const packageType = student.sponsorship_package;
                program.metadata.sponsorship_types[packageType] = (program.metadata.sponsorship_types[packageType] || 0) + 1;
            });
            
            // Calculate program financials
            let programMonthlyCostUGX = 0;
            program.students.forEach(student => {
                const financials = calculateStudentFinancials(student.financial_data || {});
                programMonthlyCostUGX += financials.monthly_output_ugx;
                
                // Update student financial data with calculated values
                student.financial_data = { ...student.financial_data, ...financials };
            });
            
            program.metadata.monthly_costs_ugx = programMonthlyCostUGX;
            program.metadata.monthly_costs_eur = programMonthlyCostUGX / EXCHANGE_RATE;
        }
    });
    
    // Calculate registry-level metadata
    Object.entries(database.sponsorship_registry).forEach(([programName, registry]) => {
        if (registry.students) {
            registry.metadata = registry.metadata || {};
            registry.metadata.total_students = registry.students.length;
            
            const activeStudents = registry.students.filter(s => 
                s.sponsorship_status === 'active' || !s.sponsorship_status
            );
            registry.metadata.active_students = activeStudents.length;
            totalActiveSponsorships += activeStudents.length;
            
            const inactiveStudents = registry.students.filter(s => 
                s.sponsorship_status === 'inactive'
            );
            registry.metadata.inactive_students = inactiveStudents.length;
            
            const programFunding = activeStudents.reduce((sum, student) => {
                return sum + (parseFloat(student.amount) || 0);
            }, 0);
            
            registry.metadata.total_monthly_funding = programFunding;
            totalMonthlyFundingEUR += programFunding;
        }
    });
    
    // Update global metadata
    database.metadata.programs_summary = {
        total_students_across_all_programs: totalStudents,
        total_active_sponsorships: totalActiveSponsorships,
        total_monthly_funding_euros: totalMonthlyFundingEUR,
        total_monthly_funding_ugx: totalMonthlyFundingEUR * EXCHANGE_RATE
    };
    
    return database;
}

function calculateStudentFinancials(financialData) {
    const result = { ...financialData };
    
    // Ensure all values are numbers
    Object.keys(result).forEach(key => {
        if (typeof result[key] === 'string' && result[key].startsWith('=')) {
            result[key] = evaluateFormula(result[key], result);
        } else if (result[key] === null || result[key] === undefined) {
            result[key] = 0;
        } else if (typeof result[key] === 'string') {
            result[key] = parseFloat(result[key]) || 0;
        }
    });
    
    // Calculate monthly costs in UGX
    const termlyFeesMonthly = (result.termly_school_fees || 0) / 3;
    const directSpending = result.direct_spending_school_fees_ugx_monthly || 0;
    const food = result.food || 0;
    const medical = result.average_medical || 0;
    const transport = result.school_personal_requirements_transport || 0;
    const admin = result.admin_utilities || 0;
    
    result.monthly_output_ugx = termlyFeesMonthly + directSpending + food + medical + transport + admin;
    result.monthly_output_euro = result.monthly_output_ugx / EXCHANGE_RATE;
    
    // Calculate cash received in UGX
    result.cash_received_ugx = (result.cash_received_euro || 0) * EXCHANGE_RATE;
    
    // Calculate balance
    result.plus_minus_diff_ugx = result.cash_received_ugx - result.monthly_output_ugx;
    result.plus_minus_diff_euro = result.plus_minus_diff_ugx / EXCHANGE_RATE;
    
    return result;
}

function evaluateFormula(formula, data) {
    try {
        let expression = formula.substring(1);
        const replacements = {
            '\\$E\\$3': data.termly_school_fees || 0,
            '\\$D\\$3': data.termly_school_fees || 0,
            'D3': data.termly_school_fees || 0,
            'E3': data.direct_spending_school_fees_ugx_monthly || 0,
            'F3': data.direct_spending_school_fees_euros_monthly || 0,
            'G3': data.food || 0,
            'H3': data.average_medical || 0,
            'I3': data.school_personal_requirements_transport || 0,
            'J3': data.admin_utilities || 0
        };
        
        Object.entries(replacements).forEach(([pattern, value]) => {
            expression = expression.replace(new RegExp(pattern, 'g'), value);
        });
        
        expression = expression
            .replace(/(\d+)%/g, '$1/100')
            .replace(/SUM\(([^)]+)\)/g, (match, range) => {
                const parts = range.split(':');
                return parts.reduce((sum, part) => sum + (parseFloat(data[part]) || 0), 0);
            });
        
        const result = Function('"use strict"; return (' + expression + ')')();
        return typeof result === 'number' ? result : 0;
    } catch (error) {
        console.warn('Error evaluating formula:', formula, error);
        return 0;
    }
}

// ===== ENHANCED IMPORT PROCESSING FUNCTIONS =====

function processExcelImport(fileBuffer, importType) {
    try {
        console.log('Processing Excel import for type:', importType);
        
        // Read Excel file
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const result = {
            type: importType,
            data: {},
            metadata: {
                source_file: 'uploaded_file.xlsx',
                import_date: new Date().toISOString(),
                sheets_processed: workbook.SheetNames,
                total_records: 0
            }
        };

        // Process based on import type
        switch (importType) {
            case 'students':
                result.data = processStudentsData(workbook);
                break;
            case 'sponsors':
                result.data = processSponsorsData(workbook);
                break;
            case 'expenses':
                result.data = processExpensesData(workbook);
                break;
            case 'all':
                result.data = processAllData(workbook);
                break;
            default:
                throw new Error(`Unknown import type: ${importType}`);
        }

        // Calculate total records
        result.metadata.total_records = calculateTotalRecords(result.data);
        
        console.log('Import processing completed successfully');
        return result;

    } catch (error) {
        console.error('Error processing Excel import:', error);
        throw new Error(`Failed to process Excel file: ${error.message}`);
    }
}

function processStudentsData(workbook) {
    const programs = ['CH', 'YSP', 'ICCSP', 'OTM_GA'];
    const studentsData = {};

    programs.forEach(program => {
        const sheet = workbook.Sheets[program];
        if (sheet) {
            const jsonData = XLSX.utils.sheet_to_json(sheet);
            console.log(`Processing ${program} students:`, jsonData.length);
            
            studentsData[program] = jsonData.map((row, index) => {
                // Calculate financial data
                const financialData = {
                    termly_school_fees: parseFloat(row['Termly Fees (UGX)']) || 0,
                    direct_spending_school_fees_ugx_monthly: parseFloat(row['Direct Spending (UGX)']) || 0,
                    food: parseFloat(row['Food (UGX)']) || 0,
                    average_medical: parseFloat(row['Medical (UGX)']) || 0,
                    school_personal_requirements_transport: parseFloat(row['Transport (UGX)']) || 0,
                    admin_utilities: parseFloat(row['Admin (UGX)']) || 0,
                    cash_received_euro: parseFloat(row['Cash Received (EUR)']) || 0
                };

                // Calculate derived financial values
                const calculatedFinancials = calculateStudentFinancials(financialData);

                return {
                    serial_number: index + 1,
                    full_name: row['Full Name'] || row['Student Name'] || `Student ${index + 1}`,
                    sponsorship_package: row['Sponsorship Package'] || row['Package'] || 'Day',
                    financial_data: calculatedFinancials,
                    notes: row['Notes'] || ''
                };
            });
        } else {
            console.log(`No sheet found for program: ${program}`);
            studentsData[program] = [];
        }
    });

    return { sponsorship_programs: studentsData };
}

function processSponsorsData(workbook) {
    const programs = ['CH', 'YSP', 'ICCSP', 'OTM_GA'];
    const sponsorsData = {};

    programs.forEach(program => {
        const sheetNames = [
            `${program}_Sponsors`,
            `${program}_sponsors`,
            program
        ];

        let sheet = null;
        for (const sheetName of sheetNames) {
            if (workbook.Sheets[sheetName]) {
                sheet = workbook.Sheets[sheetName];
                break;
            }
        }

        if (sheet) {
            const jsonData = XLSX.utils.sheet_to_json(sheet);
            console.log(`Processing ${program} sponsors:`, jsonData.length);
            
            sponsorsData[program] = jsonData.map((row, index) => ({
                cid: index + 1,
                full_name: row['Sponsored Student'] || row['Student Name'] || '',
                sponsor: row['Sponsor Name'] || row['Sponsor'] || `Sponsor ${index + 1}`,
                amount: parseFloat(row['Amount (EUR)']) || 0,
                sponsorship_status: (row['Status'] || 'active').toLowerCase(),
                category: row['Category'] || 'Individual',
                start_date: formatDate(row['Start Date']) || new Date().toISOString().split('T')[0],
                notes: row['Notes'] || ''
            }));
        } else {
            console.log(`No sponsor sheet found for program: ${program}`);
            sponsorsData[program] = [];
        }
    });

    return { sponsorship_registry: sponsorsData };
}

function processExpensesData(workbook) {
    const sheetNames = ['Expenses', 'expenses', 'Daily Expenses'];
    let sheet = null;

    for (const sheetName of sheetNames) {
        if (workbook.Sheets[sheetName]) {
            sheet = workbook.Sheets[sheetName];
            break;
        }
    }

    if (!sheet) {
        // Try first sheet if no specific expense sheet found
        sheet = workbook.Sheets[workbook.SheetNames[0]];
    }

    const jsonData = XLSX.utils.sheet_to_json(sheet);
    console.log('Processing expenses:', jsonData.length);

    const expenses = jsonData.map((row, index) => ({
        id: index + 1,
        date: formatDate(row['Date']) || new Date().toISOString().split('T')[0],
        studentId: null, // Will be linked during database integration
        studentName: row['Student'] || '',
        category: (row['Category'] || 'other').toLowerCase(),
        amount: parseFloat(row['Amount (UGX)']) || 0,
        description: row['Description'] || '',
        currency: 'UGX'
    }));

    return { daily_expenses: expenses };
}

function processAllData(workbook) {
    const students = processStudentsData(workbook);
    const sponsors = processSponsorsData(workbook);
    const expenses = processExpensesData(workbook);

    return {
        ...students,
        ...sponsors,
        ...expenses
    };
}

function formatDate(dateValue) {
    if (!dateValue) return null;
    
    try {
        if (dateValue instanceof Date) {
            return dateValue.toISOString().split('T')[0];
        }
        
        if (typeof dateValue === 'string') {
            // Try to parse various date formats
            const date = new Date(dateValue);
            if (!isNaN(date.getTime())) {
                return date.toISOString().split('T')[0];
            }
        }
        
        // Handle Excel serial date numbers
        if (typeof dateValue === 'number') {
            const date = XLSX.SSF.parse_date_code(dateValue);
            if (date) {
                return new Date(date.y, date.m - 1, date.d).toISOString().split('T')[0];
            }
        }
        
        return null;
    } catch (error) {
        console.warn('Error formatting date:', dateValue, error);
        return null;
    }
}

function calculateTotalRecords(data) {
    let total = 0;
    
    if (data.sponsorship_programs) {
        Object.values(data.sponsorship_programs).forEach(students => {
            total += students.length;
        });
    }
    
    if (data.sponsorship_registry) {
        Object.values(data.sponsorship_registry).forEach(sponsors => {
            total += sponsors.length;
        });
    }
    
    if (data.daily_expenses) {
        total += data.daily_expenses.length;
    }
    
    return total;
}

// ===== DATABASE INTEGRATION FUNCTIONS =====

function integrateImportedData(database, importedData, importType, mergeStrategy = 'replace') {
    console.log(`Integrating ${importType} data with ${mergeStrategy} strategy`);
    
    const result = JSON.parse(JSON.stringify(database)); // Deep clone
    
    try {
        switch (importType) {
            case 'students':
                integrateStudentsData(result, importedData, mergeStrategy);
                break;
            case 'sponsors':
                integrateSponsorsData(result, importedData, mergeStrategy);
                break;
            case 'expenses':
                integrateExpensesData(result, importedData, mergeStrategy);
                break;
            case 'all':
                integrateAllData(result, importedData, mergeStrategy);
                break;
            default:
                throw new Error(`Unknown import type: ${importType}`);
        }

        // Update metadata and source files
        if (!result.metadata.source_files) {
            result.metadata.source_files = [];
        }
        
        result.metadata.source_files.push({
            filename: importedData.metadata.source_file,
            import_date: importedData.metadata.import_date,
            type: importType,
            records: importedData.metadata.total_records
        });

        // Recalculate all metadata
        return recalculateAllMetadata(result);

    } catch (error) {
        console.error('Error integrating imported data:', error);
        throw new Error(`Failed to integrate data: ${error.message}`);
    }
}

function integrateStudentsData(database, importedData, mergeStrategy) {
    const programs = ['CH', 'YSP', 'ICCSP', 'OTM_GA'];
    
    programs.forEach(program => {
        const importedStudents = importedData.sponsorship_programs[program] || [];
        
        if (mergeStrategy === 'replace') {
            // Replace all students in the program
            database.sponsorship_programs[program].students = importedStudents;
        } else if (mergeStrategy === 'merge') {
            // Merge students, updating existing and adding new
            const existingStudents = database.sponsorship_programs[program].students;
            
            importedStudents.forEach(importedStudent => {
                const existingIndex = existingStudents.findIndex(s => 
                    s.full_name === importedStudent.full_name
                );
                
                if (existingIndex >= 0) {
                    // Update existing student
                    existingStudents[existingIndex] = {
                        ...existingStudents[existingIndex],
                        ...importedStudent,
                        serial_number: existingStudents[existingIndex].serial_number // Keep original serial number
                    };
                } else {
                    // Add new student
                    importedStudent.serial_number = existingStudents.length + 1;
                    existingStudents.push(importedStudent);
                }
            });
        } else if (mergeStrategy === 'append') {
            // Append all new students
            importedStudents.forEach(student => {
                student.serial_number = database.sponsorship_programs[program].students.length + 1;
                database.sponsorship_programs[program].students.push(student);
            });
        }
        
        console.log(`Integrated ${importedStudents.length} students for ${program}`);
    });
}

function integrateSponsorsData(database, importedData, mergeStrategy) {
    const programs = ['CH', 'YSP', 'ICCSP', 'OTM_GA'];
    
    programs.forEach(program => {
        const importedSponsors = importedData.sponsorship_registry[program] || [];
        
        if (mergeStrategy === 'replace') {
            // Replace all sponsors in the program
            database.sponsorship_registry[program].students = importedSponsors;
        } else if (mergeStrategy === 'merge') {
            // Merge sponsors, updating existing and adding new
            const existingSponsors = database.sponsorship_registry[program].students;
            
            importedSponsors.forEach(importedSponsor => {
                const existingIndex = existingSponsors.findIndex(s => 
                    s.full_name === importedSponsor.full_name && s.sponsor === importedSponsor.sponsor
                );
                
                if (existingIndex >= 0) {
                    // Update existing sponsor
                    existingSponsors[existingIndex] = {
                        ...existingSponsors[existingIndex],
                        ...importedSponsor,
                        cid: existingSponsors[existingIndex].cid // Keep original cid
                    };
                } else {
                    // Add new sponsor
                    importedSponsor.cid = existingSponsors.length + 1;
                    existingSponsors.push(importedSponsor);
                }
            });
        } else if (mergeStrategy === 'append') {
            // Append all new sponsors
            importedSponsors.forEach(sponsor => {
                sponsor.cid = database.sponsorship_registry[program].students.length + 1;
                database.sponsorship_registry[program].students.push(sponsor);
            });
        }
        
        console.log(`Integrated ${importedSponsors.length} sponsors for ${program}`);
    });
}

function integrateExpensesData(database, importedData, mergeStrategy) {
    const importedExpenses = importedData.daily_expenses || [];
    
    if (mergeStrategy === 'replace') {
        // Replace all expenses
        database.daily_expenses = importedExpenses;
    } else if (mergeStrategy === 'merge') {
        // Merge expenses based on date, student, and description
        const existingExpenses = database.daily_expenses;
        
        importedExpenses.forEach(importedExpense => {
            const existingIndex = existingExpenses.findIndex(e => 
                e.date === importedExpense.date &&
                e.studentName === importedExpense.studentName &&
                e.description === importedExpense.description
            );
            
            if (existingIndex >= 0) {
                // Update existing expense
                existingExpenses[existingIndex] = {
                    ...existingExpenses[existingIndex],
                    ...importedExpense,
                    id: existingExpenses[existingIndex].id // Keep original id
                };
            } else {
                // Add new expense
                importedExpense.id = existingExpenses.length > 0 ? 
                    Math.max(...existingExpenses.map(e => e.id)) + 1 : 1;
                existingExpenses.push(importedExpense);
            }
        });
    } else if (mergeStrategy === 'append') {
        // Append all new expenses
        const maxId = database.daily_expenses.length > 0 ? 
            Math.max(...database.daily_expenses.map(e => e.id)) : 0;
        
        importedExpenses.forEach((expense, index) => {
            expense.id = maxId + index + 1;
            database.daily_expenses.push(expense);
        });
    }
    
    console.log(`Integrated ${importedExpenses.length} expenses`);
}

function integrateAllData(database, importedData, mergeStrategy) {
    integrateStudentsData(database, importedData, mergeStrategy);
    integrateSponsorsData(database, importedData, mergeStrategy);
    integrateExpensesData(database, importedData, mergeStrategy);
}

// ===== ROUTES =====

// Serve main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ===== CORE DATA ROUTES =====

// Get complete database
app.get('/api/data', (req, res) => {
    const database = loadDatabase();
    res.json({
        success: true,
        data: database,
        message: 'Data loaded successfully'
    });
});

// Get financial summary
app.get('/api/financial-summary', (req, res) => {
    const database = loadDatabase();
    const summary = calculateFinancialSummary(database);
    res.json({
        success: true,
        data: summary,
        message: 'Financial summary loaded successfully'
    });
});

// Get funding gap analysis
app.get('/api/funding-gap', (req, res) => {
    const database = loadDatabase();
    const gap = calculateFundingGap(database);
    res.json({
        success: true,
        data: gap,
        message: 'Funding gap analysis loaded successfully'
    });
});

// Get specific program
app.get('/api/programs/:program', (req, res) => {
    const program = req.params.program;
    const database = loadDatabase();
    const programData = database.sponsorship_programs[program] || {};
    res.json({
        success: true,
        data: programData,
        message: `Program ${program} loaded successfully`
    });
});

// Create or update program
app.post('/api/programs/:program', (req, res) => {
    const program = req.params.program;
    const programData = req.body;
    
    let database = loadDatabase();
    
    if (!database.sponsorship_programs[program]) {
        database.sponsorship_programs[program] = {
            program_name: programData.program_name || `${program} FINANCIAL ANALYSIS REPORT`,
            students: [],
            metadata: { total_students: 0, sponsorship_types: {} }
        };
    }
    
    // Update program data
    Object.assign(database.sponsorship_programs[program], programData);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'program_updated',
            program: program,
            data: updatedDatabase.sponsorship_programs[program],
            database: updatedDatabase,
            message: `Program ${program} updated successfully`
        });
        res.json({ 
            success: true, 
            data: updatedDatabase.sponsorship_programs[program],
            message: `Program ${program} updated successfully` 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save database' });
    }
});

// ===== STUDENT MANAGEMENT ROUTES =====

// Add student to program
app.post('/api/programs/:program/students', (req, res) => {
    const program = req.params.program;
    const studentData = req.body;
    
    let database = loadDatabase();
    
    if (!database.sponsorship_programs[program]) {
        return res.status(404).json({ success: false, message: `Program ${program} not found` });
    }
    
    // Generate serial number if not provided
    if (!studentData.serial_number) {
        studentData.serial_number = database.sponsorship_programs[program].students.length + 1;
    }
    
    // Calculate financial data
    const calculatedFinancials = calculateStudentFinancials(studentData.financial_data || {});
    studentData.financial_data = calculatedFinancials;
    
    database.sponsorship_programs[program].students.push(studentData);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'student_added',
            program: program,
            student: studentData,
            database: updatedDatabase,
            message: `Student added to ${program} successfully`
        });
        res.json({ 
            success: true, 
            data: studentData,
            database: updatedDatabase,
            message: 'Student added successfully'
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save database' });
    }
});

// Update student
app.put('/api/programs/:program/students/:studentId', (req, res) => {
    const program = req.params.program;
    const studentId = parseInt(req.params.studentId);
    const updates = req.body;
    
    let database = loadDatabase();
    
    if (!database.sponsorship_programs[program]) {
        return res.status(404).json({ success: false, message: `Program ${program} not found` });
    }
    
    const studentIndex = database.sponsorship_programs[program].students.findIndex(
        s => s.serial_number === studentId
    );
    
    if (studentIndex === -1) {
        return res.status(404).json({ success: false, message: 'Student not found' });
    }
    
    // Update student data and recalculate financials
    if (updates.financial_data) {
        updates.financial_data = calculateStudentFinancials(updates.financial_data);
    }
    
    Object.assign(database.sponsorship_programs[program].students[studentIndex], updates);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'student_updated',
            program: program,
            studentId: studentId,
            student: database.sponsorship_programs[program].students[studentIndex],
            database: updatedDatabase,
            message: `Student updated successfully`
        });
        res.json({ 
            success: true, 
            data: database.sponsorship_programs[program].students[studentIndex],
            message: 'Student updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save database' });
    }
});

// Delete student
app.delete('/api/programs/:program/students/:studentId', (req, res) => {
    const program = req.params.program;
    const studentId = parseInt(req.params.studentId);
    
    let database = loadDatabase();
    
    if (!database.sponsorship_programs[program]) {
        return res.status(404).json({ success: false, message: `Program ${program} not found` });
    }
    
    const studentIndex = database.sponsorship_programs[program].students.findIndex(
        s => s.serial_number === studentId
    );
    
    if (studentIndex === -1) {
        return res.status(404).json({ success: false, message: 'Student not found' });
    }
    
    const deletedStudent = database.sponsorship_programs[program].students.splice(studentIndex, 1)[0];
    
    // Reassign serial numbers
    database.sponsorship_programs[program].students.forEach((student, index) => {
        student.serial_number = index + 1;
    });
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'student_deleted',
            program: program,
            studentId: studentId,
            database: updatedDatabase,
            message: `Student deleted successfully`
        });
        res.json({ 
            success: true, 
            database: updatedDatabase,
            message: 'Student deleted successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save database' });
    }
});

// Get all students for a program (for sponsor dropdown)
app.get('/api/programs/:program/students-list', (req, res) => {
    const program = req.params.program;
    const database = loadDatabase();
    const programData = database.sponsorship_programs[program] || {};
    
    const studentsList = programData.students ? programData.students.map(student => ({
        serial_number: student.serial_number,
        full_name: student.full_name,
        sponsorship_package: student.sponsorship_package
    })) : [];
    
    res.json({
        success: true,
        data: studentsList,
        message: `Students list for ${program} loaded successfully`
    });
});

// ===== SPONSORSHIP REGISTRY ROUTES =====

// Get all sponsors for a program
app.get('/api/registry/:program', (req, res) => {
    const program = req.params.program;
    const database = loadDatabase();
    const registryData = database.sponsorship_registry[program] || {};
    
    res.json({
        success: true,
        data: registryData,
        message: `Registry data for ${program} loaded successfully`
    });
});

// Add sponsor to registry
app.post('/api/registry/:program/students', (req, res) => {
    const program = req.params.program;
    const sponsorData = req.body;
    
    let database = loadDatabase();
    
    if (!database.sponsorship_registry[program]) {
        database.sponsorship_registry[program] = {
            students: [],
            metadata: { total_students: 0, active_students: 0, total_monthly_funding: 0 }
        };
    }
    
    // Generate CID if not provided
    if (!sponsorData.cid) {
        sponsorData.cid = database.sponsorship_registry[program].students.length + 1;
    }
    
    // Validate that the student exists in the program
    const programStudents = database.sponsorship_programs[program]?.students || [];
    const studentExists = programStudents.some(student => 
        student.full_name === sponsorData.full_name
    );
    
    if (!studentExists) {
        return res.status(400).json({ 
            success: false, 
            message: `Student "${sponsorData.full_name}" not found in program ${program}` 
        });
    }
    
    // Set default values
    if (!sponsorData.sponsorship_status) {
        sponsorData.sponsorship_status = 'active';
    }
    if (!sponsorData.category) {
        sponsorData.category = 'Individual';
    }
    
    database.sponsorship_registry[program].students.push(sponsorData);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'sponsor_added',
            program: program,
            sponsor: sponsorData,
            database: updatedDatabase,
            message: `Sponsor added to ${program} registry successfully`
        });
        res.json({ 
            success: true, 
            data: sponsorData,
            database: updatedDatabase,
            message: 'Sponsor added successfully'
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save database' });
    }
});

// Update sponsor in registry
app.put('/api/registry/:program/students/:sponsorId', (req, res) => {
    const program = req.params.program;
    const sponsorId = parseInt(req.params.sponsorId);
    const updates = req.body;
    
    let database = loadDatabase();
    
    if (!database.sponsorship_registry[program]) {
        return res.status(404).json({ success: false, message: `Program ${program} not found in registry` });
    }
    
    const sponsorIndex = database.sponsorship_registry[program].students.findIndex(
        s => s.cid === sponsorId
    );
    
    if (sponsorIndex === -1) {
        return res.status(404).json({ success: false, message: 'Sponsor not found' });
    }
    
    // If student name is being updated, validate it exists in the program
    if (updates.full_name) {
        const programStudents = database.sponsorship_programs[program]?.students || [];
        const studentExists = programStudents.some(student => 
            student.full_name === updates.full_name
        );
        
        if (!studentExists) {
            return res.status(400).json({ 
                success: false, 
                message: `Student "${updates.full_name}" not found in program ${program}` 
            });
        }
    }
    
    // Update sponsor data
    Object.assign(database.sponsorship_registry[program].students[sponsorIndex], updates);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'sponsor_updated',
            program: program,
            sponsorId: sponsorId,
            sponsor: database.sponsorship_registry[program].students[sponsorIndex],
            database: updatedDatabase,
            message: `Sponsor updated successfully`
        });
        res.json({ 
            success: true, 
            data: database.sponsorship_registry[program].students[sponsorIndex],
            database: updatedDatabase,
            message: 'Sponsor updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save database' });
    }
});

// Delete sponsor from registry
app.delete('/api/registry/:program/students/:sponsorId', (req, res) => {
    const program = req.params.program;
    const sponsorId = parseInt(req.params.sponsorId);
    
    let database = loadDatabase();
    
    if (!database.sponsorship_registry[program]) {
        return res.status(404).json({ success: false, message: `Program ${program} not found in registry` });
    }
    
    const sponsorIndex = database.sponsorship_registry[program].students.findIndex(
        s => s.cid === sponsorId
    );
    
    if (sponsorIndex === -1) {
        return res.status(404).json({ success: false, message: 'Sponsor not found' });
    }
    
    const deletedSponsor = database.sponsorship_registry[program].students.splice(sponsorIndex, 1)[0];
    
    // Reassign CID numbers
    database.sponsorship_registry[program].students.forEach((sponsor, index) => {
        sponsor.cid = index + 1;
    });
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'sponsor_deleted',
            program: program,
            sponsorId: sponsorId,
            database: updatedDatabase,
            message: `Sponsor deleted successfully`
        });
        res.json({ 
            success: true, 
            database: updatedDatabase,
            message: 'Sponsor deleted successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save database' });
    }
});

// Get all sponsors (across all programs)
app.get('/api/registry', (req, res) => {
    const database = loadDatabase();
    const allSponsors = [];
    
    Object.entries(database.sponsorship_registry).forEach(([program, registry]) => {
        if (registry.students) {
            registry.students.forEach(sponsor => {
                allSponsors.push({
                    ...sponsor,
                    program: program
                });
            });
        }
    });
    
    res.json({
        success: true,
        data: allSponsors,
        message: 'All sponsors loaded successfully'
    });
});

// Get sponsor statistics
app.get('/api/registry-stats', (req, res) => {
    const database = loadDatabase();
    const stats = calculateSponsorStatistics(database);
    
    res.json({
        success: true,
        data: stats,
        message: 'Sponsor statistics loaded successfully'
    });
});

// ===== DAILY EXPENSES ROUTES =====

// Get all daily expenses
app.get('/api/expenses', (req, res) => {
    const database = loadDatabase();
    res.json({
        success: true,
        data: database.daily_expenses || [],
        message: 'Daily expenses loaded successfully'
    });
});

// Add new expense
app.post('/api/expenses', (req, res) => {
    const expenseData = req.body;
    
    let database = loadDatabase();
    
    if (!database.daily_expenses) {
        database.daily_expenses = [];
    }
    
    // Generate ID if not provided
    if (!expenseData.id) {
        expenseData.id = database.daily_expenses.length > 0 ? 
            Math.max(...database.daily_expenses.map(e => e.id)) + 1 : 1;
    }
    
    // Set default currency if not provided
    if (!expenseData.currency) {
        expenseData.currency = 'UGX';
    }
    
    database.daily_expenses.push(expenseData);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'expense_added',
            expense: expenseData,
            database: updatedDatabase,
            message: 'Expense added successfully'
        });
        res.json({ 
            success: true, 
            data: expenseData,
            message: 'Expense added successfully'
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save expense' });
    }
});

// Update expense
app.put('/api/expenses/:expenseId', (req, res) => {
    const expenseId = parseInt(req.params.expenseId);
    const updates = req.body;
    
    let database = loadDatabase();
    
    if (!database.daily_expenses) {
        return res.status(404).json({ success: false, message: 'No expenses found' });
    }
    
    const expenseIndex = database.daily_expenses.findIndex(e => e.id === expenseId);
    
    if (expenseIndex === -1) {
        return res.status(404).json({ success: false, message: 'Expense not found' });
    }
    
    Object.assign(database.daily_expenses[expenseIndex], updates);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'expense_updated',
            expenseId: expenseId,
            expense: database.daily_expenses[expenseIndex],
            database: updatedDatabase,
            message: 'Expense updated successfully'
        });
        res.json({ 
            success: true, 
            data: database.daily_expenses[expenseIndex],
            message: 'Expense updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save expense' });
    }
});

// Delete expense
app.delete('/api/expenses/:expenseId', (req, res) => {
    const expenseId = parseInt(req.params.expenseId);
    
    let database = loadDatabase();
    
    if (!database.daily_expenses) {
        return res.status(404).json({ success: false, message: 'No expenses found' });
    }
    
    const expenseIndex = database.daily_expenses.findIndex(e => e.id === expenseId);
    
    if (expenseIndex === -1) {
        return res.status(404).json({ success: false, message: 'Expense not found' });
    }
    
    const deletedExpense = database.daily_expenses.splice(expenseIndex, 1)[0];
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'expense_deleted',
            expenseId: expenseId,
            database: updatedDatabase,
            message: 'Expense deleted successfully'
        });
        res.json({ 
            success: true, 
            message: 'Expense deleted successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to delete expense' });
    }
});

// ===== EVENTS MANAGEMENT ROUTES =====

// Get all events
app.get('/api/events', (req, res) => {
    const database = loadDatabase();
    res.json({
        success: true,
        data: database.events || [],
        message: 'Events loaded successfully'
    });
});

// Add new event
app.post('/api/events', (req, res) => {
    const eventData = req.body;
    
    let database = loadDatabase();
    
    if (!database.events) {
        database.events = [];
    }
    
    // Generate ID if not provided
    if (!eventData.id) {
        eventData.id = database.events.length > 0 ? 
            Math.max(...database.events.map(e => e.id)) + 1 : 1;
    }
    
    database.events.push(eventData);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'event_added',
            event: eventData,
            database: updatedDatabase,
            message: 'Event added successfully'
        });
        res.json({ 
            success: true, 
            data: eventData,
            message: 'Event added successfully'
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save event' });
    }
});

// Update event
app.put('/api/events/:eventId', (req, res) => {
    const eventId = parseInt(req.params.eventId);
    const updates = req.body;
    
    let database = loadDatabase();
    
    if (!database.events) {
        return res.status(404).json({ success: false, message: 'No events found' });
    }
    
    const eventIndex = database.events.findIndex(e => e.id === eventId);
    
    if (eventIndex === -1) {
        return res.status(404).json({ success: false, message: 'Event not found' });
    }
    
    Object.assign(database.events[eventIndex], updates);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'event_updated',
            eventId: eventId,
            event: database.events[eventIndex],
            database: updatedDatabase,
            message: 'Event updated successfully'
        });
        res.json({ 
            success: true, 
            data: database.events[eventIndex],
            message: 'Event updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save event' });
    }
});

// Delete event
app.delete('/api/events/:eventId', (req, res) => {
    const eventId = parseInt(req.params.eventId);
    
    let database = loadDatabase();
    
    if (!database.events) {
        return res.status(404).json({ success: false, message: 'No events found' });
    }
    
    const eventIndex = database.events.findIndex(e => e.id === eventId);
    
    if (eventIndex === -1) {
        return res.status(404).json({ success: false, message: 'Event not found' });
    }
    
    const deletedEvent = database.events.splice(eventIndex, 1)[0];
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'event_deleted',
            eventId: eventId,
            database: updatedDatabase,
            message: 'Event deleted successfully'
        });
        res.json({ 
            success: true, 
            message: 'Event deleted successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to delete event' });
    }
});

// ===== SETTINGS MANAGEMENT ROUTES =====

// Get system settings
app.get('/api/settings', (req, res) => {
    const database = loadDatabase();
    res.json({
        success: true,
        data: database.system_settings || {},
        message: 'Settings loaded successfully'
    });
});

// Update system settings
app.put('/api/settings', (req, res) => {
    const settings = req.body;
    
    let database = loadDatabase();
    
    if (!database.system_settings) {
        database.system_settings = {};
    }
    
    Object.assign(database.system_settings, settings);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'settings_updated',
            settings: database.system_settings,
            database: updatedDatabase,
            message: 'Settings updated successfully'
        });
        res.json({ 
            success: true, 
            data: database.system_settings,
            message: 'Settings updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save settings' });
    }
});

// Update organization settings
app.put('/api/settings/organization', (req, res) => {
    const orgSettings = req.body;
    
    let database = loadDatabase();
    
    if (!database.system_settings) {
        database.system_settings = {};
    }
    
    if (!database.system_settings.organization) {
        database.system_settings.organization = {};
    }
    
    Object.assign(database.system_settings.organization, orgSettings);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'organization_settings_updated',
            settings: database.system_settings.organization,
            database: updatedDatabase,
            message: 'Organization settings updated successfully'
        });
        res.json({ 
            success: true, 
            data: database.system_settings.organization,
            message: 'Organization settings updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save organization settings' });
    }
});

// Update forex settings
app.put('/api/settings/forex', (req, res) => {
    const forexSettings = req.body;
    
    let database = loadDatabase();
    
    if (!database.system_settings) {
        database.system_settings = {};
    }
    
    if (!database.system_settings.forex) {
        database.system_settings.forex = {};
    }
    
    Object.assign(database.system_settings.forex, forexSettings);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'forex_settings_updated',
            settings: database.system_settings.forex,
            database: updatedDatabase,
            message: 'Forex settings updated successfully'
        });
        res.json({ 
            success: true, 
            data: database.system_settings.forex,
            message: 'Forex settings updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save forex settings' });
    }
});

// Update notification settings
app.put('/api/settings/notifications', (req, res) => {
    const notificationSettings = req.body;
    
    let database = loadDatabase();
    
    if (!database.system_settings) {
        database.system_settings = {};
    }
    
    if (!database.system_settings.notifications) {
        database.system_settings.notifications = {};
    }
    
    Object.assign(database.system_settings.notifications, notificationSettings);
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'notification_settings_updated',
            settings: database.system_settings.notifications,
            database: updatedDatabase,
            message: 'Notification settings updated successfully'
        });
        res.json({ 
            success: true, 
            data: database.system_settings.notifications,
            message: 'Notification settings updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save notification settings' });
    }
});

// Upload logo
app.post('/api/settings/logo', (req, res) => {
    const { logoData } = req.body;
    
    if (!logoData) {
        return res.status(400).json({ success: false, message: 'No logo data provided' });
    }
    
    let database = loadDatabase();
    
    if (!database.system_settings) {
        database.system_settings = {};
    }
    
    if (!database.system_settings.organization) {
        database.system_settings.organization = {};
    }
    
    database.system_settings.organization.logo = logoData;
    
    if (saveDatabase(database)) {
        const updatedDatabase = loadDatabase();
        broadcast({
            type: 'logo_updated',
            logo: logoData,
            database: updatedDatabase,
            message: 'Logo updated successfully'
        });
        res.json({ 
            success: true, 
            message: 'Logo updated successfully' 
        });
    } else {
        res.status(500).json({ success: false, message: 'Failed to save logo' });
    }
});

// ===== ANALYTICS AND REPORTS =====

// Get analytics data
app.get('/api/analytics', (req, res) => {
    const database = loadDatabase();
    const analytics = calculateAnalytics(database);
    
    res.json({
        success: true,
        data: analytics,
        message: 'Analytics data loaded successfully'
    });
});

// Generate financial report
app.get('/api/reports/financial', (req, res) => {
    const database = loadDatabase();
    const report = generateFinancialReport(database);
    
    res.json({
        success: true,
        data: report,
        message: 'Financial report generated successfully'
    });
});

// ===== ENHANCED IMPORT/EXPORT ROUTES =====

// Export data to JSON
app.get('/api/export/json', (req, res) => {
    const database = loadDatabase();
    const timestamp = new Date().toISOString().split('T')[0];
    
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Content-Disposition', `attachment; filename=sponsorship_data_${timestamp}.json`);
    res.send(JSON.stringify(database, null, 2));
});

// Import data from JSON
app.post('/api/import/json', (req, res) => {
    try {
        const importedData = req.body;
        
        if (!importedData || typeof importedData !== 'object') {
            return res.status(400).json({ success: false, message: 'Invalid data format' });
        }
        
        console.log('Processing JSON import:', importedData.type);
        
        let database = loadDatabase();
        
        // Integrate imported data
        const mergeStrategy = req.query.merge || 'replace';
        database = integrateImportedData(database, importedData, importedData.type, mergeStrategy);
        
        if (saveDatabase(database)) {
            const updatedDatabase = loadDatabase();
            broadcast({
                type: 'database_imported',
                database: updatedDatabase,
                message: 'Data imported successfully'
            });
            res.json({ 
                success: true, 
                message: 'Data imported successfully',
                records: importedData.metadata.total_records
            });
        } else {
            res.status(500).json({ success: false, message: 'Failed to save imported data' });
        }
    } catch (error) {
        console.error('Import error:', error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// Enhanced Excel import endpoint
app.post('/api/import/excel', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, message: 'No file uploaded' });
        }
        
        const importType = req.body.type || 'all';
        const mergeStrategy = req.body.merge || 'replace';
        
        console.log(`Processing Excel import: ${importType} with ${mergeStrategy} strategy`);
        
        // Process Excel file
        const importedData = processExcelImport(req.file.buffer, importType);
        
        // Load current database
        let database = loadDatabase();
        
        // Integrate imported data
        database = integrateImportedData(database, importedData, importType, mergeStrategy);
        
        // Save updated database
        if (saveDatabase(database)) {
            const updatedDatabase = loadDatabase();
            broadcast({
                type: 'database_imported',
                database: updatedDatabase,
                message: 'Excel data imported successfully'
            });
            
            res.json({
                success: true,
                message: 'Excel file imported successfully',
                records: importedData.metadata.total_records,
                type: importType,
                merge_strategy: mergeStrategy
            });
        } else {
            res.status(500).json({ success: false, message: 'Failed to save imported data' });
        }
        
    } catch (error) {
        console.error('Excel import error:', error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// Generate sample Excel templates
app.get('/api/export/template/:type', (req, res) => {
    try {
        const type = req.params.type;
        const wb = XLSX.utils.book_new();
        
        switch (type) {
            case 'students':
                generateStudentsTemplate(wb);
                break;
            case 'sponsors':
                generateSponsorsTemplate(wb);
                break;
            case 'expenses':
                generateExpensesTemplate(wb);
                break;
            case 'all':
                generateStudentsTemplate(wb);
                generateSponsorsTemplate(wb);
                generateExpensesTemplate(wb);
                break;
            default:
                return res.status(400).json({ success: false, message: 'Invalid template type' });
        }
        
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=sponsorship_${type}_template.xlsx`);
        res.send(buffer);
        
    } catch (error) {
        console.error('Template generation error:', error);
        res.status(500).json({ success: false, message: 'Failed to generate template' });
    }
});

// ===== HELPER FUNCTIONS =====

function calculateSponsorStatistics(database) {
    let totalSponsors = 0;
    let activeSponsors = 0;
    let totalMonthlyFunding = 0;
    const programStats = {};
    
    Object.entries(database.sponsorship_registry).forEach(([program, registry]) => {
        if (registry.students) {
            const programSponsors = registry.students.length;
            totalSponsors += programSponsors;
            
            const programActiveSponsors = registry.students.filter(s => 
                s.sponsorship_status === 'active' || !s.sponsorship_status
            ).length;
            activeSponsors += programActiveSponsors;
            
            const programFunding = registry.students
                .filter(s => s.sponsorship_status === 'active' || !s.sponsorship_status)
                .reduce((sum, student) => sum + (parseFloat(student.amount) || 0), 0);
            totalMonthlyFunding += programFunding;
            
            programStats[program] = {
                total_sponsors: programSponsors,
                active_sponsors: programActiveSponsors,
                monthly_funding: programFunding
            };
        }
    });
    
    return {
        total_sponsors: totalSponsors,
        active_sponsors: activeSponsors,
        total_monthly_funding: totalMonthlyFunding,
        program_stats: programStats
    };
}

function calculateFinancialSummary(db) {
    const totalIncomeEUR = db.metadata.programs_summary?.total_monthly_funding_euros || 0;
    const totalIncomeUGX = totalIncomeEUR * EXCHANGE_RATE;
    
    // Calculate total costs from all programs
    let totalCostsUGX = 0;
    Object.values(db.sponsorship_programs).forEach(program => {
        totalCostsUGX += program.metadata?.monthly_costs_ugx || 0;
    });
    const totalCostsEUR = totalCostsUGX / EXCHANGE_RATE;
    
    const totalDeficitUGX = totalIncomeUGX - totalCostsUGX;
    const totalDeficitEUR = totalDeficitUGX / EXCHANGE_RATE;
    
    return {
        totalStudents: db.metadata.programs_summary?.total_students_across_all_programs || 0,
        totalIncomeEUR,
        totalIncomeUGX,
        totalCostsEUR,
        totalCostsUGX,
        totalDeficitEUR,
        totalDeficitUGX
    };
}

function calculateFundingGap(db) {
    const programs = Object.keys(db.sponsorship_programs);
    const gaps = [];
    
    programs.forEach(program => {
        const programData = db.sponsorship_programs[program];
        const registryData = db.sponsorship_registry[program];
        
        if (programData) {
            const incomeEUR = registryData?.metadata?.total_monthly_funding || 0;
            const incomeUGX = incomeEUR * EXCHANGE_RATE;
            
            const costsUGX = programData.metadata?.monthly_costs_ugx || 0;
            const costsEUR = costsUGX / EXCHANGE_RATE;
            
            const deficitUGX = incomeUGX - costsUGX;
            const deficitEUR = deficitUGX / EXCHANGE_RATE;
            
            gaps.push({
                name: program,
                incomeEUR,
                incomeUGX,
                costsEUR,
                costsUGX,
                deficitEUR,
                deficitUGX,
                studentCount: programData.metadata?.total_students || 0
            });
        }
    });
    
    const totalDeficitUGX = gaps.reduce((sum, gap) => sum + gap.deficitUGX, 0);
    const totalDeficitEUR = totalDeficitUGX / EXCHANGE_RATE;
    
    return {
        programs: gaps,
        totalDeficitEUR,
        totalDeficitUGX
    };
}

function calculateAnalytics(db) {
    const programs = Object.keys(db.sponsorship_programs);
    const analytics = {
        program_distribution: {},
        expense_breakdown: {},
        financial_trends: {},
        sponsor_categories: {}
    };
    
    // Program distribution
    programs.forEach(program => {
        const programData = db.sponsorship_programs[program];
        if (programData && programData.metadata) {
            analytics.program_distribution[program] = {
                student_count: programData.metadata.total_students || 0,
                monthly_costs: programData.metadata.monthly_costs_ugx || 0,
                sponsorship_types: programData.metadata.sponsorship_types || {}
            };
        }
    });
    
    // Expense breakdown
    if (db.daily_expenses) {
        const categories = {};
        db.daily_expenses.forEach(expense => {
            const category = expense.category || 'other';
            categories[category] = (categories[category] || 0) + (expense.amount || 0);
        });
        analytics.expense_breakdown = categories;
    }
    
    // Sponsor categories
    Object.entries(db.sponsorship_registry).forEach(([program, registry]) => {
        if (registry.students) {
            registry.students.forEach(sponsor => {
                const category = sponsor.category || 'Individual';
                analytics.sponsor_categories[category] = (analytics.sponsor_categories[category] || 0) + 1;
            });
        }
    });
    
    return analytics;
}

function generateFinancialReport(db) {
    const financialSummary = calculateFinancialSummary(db);
    const fundingGap = calculateFundingGap(db);
    const sponsorStats = calculateSponsorStatistics(db);
    
    return {
        summary: financialSummary,
        funding_gap: fundingGap,
        sponsor_statistics: sponsorStats,
        generated_at: new Date().toISOString(),
        exchange_rate: EXCHANGE_RATE
    };
}

// Template generation functions
function generateStudentsTemplate(wb) {
    const programs = ['CH', 'YSP', 'ICCSP', 'OTM_GA'];
    
    programs.forEach(program => {
        const sampleData = [
            {
                'Full Name': 'Example Student 1',
                'Sponsorship Package': 'Day',
                'Termly Fees (UGX)': 208000,
                'Direct Spending (UGX)': 0,
                'Food (UGX)': 100000,
                'Medical (UGX)': 20000,
                'Transport (UGX)': 20000,
                'Admin (UGX)': 13500,
                'Cash Received (EUR)': 70,
                'Notes': 'Sample student record'
            },
            {
                'Full Name': 'Example Student 2',
                'Sponsorship Package': 'Boarding',
                'Termly Fees (UGX)': 557000,
                'Direct Spending (UGX)': 0,
                'Food (UGX)': 0,
                'Medical (UGX)': 0,
                'Transport (UGX)': 0,
                'Admin (UGX)': 0,
                'Cash Received (EUR)': 70,
                'Notes': 'Another sample record'
            }
        ];
        
        const ws = XLSX.utils.json_to_sheet(sampleData);
        XLSX.utils.book_append_sheet(wb, ws, program);
    });
}

function generateSponsorsTemplate(wb) {
    const programs = ['CH', 'YSP', 'ICCSP', 'OTM_GA'];
    
    programs.forEach(program => {
        const sampleData = [
            {
                'Sponsor Name': 'Example Sponsor 1',
                'Sponsored Student': 'Example Student 1',
                'Amount (EUR)': 70,
                'Status': 'active',
                'Category': 'Individual',
                'Start Date': new Date().toISOString().split('T')[0],
                'Notes': 'Sample sponsor record'
            },
            {
                'Sponsor Name': 'Example Sponsor 2',
                'Sponsored Student': 'Example Student 2',
                'Amount (EUR)': 100,
                'Status': 'active',
                'Category': 'Organization',
                'Start Date': new Date().toISOString().split('T')[0],
                'Notes': 'Another sample record'
            }
        ];
        
        const ws = XLSX.utils.json_to_sheet(sampleData);
        XLSX.utils.book_append_sheet(wb, ws, `${program}_Sponsors`);
    });
}

function generateExpensesTemplate(wb) {
    const sampleData = [
        {
            'Date': new Date().toISOString().split('T')[0],
            'Student': 'Example Student 1',
            'Category': 'school',
            'Description': 'School fees payment',
            'Amount (UGX)': 50000
        },
        {
            'Date': new Date().toISOString().split('T')[0],
            'Student': 'Example Student 2',
            'Category': 'medical',
            'Description': 'Medical treatment',
            'Amount (UGX)': 25000
        },
        {
            'Date': new Date().toISOString().split('T')[0],
            'Student': 'Example Student 1',
            'Category': 'transport',
            'Description': 'Transport to school',
            'Amount (UGX)': 15000
        }
    ];
    
    const ws = XLSX.utils.json_to_sheet(sampleData);
    XLSX.utils.book_append_sheet(wb, ws, 'Expenses');
}

// Error handling middleware
app.use((error, req, res, next) => {
    console.error('Server error:', error);
    res.status(500).json({
        success: false,
        message: error.message || 'Internal server error'
    });
});

// 404 handler
app.use((req, res) => {
    res.status(404).json({
        success: false,
        message: 'Endpoint not found'
    });
});

// Start server
server.listen(PORT, () => {
    console.log('🚀 Sponsorship Management System with Enhanced Import Functionality');
    console.log(`📍 Local: http://localhost:${PORT}`);
    console.log(`💾 Database: ${DB_FILE}`);
    console.log(`🔗 WebSockets: Enabled for real-time updates`);
    console.log(`💰 Exchange Rate: 1 EUR = ${EXCHANGE_RATE} UGX`);
    console.log(`📊 Import Features: Excel file processing, template generation`);
    console.log('✅ Server running successfully!');
    
    // Initialize database file if it doesn't exist
    loadDatabase();
});
