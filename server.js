const express = require('express');
const http = require('http');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs');

const app = express();
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

const PORT = 3000;
const DB_FILE = 'database.json';
const EXCHANGE_RATE = 4100; // 1 EUR = 4100 UGX

// Middleware
app.use(express.json());
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

// Routes

// Serve main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

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

// Helper function to update registry metadata
function updateRegistryMetadata(program, database) {
    const registry = database.sponsorship_registry[program];
    if (registry && registry.students) {
        registry.metadata = registry.metadata || {};
        registry.metadata.total_students = registry.students.length;
        
        const activeStudents = registry.students.filter(s => 
            s.sponsorship_status === 'active' || !s.sponsorship_status
        );
        registry.metadata.active_students = activeStudents.length;
        
        const inactiveStudents = registry.students.filter(s => 
            s.sponsorship_status === 'inactive'
        );
        registry.metadata.inactive_students = inactiveStudents.length;
        
        const programFunding = activeStudents.reduce((sum, student) => {
            return sum + (parseFloat(student.amount) || 0);
        }, 0);
        
        registry.metadata.total_monthly_funding = programFunding;
    }
    return database;
}

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

// Financial calculation functions
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

// Start server
server.listen(PORT, () => {
    console.log('🚀 Sponsorship Management System with WebSockets');
    console.log(`📍 Local: http://localhost:${PORT}`);
    console.log(`💾 Database: ${DB_FILE}`);
    console.log(`🔗 WebSockets: Enabled for real-time updates`);
    console.log(`💰 Exchange Rate: 1 EUR = ${EXCHANGE_RATE} UGX`);
    console.log('✅ Server running successfully!');
    
    // Initialize database file if it doesn't exist
    loadDatabase();
});