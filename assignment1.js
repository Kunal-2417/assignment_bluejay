const fs = require('fs');
const xlsx = require('xlsx');

const filePath = 'C:\\Users\\ashish chaudhary\\Downloads\\Assignment_Timecard.xlsx';

function analyzeEmployeeSchedule(filePath) {
    // Read the Excel file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert Excel data to JSON
    const jsonData = xlsx.utils.sheet_to_json(sheet);

    // Process the data
    const employees = {};

    jsonData.forEach(entry => {
        const employee = entry['Employee Name'];
        const position = entry['Position ID'];
        const shiftStart = new Date(entry['Time Out'] * 24 * 60 * 60 * 1000);
        const shiftEnd = new Date(entry['Time'] * 24 * 60 * 60 * 1000);
        const shiftDuration = (shiftEnd - shiftStart) / (1000 * 60 * 60); // in hours

        if (!employees[employee]) {
            employees[employee] = { position, shifts: [] };
        }

        employees[employee].shifts.push({ shiftStart, shiftEnd, shiftDuration });
    });

    // Analyze the data
    const sevenConsecutiveDaysEmployees = [];
    const betweenShiftsEmployees = [];
    const longShiftEmployees = [];

    for (const employee in employees) {
        const shifts = employees[employee].shifts;

        // a) Employees who have worked for 7 consecutive days
        const consecutiveDaysShifts = shifts.filter((shift, index) =>
            index < shifts.length - 1 &&
            (shifts[index + 1].shiftStart - shift.shiftEnd) / (1000 * 60 * 60 * 24) === 1
        );

        if (consecutiveDaysShifts.length >= 6) {
            sevenConsecutiveDaysEmployees.push({ employee, position: employees[employee].position });
        }

        // b) Employees with less than 10 hours between shifts but greater than 1 hour
        const betweenShifts = shifts.filter((shift, index) =>
            index < shifts.length - 1 &&
            (shifts[index + 1].shiftStart - shift.shiftEnd) / (1000 * 60 * 60) < 10 &&
            (shifts[index + 1].shiftStart - shift.shiftEnd) / (1000 * 60 * 60) > 1
        );

        if (betweenShifts.length > 0) {
            betweenShiftsEmployees.push({ employee, position: employees[employee].position });
        }

        // c) Employees who have worked for more than 14 hours in a single shift
        const longShifts = shifts.filter(shift => shift.shiftDuration > 14);

        if (longShifts.length > 0) {
            longShiftEmployees.push({ employee, position: employees[employee].position });
        }
    }

    // Output the results to the console
    console.log("Employees who have worked for 7 consecutive days:");
    console.log(sevenConsecutiveDaysEmployees);

    console.log("\nEmployees with less than 10 hours between shifts but greater than 1 hour:");
    console.log(betweenShiftsEmployees);

    console.log("\nEmployees who have worked for more than 14 hours in a single shift:");
    console.log(longShiftEmployees);

    // Write the output to a file
    const output = `
Employees who have worked for 7 consecutive days:
${JSON.stringify(sevenConsecutiveDaysEmployees)}

Employees with less than 10 hours between shifts but greater than 1 hour:
${JSON.stringify(betweenShiftsEmployees)}

Employees who have worked for more than 14 hours in a single shift:
${JSON.stringify(longShiftEmployees)}
`;

    fs.writeFileSync('output.txt', output);
}

// Run the analysis
analyzeEmployeeSchedule(filePath);
