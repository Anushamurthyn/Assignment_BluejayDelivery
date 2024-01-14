// Import the exceljs module
const ExcelJS = require('exceljs');

// Create a new workbook object
const workbook = new ExcelJS.Workbook();

// Define a function to convert a time string to a number of minutes
const timeToMinutes = (time) => {
  // Check if the time value is defined and a string
  if (time && typeof time === "string") {
    // Split the time string by colon
    let [hour, minute] = time.split(':');
    // Convert hour and minute to numbers
    hour = Number(hour);
    minute = Number(minute);
    // Return the total number of minutes
    return hour * 60 + minute;
  } else {
    // Return null or throw an error
    return null;
    // or
    // throw new Error("Invalid time value");
  }
};


// Define a function to check if an employee has worked for 7 consecutive days
const hasWorked7Days = (rows) => {
  // Initialize a counter and a flag
  let count = 0;
  let flag = false;
  // Loop through the rows
  for (let i = 0; i < rows.length; i++) {
    // Get the current and previous dates
    let currentDate = rows[i].values[6];
    let previousDate = i > 0 ? rows[i - 1].values[6] : null;
    // Check if the current and previous dates are consecutive
    if (previousDate && isConsecutive(currentDate, previousDate)) {
      // Increment the counter
      count++;
    } else {
      // Reset the counter
      count = 0;
    }
    // Check if the counter reaches 7
    if (count === 7) {
      // Set the flag to true and break the loop
      flag = true;
      break;
    }
  }
  // Return the flag
  return flag;
};

// Define a function to check if an employee has less than 10 hours of time between shifts but greater than 1 hour
const hasLessThan10HoursBetweenShifts = (rows) => {
  // Initialize a flag
  let flag = false;
  // Loop through the rows
  for (let i = 0; i < rows.length - 1; i++) {
    // Get the current and next dates
    let currentDate = rows[i].values[6];
    let nextDate = rows[i + 1].values[6];
    // Check if the dates are defined and the same
    if (currentDate && nextDate && currentDate.getTime() === nextDate.getTime()) {
      // Get the current and next time out and time in
      let currentTimeOut = rows[i].values[4];
      let nextTimeIn = rows[i + 1].values[3];
      // Convert the time strings to minutes
      let currentMinutes = timeToMinutes(currentTimeOut);
      let nextMinutes = timeToMinutes(nextTimeIn);
      // Calculate the difference in minutes
      let diff = nextMinutes - currentMinutes;
      // Check if the difference is less than 10 hours (600 minutes) but greater than 1 hour (60 minutes)
      if (diff < 600 && diff > 60) {
        // Set the flag to true and break the loop
        flag = true;
        break;
      }
    }
  }
  // Return the flag
  return flag;
};


// Define a function to check if an employee has worked for more than 14 hours in a single shift
const hasWorkedMoreThan14Hours = (rows) => {
  // Initialize a flag
  let flag = false;
  // Loop through the rows
  for (let row of rows) {
    // Get the time in and time out
    let timeIn = row.values[3];
    let timeOut = row.values[4];
    // Convert the time strings to minutes
    let inMinutes = timeToMinutes(timeIn);
    let outMinutes = timeToMinutes(timeOut);
    // Calculate the difference in minutes
    let diff = outMinutes - inMinutes;
    // Check if the difference is more than 14 hours (840 minutes)
    if (diff > 840) {
      // Set the flag to true and break the loop
      flag = true;
      break;
    }
  }
  // Return the flag
  return flag;
};

// Define a function to print the name and position of employees who meet the criteria
const printEmployees = (rows) => {
  // Initialize an array to store the employee names and positions
  let employees = [];
  // Loop through the rows
  for (let row of rows) {
    // Get the employee name and position
    let name = row.values[8];
    let position = row.values[1];
    // Check if the name and position are not already in the array
    if (!employees.some((e) => e.name === name && e.position === position)) {
      // Push the name and position to the array
      employees.push({ name, position });
    }
  }
  // Loop through the array and print the name and position
  for (let employee of employees) {
    console.log(`Name: ${employee.name}, Position: ${employee.position}`);
  }
};

// Read the excel file
workbook.xlsx.readFile('Assignment_Timecard.xlsx')
  .then(() => {
    // Get the first worksheet
    const worksheet = workbook.getWorksheet(1);
    // Group the rows by employee name and position
    let groups = worksheet.getSheetValues().reduce((acc, row) => {
      // Get the employee name and position
      let name = row[8];
      let position = row[1];
      // Check if the name and position are valid
      if (name && position) {
        // Create a key from the name and position
        let key = `${name}-${position}`;
        // Check if the key exists in the accumulator object
        if (acc[key]) {
          // Push the row to the existing array
          acc[key].push(row);
        } else {
          // Create a new array with the row
          acc[key] = [row];
        }
      }
      // Return the accumulator object
      return acc;
    }, {});
    // Print the employees who have worked for 7 consecutive days
    console.log('Employees who have worked for 7 consecutive days:');
    for (let key in groups) {
      // Get the rows for the key
      let rows = groups[key];
      // Check if the employee has worked for 7 consecutive days
      if (hasWorked7Days(rows)) {
        // Print the name and position
        printEmployees(rows);
      }
    }
    // Print the employees who have less than 10 hours of time between shifts but greater than 1 hour
    console.log('Employees who have less than 10 hours of time between shifts but greater than 1 hour:');
    for (let key in groups) {
      // Get the rows for the key
      let rows = groups[key];
      // Check if the employee has less than 10 hours of time between shifts but greater than 1 hour
      if (hasLessThan10HoursBetweenShifts(rows)) {
        // Print the name and position
        printEmployees(rows);
      }
    }
    // Print the employees who have worked for more than 14 hours in a single shift
    console.log('Employees who have worked for more than 14 hours in a single shift:');
    for (let key in groups) {
      // Get the rows for the key
      let rows = groups[key];
      // Check if the employee has worked for more than 14 hours in a single shift
      if (hasWorkedMoreThan14Hours(rows)) {
        // Print the name and position
        printEmployees(rows);
      }
    }
  })
  .catch((error) => {
    // Handle any errors
    console.error(error);
  });
