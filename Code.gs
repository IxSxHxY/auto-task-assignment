function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Task Assignment System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tasksSheet = ss.getSheetByName('Tasks');
  let employeesSheet = ss.getSheetByName('Employees');

  if (!tasksSheet) {
    tasksSheet = ss.insertSheet('Tasks');
    tasksSheet.appendRow(['Task ID', 'Task Name', 'Department', 'People Needed', 'Status', 'Assigned Employees', 'Details']); //, 'Chat URL']);
  }

  if (!employeesSheet) {
    employeesSheet = ss.insertSheet('Employees');
    employeesSheet.appendRow(['Employee ID', 'Name', 'Email', 'Department', 'Availability', 'Assigned Task ID']);
  }
}

function addTaskToSheet(name, department, peopleNeeded, details) {
  initializeSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const taskId = Utilities.getUuid();
  sheet.appendRow([taskId, name, department, peopleNeeded, 'Unassigned', '', details, '']);
  return getTasks();
}

function addEmployeeToSheet(name, email, department) {
  initializeSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');

  // Check for duplicate email
  const data = sheet.getDataRange().getValues();
  const isDuplicate = data.some(row => row[2] === email);
  if (isDuplicate) {
    throw new Error("An employee with this email already exists.");
  }

  const employeeId = Utilities.getUuid();
  sheet.appendRow([employeeId, name, email, department, 'available', '']);
  return getEmployees();
}

function editTask(taskId, name, department, details, status) {
  const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const taskData = taskSheet.getDataRange().getValues();

  for (let i = 1; i < taskData.length; i++) {
    if (taskData[i][0] === taskId) {
      taskSheet.getRange(i + 1, 2).setValue(name);
      taskSheet.getRange(i + 1, 3).setValue(department);
      taskSheet.getRange(i + 1, 7).setValue(details);
      taskSheet.getRange(i + 1, 5).setValue(status);

      if (status === 'Completed') {
        const assignedEmployees = taskData[i][5].split(', ');
        const employeeData = employeeSheet.getDataRange().getValues();
        for (let j = 1; j < employeeData.length; j++) {
          if (assignedEmployees.includes(employeeData[j][1])) {
            employeeSheet.getRange(j + 1, 5).setValue('available');
            employeeSheet.getRange(j + 1, 6).setValue('');
          }
        }
        taskSheet.getRange(i + 1, 6).setValue('');
      }
      break;
    }
  }

  return getTasks();
}

function completeTask(taskId) {
  const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const taskData = taskSheet.getDataRange().getValues();

  for (let i = 1; i < taskData.length; i++) {
    if (taskData[i][0] === taskId) {
      taskSheet.getRange(i + 1, 5).setValue('Completed');
      const assignedEmployees = taskData[i][5].split(', ');
      const employeeData = employeeSheet.getDataRange().getValues();
      for (let j = 1; j < employeeData.length; j++) {
        if (assignedEmployees.includes(employeeData[j][1])) {
          employeeSheet.getRange(j + 1, 5).setValue('available');
          employeeSheet.getRange(j + 1, 6).setValue('');
        }
      }
      taskSheet.getRange(i + 1, 6).setValue('');
      break;
    }
  }

  return {
    tasks: getTasks(),
    employees: getEmployees()
    };
}

function editEmployee(employeeId, name, department, availability) {
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const employeeData = employeeSheet.getDataRange().getValues();

  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][0] === employeeId) {
      employeeSheet.getRange(i + 1, 2).setValue(name);
      employeeSheet.getRange(i + 1, 4).setValue(department);
      employeeSheet.getRange(i + 1, 5).setValue(availability);
      break;
    }
  }

  return getEmployees();
}


function deleteTask(taskId) {
  const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const taskData = taskSheet.getDataRange().getValues();

  for (let i = 1; i < taskData.length; i++) {
    if (taskData[i][0] === taskId) {
      if (taskData[i][4] === 'In Progress') {
        // Make assigned employees available
        const assignedEmployees = taskData[i][5].split(', ');
        const employeeData = employeeSheet.getDataRange().getValues();
        for (let j = 1; j < employeeData.length; j++) {
          if (assignedEmployees.includes(employeeData[j][1])) {
            employeeSheet.getRange(j + 1, 5).setValue('available');
            employeeSheet.getRange(j + 1, 6).setValue('');
          }
        }
      }
      taskSheet.deleteRow(i + 1);
      break;
    }
  }

  return {
    tasks: getTasks(),
    employees: getEmployees()
    };
}

function deleteEmployee(employeeId) {
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const employeeData = employeeSheet.getDataRange().getValues();

  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][0] === employeeId) {
      if (employeeData[i][5] !== '') {
        throw new Error("Cannot delete employee. They are currently assigned to a task.");
      }
      employeeSheet.deleteRow(i + 1);
      break;
    }
  }

  return getEmployees();
}

function autoAssignTasks() {
  initializeSheets();
  const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');

  const tasks = taskSheet.getDataRange().getValues();
  const employees = employeeSheet.getDataRange().getValues();

  for (let i = 1; i < tasks.length; i++) {
    if (tasks[i][4] === 'Unassigned') {
      const taskId = tasks[i][0];
      const department = tasks[i][2];
      const peopleNeeded = parseInt(tasks[i][3]);
      let assigned = [];
      let assignedEmails = [];

      for (let j = 1; j < employees.length && assigned.length < peopleNeeded; j++) {
        if (employees[j][3] === department && employees[j][4] === 'available') {
          assigned.push(employees[j][1]);
          assignedEmails.push(employees[j][2]);
          employeeSheet.getRange(j + 1, 5).setValue('unavailable');
          employeeSheet.getRange(j + 1, 6).setValue(taskId);
          employees[j][4] = 'unavailable';
          employees[j][5] = taskId;
        }
      }

      if (assigned.length === peopleNeeded) {
        taskSheet.getRange(i + 1, 5).setValue('In Progress');
        taskSheet.getRange(i + 1, 6).setValue(assigned.join(', '));

        // Create Google Chat group if more than one person is assigned
        // if (assigned.length > 1) {
        //   try {
        //     const chatUrl = createChatGroup(tasks[i][1], assignedEmails);
        //     taskSheet.getRange(i + 1, 8).setValue(chatUrl);
        //   } catch (error) {
        //     console.error('Error creating chat group:', error);
        //     taskSheet.getRange(i + 1, 8).setValue('Failed to create chat group');
        //   }
        // }

        // Send email notifications
        sendAssignmentEmails(tasks[i][1], assigned, assignedEmails);
      }
    }
  }

  return {
    tasks: getTasks(),
    employees: getEmployees()
    };
}

function getTasks(page = 1, pageSize = 10) {
  initializeSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  const totalTasks = data.length - 1;  // Subtract 1 for header row
  const totalPages = Math.ceil(totalTasks / pageSize);

  const startIndex = (page - 1) * pageSize + 1;  // +1 to skip header
  const endIndex = Math.min(startIndex + pageSize, data.length);

  const tasks = data.slice(startIndex, endIndex).map(row => ({
    id: row[0],
    name: row[1],
    department: row[2],
    peopleNeeded: row[3],
    status: row[4],
    assignedEmployees: row[5] ? row[5].split(', ') : [],
    details: row[6],
    // chatUrl: row[7]
  }));
  console.log(123123);
  return {
    tasks: tasks,
    currentPage: page,
    totalPages: totalPages,
    totalTasks: totalTasks
  };
}

function getEmployees(page = 1, pageSize = 10) {
  initializeSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const data = sheet.getDataRange().getValues();
  const totalEmployees = data.length - 1;  // Subtract 1 for header row
  const totalPages = Math.ceil(totalEmployees / pageSize);

  const startIndex = (page - 1) * pageSize + 1;  // +1 to skip header
  const endIndex = Math.min(startIndex + pageSize, data.length);

  const employees = data.slice(startIndex, endIndex).map(row => ({
    id: row[0],
    name: row[1],
    email: row[2],
    department: row[3],
    availability: row[4],
    assignedTaskId: row[5]
  }));
  console.log(employees);
  return {
    employees: employees,
    currentPage: page,
    totalPages: totalPages,
    totalEmployees: totalEmployees
  };
}

// function createChatGroup(taskName, emails) {
//   // This function would use the Google Chat API to create a new chat space
//   // and invite the assigned employees. For now, we'll return a placeholder URL.
//   return `https://chat.google.com/room/${Utilities.getUuid()}`;
// }

function sendAssignmentEmails(taskName, assignedNames, assignedEmails) {
  const subject = `New Task Assignment: ${taskName}`;
  const body = `You have been assigned to the task "${taskName}" along with: ${assignedNames.join(', ')}`;

  assignedEmails.forEach(email => {
    MailApp.sendEmail(email, subject, body);
  });
}

function onOpen() {
  initializeSheets();
}