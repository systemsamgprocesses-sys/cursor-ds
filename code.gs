/**
 * Task Management System - Google Apps Script Backend
 * Modified to work with your specific sheet format.
 */

// ⚙️ Configuration
const CONFIG = {
  SHEETS: {
    CREDENTIALS: 'Credentials',
    DROPDOWN: 'Dropdown',
    MASTER: 'MASTER',
    SUBMISSIONS: 'SUBMISSIONS'
  },
  TASK_ID_PREFIX: 'AT-',
  MAX_REVISIONS: 2,
  WORKING_DAYS: [1, 2, 3, 4, 5, 6], // Mon-Sat
};

/**
 * 🚀 Serve the HTML dashboard
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Task Management Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 🔐 Authenticate user against Credentials sheet
 */
function authenticateUser(params) {
  try {
    const { userId, password } = params;
    const sheet = getSheet(CONFIG.SHEETS.CREDENTIALS);
    if (!sheet) throw new Error('Credentials sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const userIdCol = headers.indexOf('User ID');
    const passwordCol = headers.indexOf('Password');
    const departmentCol = headers.indexOf('Department');

    if (userIdCol === -1 || passwordCol === -1 || departmentCol === -1) {
      return { success: false, message: 'Invalid credentials sheet format' };
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdCol] === userId && data[i][passwordCol] === password) {
        const department = data[i][departmentCol];
        const role = determineUserRole(userId, department);
        return {
          success: true,
          user: { userId, department, role }
        };
      }
    }
    return { success: false, message: 'Invalid credentials' };
  } catch (error) {
    console.error('authenticateUser error:', error);
    return { success: false, message: 'Authentication failed' };
  }
}

/**
 * 🎖️ Determine user role
 */
function determineUserRole(userId, department) {
  // Grant Super Admin ONLY to specific, exact names or the 'admin' ID
  if (userId === 'admin' || userId === 'Karan Malhotra' || userId === 'Sunaina Awasthi') {
    return 'Super Admin';
  }
  // Grant Admin to users with 'admin' in their ID or from IT/HR departments
  if (userId.toLowerCase().includes('admin') || ['IT', 'HR'].includes(department)) {
    return 'Admin';
  }
  // Everyone else is a Normal User
  return 'Normal User';
}

/**
 * 📋 Get dropdown data for task assignment
 */
function getDropdownData() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.DROPDOWN);
    if (!sheet) throw new Error('Dropdown sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 👇 FIXED: Match exact or partial header text as per your sheet
    const nameCol = headers.findIndex(h => 
      h.toLowerCase().includes("doer's name") || 
      h.toLowerCase().includes("doer") || 
      h.toLowerCase().includes("name") && h.toLowerCase().includes("doer")
    );
    const deptCol = headers.findIndex(h => 
      h.toLowerCase().includes('department name') || 
      h.toLowerCase().includes('department')
    );
    const userIdCol = headers.findIndex(h => 
      h.toLowerCase().includes('user id') || 
      h.toLowerCase().includes('userid')
    );

    // Optional: Log column indices for debugging
    console.log(`Name Col Index: ${nameCol}, Dept Col Index: ${deptCol}, User ID Col Index: ${userIdCol}`);

    const dropdownData = [];
    for (let i = 1; i < data.length; i++) {
      const name = data[i][nameCol];
      const dept = data[i][deptCol];
      const userId = data[i][userIdCol];

      // Only add if all required fields exist
      if (name && dept && userId) {
        dropdownData.push({
          name: name.trim(),
          department: dept.trim(),
          userId: userId.trim()
        });
      }
    }

    return { success: true, data: dropdownData };
  } catch (error) {
    console.error('getDropdownData error:', error);
    return { success: false, message: 'Failed to load users' };
  }
}
/**
 * 🏢 Get all departments (for Super Admin)
 */
function getDepartments() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.DROPDOWN);
    if (!sheet) throw new Error('Dropdown sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const deptCol = headers.findIndex(h => h.toLowerCase().includes('department'));

    const departments = [...new Set(data.slice(1).map(row => row[deptCol]).filter(d => d))];
    return { success: true, departments: departments.sort() };
  } catch (error) {
    console.error('getDepartments error:', error);
    return { success: false, message: 'Failed to load departments' };
  }
}

/**
 * ➕ Assign task with frequency expansion
 */
function assignTask(params) {
  try {
    const {
      givenBy,
      givenTo,
      givenToName,
      taskDescription,
      tutorialLinks,
      department,
      taskFrequency,
      plannedDate
    } = params;

    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const taskInstances = generateTaskInstances({ taskDescription, taskFrequency, plannedDate });
    const nextTaskId = getNextTaskId(sheet);

    const rows = taskInstances.map((instance, index) => [
      CONFIG.TASK_ID_PREFIX + (nextTaskId + index).toString().padStart(3, '0'), // Task Id
      givenBy,     // GIVEN BY
      givenToName, // GIVEN TO
      givenTo,     // GIVEN TO USER ID
      taskDescription, // TASK DESCRIPTION
      tutorialLinks || '', // HOW TO DO- TUTORIAL LINKS (OPTIONAL)
      department,  // DEPARTMENT
      taskFrequency, // TASK FREQUENCY
      instance.plannedDate, // PLANNED DATE (in DD/MM/YYYY format)
      'Pending',   // Task Status
      '',          // New Date if Any
      '',          // Reason
      '',          // Task Completed On
      '',          // Revision Status
      '',          // Revision 1 Date
      ''           // Revision 2 Date
    ]);

    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    }

    return {
      success: true,
      message: `Task assigned successfully! Created ${rows.length} instances.`,
      tasksCreated: rows.length
    };
  } catch (error) {
    console.error('assignTask error:', error);
    return { success: false, message: 'Failed to assign task: ' + error.message };
  }
}

/**
 * 🔄 Generate task instances based on frequency
 * Returns dates in DD/MM/YYYY format to match your sheet.
 */
function generateTaskInstances(params) {
  const { taskFrequency, plannedDate } = params;
  const startDate = new Date(plannedDate);
  const instances = [];

  switch (taskFrequency) {
    case 'Daily':
      for (let i = 0; i < 365 && instances.length < 300; i++) {
        const date = new Date(startDate);
        date.setDate(startDate.getDate() + i);
        if (CONFIG.WORKING_DAYS.includes(date.getDay())) {
          instances.push({ plannedDate: formatDateForSheetDDMMYYYY(date) });
        }
      }
      break;
    case 'Weekly':
      for (let i = 0; i < 52; i++) {
        const date = new Date(startDate);
        date.setDate(startDate.getDate() + (i * 7));
        instances.push({ plannedDate: formatDateForSheetDDMMYYYY(date) });
      }
      break;
    case 'Monthly':
      for (let i = 0; i < 12; i++) {
        const date = new Date(startDate);
        date.setMonth(startDate.getMonth() + i);
        instances.push({ plannedDate: formatDateForSheetDDMMYYYY(date) });
      }
      break;
    case 'Quarterly':
      for (let i = 0; i < 4; i++) {
        const date = new Date(startDate);
        date.setMonth(startDate.getMonth() + (i * 3));
        instances.push({ plannedDate: formatDateForSheetDDMMYYYY(date) });
      }
      break;
    case 'One Time Only':
    default:
      instances.push({ plannedDate: formatDateForSheetDDMMYYYY(startDate) });
      break;
  }
  return instances;
}

/**
 * 🔢 Get next Task ID
 */
function getNextTaskId(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 1;

  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const taskId = data[i][0]; // Column A
    if (taskId && typeof taskId === 'string' && taskId.startsWith(CONFIG.TASK_ID_PREFIX)) {
      const num = parseInt(taskId.substring(CONFIG.TASK_ID_PREFIX.length));
      if (!isNaN(num) && num > maxId) maxId = num;
    }
  }
  return maxId + 1;
}

/**
 * 📅 Format date for Sheets in DD/MM/YYYY (your format)
 */
function formatDateForSheetDDMMYYYY(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * 📊 Get overview stats
 * UPDATED: Uses column numbers and robust date handling.
 */
function getOverviewStats(params) {
  try {
    const { userRole, userDepartment } = params;
    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, stats: { total: 0, completed: 0, pending: 0, overdue: 0 }, recentTasks: [], distribution: {} };
    }

    const cols = {
      department: 6,   // G
      status: 9,       // J
      plannedDate: 8   // I
    };

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let total = 0, completed = 0, pending = 0, overdue = 0;
    const recentTasks = [];
    const distribution = { 'Pending': 0, 'Completed': 0, 'Overdue': 0, 'Revise': 0 };

    // Helper to parse date safely
    const parseDate = (value) => {
      if (!value) return null;
      let date;
      if (value instanceof Date) {
        date = value;
      } else if (typeof value === 'string') {
        if (value.includes('/')) {
          const parts = value.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            date = new Date(year, month, day);
          }
        } else {
          date = new Date(value);
        }
      }
      return (date && !isNaN(date.getTime())) ? date : null;
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (userRole !== 'Super Admin' && row[cols.department] !== userDepartment) continue;

      const status = row[cols.status] || '';
      const plannedDateStr = row[cols.plannedDate];
      const plannedDate = parseDate(plannedDateStr);

      total++;
      if (distribution[status] !== undefined) distribution[status]++;

      switch (status) {
        case 'Completed': 
          completed++; 
          break;
        case 'Pending':
          pending++;
          if (plannedDate && plannedDate < today) { 
            overdue++; 
            pending--; 
          }
          break;
        case 'Overdue': 
          overdue++; 
          break;
        case 'Revise': 
          pending++; 
          break;
      }

      if (recentTasks.length < 5) {
        recentTasks.push({
          taskId: row[0] || '',
          description: row[4] || '',
          status: status
        });
      }
    }

    return { success: true, stats: { total, completed, pending, overdue }, recentTasks, distribution };
  } catch (error) {
    console.error('getOverviewStats error:', error);
    return { success: false, message: 'Failed to load stats: ' + error.message };
  }
}

/**
 * 📑 Get tasks (filtered by section, role, etc.)
 * UPDATED: Uses column numbers and robust date handling.
 */
function getTasks(params) {
  try {
    const { section, userRole, userDepartment, filters = {} } = params;
    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, tasks: [] };

    // 🔥 FIXED: Use static column numbers (0-based index)
    const cols = {
      taskId: 0,        // A
      givenBy: 1,       // B
      givenTo: 2,       // C
      userId: 3,        // D
      description: 4,   // E
      tutorialLinks: 5, // F
      department: 6,    // G
      frequency: 7,     // H
      plannedDate: 8,   // I - PLANNED DATE
      status: 9,        // J
      newDate: 10,      // K - New Date if Any
      reason: 11,       // L
      completedOn: 12,  // M - Task Completed On
      revisionStatus: 13, // N
      newDate1: 14,     // O - Revision 1 Date
      newDate2: 15      // P - Revision 2 Date
    };

    const tasks = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Ensure user info is available
      if (!userRole || !userDepartment) {
        continue;
      }

      // Filter by department (unless Super Admin)
      if (userRole !== 'Super Admin' && row[cols.department] !== userDepartment) {
        continue;
      }

      let include = true;
      const status = row[cols.status] || '';

      // Section filter
      if (section === 'upcoming' && status !== 'Pending') include = false;
      if (section === 'pending' && !['Pending', 'Revise'].includes(status)) include = false;

      // Additional filters
      if (include && filters.department && row[cols.department] !== filters.department) include = false;
      if (include && filters.status && row[cols.status] !== filters.status) include = false;
      if (include && filters.frequency && row[cols.frequency] !== filters.frequency) include = false;
      if (include && filters.search) {
        const term = filters.search.toLowerCase();
        const searchable = [row[cols.taskId], row[cols.description], row[cols.givenTo], row[cols.userId]]
          .some(field => (field || '').toString().toLowerCase().includes(term));
        if (!searchable) include = false;
      }

      if (include) {
        // 👇 CREATE TASK OBJECT
        const task = {
          taskId: row[cols.taskId] || '',
          givenBy: row[cols.givenBy] || '',
          givenTo: row[cols.givenTo] || '',
          userId: row[cols.userId] || '',
          description: row[cols.description] || '',
          tutorialLinks: row[cols.tutorialLinks] || '',
          department: row[cols.department] || '',
          frequency: row[cols.frequency] || '',
          status: status,
          reason: row[cols.reason] || '',
          revisionStatus: row[cols.revisionStatus] || ''
        };

        // 👇 ROBUST DATE HANDLING FOR ALL DATE FIELDS

        // Helper function to safely format a date
        const formatDate = (value) => {
          if (!value) return ''; // Handle null/undefined/empty

          let date;
          if (value instanceof Date) {
            // If it's already a Date object
            date = value;
          } else if (typeof value === 'string') {
            // Try to parse common formats
            if (value.includes('/')) {
              // Assume DD/MM/YYYY
              const parts = value.split('/');
              if (parts.length === 3) {
                const day = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10) - 1; // JS months are 0-indexed
                const year = parseInt(parts[2], 10);
                date = new Date(year, month, day);
              }
            } else {
              // Try standard JS Date parsing
              date = new Date(value);
            }
          }

          // Validate the date
          if (date && !isNaN(date.getTime())) {
            // Format as YYYY-MM-DD for frontend compatibility
            return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            return ''; // Invalid date
          }
        };

        // Apply date formatting to all date fields
        task.plannedDate = formatDate(row[cols.plannedDate]);
        task.newDate = formatDate(row[cols.newDate]);
        task.completedOn = formatDate(row[cols.completedOn]);
        task.newDate1 = formatDate(row[cols.newDate1]);
        task.newDate2 = formatDate(row[cols.newDate2]);

        tasks.push(task);
      }
    }

    // Sort by planned date (newest first)
    tasks.sort((a, b) => {
      const dateA = a.plannedDate ? new Date(a.plannedDate) : new Date(0);
      const dateB = b.plannedDate ? new Date(b.plannedDate) : new Date(0);
      return dateB - dateA; // Newest first
    });

    return { success: true, tasks };
  } catch (error) {
    console.error('getTasks error:', error);
    return { success: false, message: 'Failed to load tasks: ' + error.message };
  }
}

/**
 * 📝 Get revisions for approval
 * UPDATED: Uses column numbers and robust date handling.
 */
function getRevisions(params) {
  try {
    const { userRole, userDepartment } = params;
    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, revisions: [] };

    // 🔥 FIXED: Use static column numbers (0-based index)
    const cols = {
      taskId: 0,        // A
      description: 4,   // E
      givenTo: 2,       // C
      userId: 3,        // D
      department: 6,    // G
      plannedDate: 8,   // I
      status: 9,        // J
      newDate: 10,      // K
      reason: 11        // L
    };

    const revisions = [];

    // Helper function to safely format a date
    const formatDate = (value) => {
      if (!value) return '';
      let date;
      if (value instanceof Date) {
        date = value;
      } else if (typeof value === 'string') {
        if (value.includes('/')) {
          const parts = value.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            date = new Date(year, month, day);
          }
        } else {
          date = new Date(value);
        }
      }
      if (date && !isNaN(date.getTime())) {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        return '';
      }
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (!userRole || !userDepartment) {
        continue;
      }

      if (row[cols.status] === 'Revise') {
        if (userRole === 'Super Admin' || row[cols.department] === userDepartment) {
          revisions.push({
            taskId: row[cols.taskId] || '',
            description: row[cols.description] || '',
            givenTo: row[cols.givenTo] || '',
            userId: row[cols.userId] || '',
            department: row[cols.department] || '',
            plannedDate: formatDate(row[cols.plannedDate]),
            newDate: formatDate(row[cols.newDate]),
            reason: row[cols.reason] || ''
          });
        }
      }
    }
    return { success: true, revisions };
  } catch (error) {
    console.error('getRevisions error:', error);
    return { success: false, message: 'Failed to load revisions: ' + error.message };
  }
}

/**
 * ℹ️ Get task details by taskId using fixed column indices
 */
function getTaskDetails(params) {
  try {
    const { taskId } = params;
    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'No data found' };

    // 🔥 FIXED: Use STATIC COLUMN INDICES (0-based) — MATCHES YOUR DATA STRUCTURE
    const cols = {
      taskId: 0,        // A - Task Id
      givenBy: 1,       // B - GIVEN BY
      givenTo: 2,       // C - GIVEN TO
      userId: 3,        // D - GIVEN TO USER ID
      description: 4,   // E - TASK DESCRIPTION
      tutorialLinks: 5, // F - HOW TO DO- TUTORIAL LINKS (OPTIONAL)
      department: 6,    // G - DEPARTMENT
      frequency: 7,     // H - TASK FREQUENCY
      plannedDate: 8,   // I - PLANNED DATE
      status: 9,        // J - Task Status
      newDate: 10,      // K - New Date if Any
      reason: 11,       // L - Reason
      completedOn: 12,  // M - Task Completed On
      revisionStatus: 13, // N - Revision Status
      newDate1: 14,     // O - Revision 1 Date
      newDate2: 15      // P - Revision 2 Date
    };

    for (let i = 1; i < data.length; i++) {
      if (data[i][cols.taskId] === taskId) {
        // Format dates safely
        const formatDate = (value) => {
          if (!value) return '';
          let date;
          if (value instanceof Date) date = value;
          else if (typeof value === 'string') {
            if (value.includes('/')) {
              const [day, month, year] = value.split('/');
              date = new Date(year, month - 1, day);
            } else {
              date = new Date(value);
            }
          }
          return date && !isNaN(date.getTime()) 
            ? Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd") 
            : '';
        };

        return {
          success: true,
          task: {
            taskId: data[i][cols.taskId] || '',
            givenBy: data[i][cols.givenBy] || '',
            givenTo: data[i][cols.givenTo] || '',
            userId: data[i][cols.userId] || '',
            description: data[i][cols.description] || '',
            tutorialLinks: data[i][cols.tutorialLinks] || '',
            department: data[i][cols.department] || '',
            frequency: data[i][cols.frequency] || '',
            plannedDate: formatDate(data[i][cols.plannedDate]),
            status: data[i][cols.status] || '',
            newDate: formatDate(data[i][cols.newDate]),
            reason: data[i][cols.reason] || '',
            completedOn: formatDate(data[i][cols.completedOn]),
            revisionStatus: data[i][cols.revisionStatus] || '',
            newDate1: formatDate(data[i][cols.newDate1]),
            newDate2: formatDate(data[i][cols.newDate2])
          }
        };
      }
    }
    return { success: false, message: 'Task not found' };
  } catch (error) {
    console.error('getTaskDetails error:', error);
    return { success: false, message: 'Failed to load task details: ' + error.message };
  }
}
/**
 * ✅ Approve revision
 */
function approveRevision(params) {
  try {
    const { taskId, newDate } = params;
    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const cols = {
      taskId: 0, // Task Id
      status: 9, // Task Status
      newDate: 10, // New Date if Any
      revisionStatus: 13, // Revision Status
      newDate1: 14, // Revision 1 Date
      newDate2: 15 // Revision 2 Date
    };

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][cols.taskId] === taskId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Task not found' };
    if (data[rowIndex-1][cols.status] !== 'Revise') return { success: false, message: 'Task not in Revise status' };

    const newDate1 = data[rowIndex-1][cols.newDate1];
    const newDate2 = data[rowIndex-1][cols.newDate2];
    if (newDate1 && newDate2) return { success: false, message: 'Maximum 2 revisions reached' };

    const revisionCol = !newDate1 ? cols.newDate1 + 1 : cols.newDate2 + 1;

    sheet.getRange(rowIndex, cols.status + 1).setValue('Pending');
    sheet.getRange(rowIndex, cols.newDate + 1).setValue(newDate);
    sheet.getRange(rowIndex, cols.revisionStatus + 1).setValue('Revision Approved');
    sheet.getRange(rowIndex, revisionCol).setValue(newDate);

    return { success: true, message: 'Revision approved' };
  } catch (error) {
    console.error('approveRevision error:', error);
    return { success: false, message: 'Failed to approve revision' };
  }
}

/**
 * 🧩 Helper: Parse DD/MM/YYYY date string to JavaScript Date object
 */
function parseDateDDMMYYYY(dateString) {
  if (!dateString) return null;
  const parts = dateString.split('/');
  if (parts.length !== 3) return null;
  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; // JS months are 0-indexed
  const year = parseInt(parts[2], 10);
  return new Date(year, month, day);
}

/**
 * 🧩 Helper: Get sheet by name
 */
function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * 🧩 Helper: Get MASTER sheet (fallback to SUBMISSIONS)
 */
function getMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEETS.MASTER) || ss.getSheetByName(CONFIG.SHEETS.SUBMISSIONS);
}



/**
 * LINKS MANAGEMENT FUNCTIONS
 */

// Add this to your CONFIG object
CONFIG.SHEETS.LINKS = 'LINKS';

/**
 * Get all links for Super Admin
 */
function getLinks() {
  try {
    const sheet = getLinksSheet();
    if (!sheet) {
      // Create the sheet if it doesn't exist
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.insertSheet(CONFIG.SHEETS.LINKS);
      // Set headers
      sheet.appendRow(['ID', 'Title', 'URL', 'Added By', 'Added On']);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, links: [] };
    }
    
    const headers = data[0];
    const idCol = 0; // A
    const titleCol = 1; // B
    const urlCol = 2; // C
    const addedByCol = 3; // D
    const addedOnCol = 4; // E
    
    const links = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol]) { // Only include rows with ID
        links.push({
          id: data[i][idCol],
          title: data[i][titleCol] || '',
          url: data[i][urlCol] || '',
          addedBy: data[i][addedByCol] || '',
          addedOn: data[i][addedOnCol] || ''
        });
      }
    }
    
    // Sort by added date (newest first)
    links.sort((a, b) => {
      if (a.addedOn && b.addedOn) {
        return new Date(b.addedOn) - new Date(a.addedOn);
      }
      return 0;
    });
    
    return { success: true, links };
  } catch (error) {
    console.error('getLinks error:', error);
    return { success: false, message: 'Failed to load links' };
  }
}

/**
 * Add a new link
 */
function addLink(params) {
  try {
    const { title, url } = params;
    const sheet = getLinksSheet();
    
    if (!sheet) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.insertSheet(CONFIG.SHEETS.LINKS);
      sheet.appendRow(['ID', 'Title', 'URL', 'Added By', 'Added On']);
    }
    
    // Generate unique ID
    const id = 'LINK-' + Date.now();
    
    // Get current user (you may need to pass this from frontend)
    const addedBy = Session.getActiveUser().getEmail() || 'Unknown';
    const addedOn = new Date();
    
    // Add the link
    sheet.appendRow([
      id,
      title,
      url,
      addedBy,
      addedOn
    ]);
    
    return { success: true, message: 'Link added successfully' };
  } catch (error) {
    console.error('addLink error:', error);
    return { success: false, message: 'Failed to add link: ' + error.message };
  }
}

/**
 * Delete a link
 */
function deleteLink(params) {
  try {
    const { id } = params;
    const sheet = getLinksSheet();
    
    if (!sheet) {
      return { success: false, message: 'Links sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: 'Link not found' };
    }
    
    // Find the row with the matching ID
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        rowIndex = i + 1; // Convert to 1-based index
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Link not found' };
    }
    
    // Delete the row
    sheet.deleteRow(rowIndex);
    
    return { success: true, message: 'Link deleted successfully' };
  } catch (error) {
    console.error('deleteLink error:', error);
    return { success: false, message: 'Failed to delete link: ' + error.message };
  }
}

/**
 * Helper: Get LINKS sheet
 */
function getLinksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEETS.LINKS);
}

/**
 * Get person-wise tasks for reports
 */
function getPersonTasks(params) {
  try {
    const { personId, userRole, userDepartment } = params;
    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, tasks: [], stats: { total: 0, completed: 0, pending: 0, overdue: 0 } };

    // Use static column numbers (0-based index)
    const cols = {
      taskId: 0,        // A
      givenBy: 1,       // B
      givenTo: 2,       // C
      userId: 3,        // D
      description: 4,   // E
      tutorialLinks: 5, // F
      department: 6,    // G
      frequency: 7,     // H
      plannedDate: 8,   // I
      status: 9,        // J
      newDate: 10,      // K
      reason: 11,       // L
      completedOn: 12,  // M
      revisionStatus: 13, // N
      newDate1: 14,     // O
      newDate2: 15      // P
    };

    const tasks = [];
    let total = 0, completed = 0, pending = 0, overdue = 0;
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Helper function to safely format a date
    const formatDate = (value) => {
      if (!value) return '';
      let date;
      if (value instanceof Date) {
        date = value;
      } else if (typeof value === 'string') {
        if (value.includes('/')) {
          const parts = value.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            date = new Date(year, month, day);
          }
        } else {
          date = new Date(value);
        }
      }
      if (date && !isNaN(date.getTime())) {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        return '';
      }
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip if not the selected person
      if (row[cols.userId] !== personId) continue;
      
      // Skip if not Super Admin and not in user's department
      if (userRole !== 'Super Admin' && row[cols.department] !== userDepartment) continue;
      
      const status = row[cols.status] || '';
      const plannedDate = formatDate(row[cols.plannedDate]);
      
      // Count stats
      total++;
      if (status === 'Completed') completed++;
      else if (status === 'Pending') {
        // Check if overdue
        const plannedDateObj = new Date(plannedDate);
        if (plannedDateObj < today) {
          overdue++;
        } else {
          pending++;
        }
      } else if (status === 'Overdue') {
        overdue++;
      }

      // Create task object
      const task = {
        taskId: row[cols.taskId] || '',
        givenBy: row[cols.givenBy] || '',
        givenTo: row[cols.givenTo] || '',
        userId: row[cols.userId] || '',
        description: row[cols.description] || '',
        tutorialLinks: row[cols.tutorialLinks] || '',
        department: row[cols.department] || '',
        frequency: row[cols.frequency] || '',
        status: status,
        plannedDate: plannedDate,
        completedOn: formatDate(row[cols.completedOn]),
        assignedDate: formatDate(row[cols.plannedDate]), // Using planned date as assigned date
        reason: row[cols.reason] || '',
        revisionStatus: row[cols.revisionStatus] || '',
        newDate1: formatDate(row[cols.newDate1]),
        newDate2: formatDate(row[cols.newDate2])
      };

      tasks.push(task);
    }

    // Sort by planned date (newest first)
    tasks.sort((a, b) => {
      const dateA = a.plannedDate ? new Date(a.plannedDate) : new Date(0);
      const dateB = b.plannedDate ? new Date(b.plannedDate) : new Date(0);
      return dateB - dateA;
    });

    return { 
      success: true, 
      tasks,
      stats: { total, completed, pending, overdue }
    };
  } catch (error) {
    console.error('getPersonTasks error:', error);
    return { success: false, message: 'Failed to load person tasks: ' + error.message };
  }
}

/**
 * Complete a task
 */
function completeTask(params) {
  try {
    const { taskId, userId } = params;
    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const cols = {
      taskId: 0,        // A
      userId: 3,        // D
      status: 9,        // J
      completedOn: 12   // M
    };

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][cols.taskId] === taskId) {
        // Verify user has permission to complete this task
        if (data[i][cols.userId] !== userId) {
          return { success: false, message: 'You do not have permission to complete this task.' };
        }
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Task not found' };

    // Update status and completed date
    sheet.getRange(rowIndex, cols.status + 1).setValue('Completed');
    sheet.getRange(rowIndex, cols.completedOn + 1).setValue(new Date());

    return { success: true, message: 'Task marked as completed' };
  } catch (error) {
    console.error('completeTask error:', error);
    return { success: false, message: 'Failed to complete task: ' + error.message };
  }
}

/**
 * Request revision for a task
 */
function requestRevision(params) {
  try {
    const { taskId, newDate, reason, userId } = params;
    const sheet = getMasterSheet();
    if (!sheet) throw new Error('Master sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const cols = {
      taskId: 0,        // A
      userId: 3,        // D
      status: 9,        // J
      newDate: 10,      // K
      reason: 11        // L
    };

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][cols.taskId] === taskId) {
        // Verify user has permission to request revision for this task
        if (data[i][cols.userId] !== userId) {
          return { success: false, message: 'You do not have permission to request revision for this task.' };
        }
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Task not found' };

    // Update status, new date, and reason
    sheet.getRange(rowIndex, cols.status + 1).setValue('Revise');
    sheet.getRange(rowIndex, cols.newDate + 1).setValue(newDate);
    sheet.getRange(rowIndex, cols.reason + 1).setValue(reason);

    return { success: true, message: 'Revision requested successfully' };
  } catch (error) {
    console.error('requestRevision error:', error);
    return { success: false, message: 'Failed to request revision: ' + error.message };
  }
}
