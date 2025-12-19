// ==================================================================================
// GOOGLE SHEETS ACCESS MANAGER 12-05-2025 
// created by: Keith Andersen with assistance of Claude Ai
// Main Apps Script File
// ==================================================================================
// This script manages access to Google Spreadsheets by allowing users to check out
// and check in spreadsheets for editing. Admins can manage users and spreadsheets.
// All data is stored in Script Properties Service (500KB limit).
// ==================================================================================

// ==================================================================================
// CONFIGURATION SETTINGS - Customize these for your deployment
// ==================================================================================

//***  FILE IS VIEW ONLY, MAKE COPY OF THIS APP SCRIPT FILE TO GAIN EDITOR ACCESS  ***
// LOOK FOR ⓘ ON LEFT ( LEFT OF Files ) CLICK ⓘ TO OPEN OVERVIEW. MAKE COPY ICON ON LEFT

//--- OPEN COPY AND PROCEED------

//STEP #1 SET SUPER_ADMIN_EMAIL (EMAIL OF THIS GMAIL ACCOUNT )
/**
 * SUPER ADMIN EMAIL
 * This account CANNOT be modified or deleted by any other admin.
 * This should be the script owner/creator's email.
 */
const SUPER_ADMIN_EMAIL = 'YOUR-EMAIL-ADDRESS HERE';//<----CHANGE THIS TO YOUR EMAIL
//*IMPORTANT - SELECT createSuperAdmin FROM FUNCTIONS MENU ABOVE TO CREATE YOUR SUPER_ADMIN ACCOUNT
//WHEN DEPLOY WEB APP - LOGIN AND CHANGE PASSWORD FROM default1234 TO A SECURE PASSWORD


//STEP #2 - SET CONTACT EMAIL ( CAN BE SUPER_ADMIN_EMAIL OR ANOTHER ADMIN'S)
/**
 * ADMIN CONTACT EMAIL
 * This is the email address where user messages will be sent
 * Change this to your preferred contact email address
 * Then select testEmailPermissions from function menu above and run. Give permissions
 */
const ADMIN_CONTACT_EMAIL = 'ADMIN-EMAIL-ADDRESS-HERE';//<--- CHANGE TO AN ADMIN EMAIL


//STEP #3 - SELECT "testEmailPermissions" from the above menu and RUN. This will prompt the permissions pop-ups
// to give permission for the Contact Admin email feature in the Admin Panel to work. Give all permissions prompted.


//STEP #4 - SET SYSTEMM NAME
/**
 * SYSTEM NAME
 * Customize the name that appears in email notifications
 */
const SYSTEM_NAME = 'Google Sheets Access Manager';


//STEP #5 (OPTIONAL) INCREASE MAX LOG ENTRIES STORED IN PROPERTY SERVICES
//*** BE AWARE MORE LOGS CONSUME KB's AND DECREASE USER / SPREADSHEET SPACE
/**
 * MAXIMUM ACTIVITY LOG ENTRIES
 * Number of activity log entries to retain (affects storage)
 * NO NEED TO CHANGE
 */
const MAX_ACTIVITY_LOG_ENTRIES = 300;


//STEP #6 - SET SESSION TIMEOUT
/**
 * SESSION TIMEOUT (in hours)
 * How long before an inactive session expires (future enhancement)
 */
const SESSION_TIMEOUT_HOURS = 24;
//EITHER MANUALLY CREATE A MIDNIGHT TRIGGER OR SELECT createMidnightTrigger FROM FUNCTION MENU ABOVE
//AND RUN ONLY ONCE


//STEP #7 - DEPLOY WEB APP ( instructions below )
// ==================================================================================
// DEPLOYMENT INSTRUCTIONS
// ==================================================================================
// To deploy this as a web app:
// 
// 1. In Apps Script Editor, click "Deploy" > "New deployment"
// 2. Click the gear icon, select "Web app"
// 3. Fill in the settings:
//    - Description: "Google Sheets Access Manager v1.0"
//    - Execute as: "Me (your email)"
//    - Who has access: "Anyone" (or "Anyone with Google account" for more security - RECOMMENDED! )
// 4. Click "Deploy"
// 5. Copy the Web app URL 
// 6. Grant any required permissions when prompted
//
// After any updates to Code.gs or index.html:
// - Click "Deploy" > "Manage deployments"
// - Click the pencil icon to edit
// - Change version to "New version"
// - Click "Deploy"
// - Copy the Web app URL and put in browser
// 
// IMPORTANT: Every time you make changes to the code, you must create a new 
// deployment version for the changes to appear in the web app.
// ==================================================================================

// STEP #8 IMMPORTANT - After deployment - LOGIN WITH YOUR EMAIL AND CHANGE DEFAULT PASSWORD!!

//-----------------------END CUSTOMIZATIOS-------------------------------------------

/**
 * FUNCTIONS BELOW 
 * doGet ---------------------WEB ENTRY POINT + FILES
 * getDocumentationHtml
 * EVENT_TYPES----------------EVENT_TYPES
 * normalizeEmail
 * sanitizeInput--------------SECURITY AND VALIDATION
 * withLock
 * hashPassword
 * verifyPassword
 * generateSessionToken
 * getAllUsers----------------PROPERTY SERVICE MANAGEMENT
 * saveAllUsers
 * getAllSpreadsheets
 * saveAllSpreadsheets
 * viewCurrentUsers
 * getAdminDashboardData
 * calculateStorageUsage
 * clearAllData--(*not currently integrated into web app controls)
 * loginUser------------------AUTHENTICATION FUNCTIONS
 * logoutUser
 * resetUserPassword 
 * verifySessionToken
 * logActivity----------------ACTIVITY LOGGING SYSTEM
 * getActivityLog
 * clearOldActivityLogs
 * midnightCleanup------------MIDNIGHT CLEAN UP FUNCTIONS
 * emailAdminMidnightErrors
 * registerNewUser------------USERS REGISTRATION FUNCTIONS
 * getSpreadsheetDetails------SPREADSHEET MANAGEMENT FUNCTIONS
 * updateSpreadsheet
 * addSpreadsheet
 * deleteSpreadsheet
 * getSpreadsheetByID
 * getUserAvailableSpreadsheets
 * getAllCheckedOutSpreadsheets
 * releaseUserSpreadsheets
 * getUserCheckedOutCount
 * requestSpreadsheetAccess---USER SECTION FUNCTIONS
 * sendContactMessage
 * isSessionSuperAdmin--------ADMIN SECTION FUNCTIONS
 * isSuperAdmin
 * getUserDetails
 * grantPermissionsForAssignments
 * updateUser
 * deleteUser
 * getLoginLockoutStatus------USER AUTHENTICATION LOGIN LOCK
 * setLoginLockoutStatus
 * isLoginLocked
 * clearAllActiveSessions
 * revokeActivePermissions
 * getUsersBackup
 * getSpreadsheetsBackup
 * buildUpdateMessage
 * emailAdminAboutErrors
 * logDetailedUserUpdate
 * adminResetUserPassword
 */

// ==================================================================================
// WEB APP ENTRY POINT + FILES
// ==================================================================================
/**
 * This function is automatically called when someone visits the web app URL.
 * It serves the index.html file to the user's browser.
 */
function doGet() {
  // Create HTML output from the index.html file
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Google Sheets Access Manager')  // Browser tab title
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);  // Allow embedding if needed
  
  return htmlOutput;
}

/**
 * Return the Documentation.html content as a string
 */
function getDocumentationHtml() {
  try {
    return HtmlService.createHtmlOutputFromFile('Documentation').getContent();
  } catch (error) {
    Logger.log('Error loading documentation: ' + error.toString());
    throw new Error('Failed to load documentation');
  }
}

//===================================================================
//EVENT TYPES
//===================================================================
/**
 * EVENT TYPE CONSTANTS
 * Standardized event type names for consistency
 */
const EVENT_TYPES = {
  // User events
  USER_REGISTERED: 'USER_REGISTERED',
  USER_APPROVED: 'USER_APPROVED',
  USER_REJECTED: 'USER_REJECTED',
  USER_DELETED: 'USER_DELETED',
  USER_MODIFIED: 'USER_MODIFIED',
  LOGIN_SUCCESS: 'LOGIN_SUCCESS',
  LOGIN_FAILED: 'LOGIN_FAILED',
  LOGOUT: 'LOGOUT',
  
  // Spreadsheet events
  SPREADSHEET_ADDED: 'SPREADSHEET_ADDED',
  SPREADSHEET_DELETED: 'SPREADSHEET_DELETED',
  SPREADSHEET_MODIFIED: 'SPREADSHEET_MODIFIED',
  SPREADSHEET_CHECKOUT: 'SPREADSHEET_CHECKOUT',
  SPREADSHEET_CHECKIN: 'SPREADSHEET_CHECKIN',
  SPREADSHEET_ACCESS_DENIED: 'SPREADSHEET_ACCESS_DENIED',
  
  // System events
  SYSTEM_CLEANUP: 'SYSTEM_CLEANUP',
  SESSION_EXPIRED: 'SESSION_EXPIRED',
  PERMISSION_GRANTED: 'PERMISSION_GRANTED',
  PERMISSION_REMOVED: 'PERMISSION_REMOVED',
  CONTACT_ADMIN: 'CONTACT_ADMIN',
  LOGIN_LOCKOUT_ENABLED: 'LOGIN_LOCKOUT_ENABLED',
  LOGIN_LOCKOUT_DISABLED: 'LOGIN_LOCKOUT_DISABLED'
};

/**
 * NORMALIZE EMAIL ADDRESS
 * Converts email to lowercase and trims whitespace for consistent storage/lookup
 */
function normalizeEmail(email) {
  if (!email || typeof email !== 'string') return '';
  return email.toLowerCase().trim();
}


// ==================================================================================
// SECURITY & VALIDATION FUNCTIONS
// ==================================================================================

/**
 * USER FORM INPUT SECURITY VALIDATION
 * REMOVES DANGEROUS CHARACTERS FROM USER INPUT AND ENFORCES A CHARACTER LIMIT
 * RETURNING THE CLEANED USER INPUT
 */
function sanitizeInput(input, maxLength) {
  // Handle undefined parameter
  maxLength = maxLength || null;
  
  // Enhanced danger character set
  const dangerousChars = /[=|<(\];`$]/g; // the backslash \ will be an empty space
  let cleaned = (!input || typeof input !== 'string') ? input : input.replace(dangerousChars, '*');
  
  // Apply character limit if specified
  if (maxLength && typeof cleaned === 'string' && cleaned.length > maxLength) {
    cleaned = cleaned.substring(0, maxLength);
  }
  
  return cleaned;
}

/**
 * LOCK WRAPPER FOR CONCURRENCY PROTECTION
 * Prevents simultaneous operations that could corrupt data
 */
function withLock(lockName, callback) {
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 10 seconds for the lock
    lock.waitLock(10000);
    return callback();
  } catch (e) {
    Logger.log('Lock error for ' + lockName + ': ' + e.toString());
    throw new Error('System is busy, please try again in a moment');
  } finally {
    lock.releaseLock();
  }
}

/**
 * PASSWORD HASHING
 * One-way hash using SHA-256 for secure password storage
 */
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password
  );
  // Convert byte array to hex string
  return rawHash.map(function(byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}

/**
 * VERIFY PASSWORD
 * Compare input password with stored hash
 */
function verifyPassword(inputPassword, storedHash) {
  const inputHash = hashPassword(inputPassword);
  return inputHash === storedHash;
}

/**
 * GENERATE SESSION TOKEN
 * Creates a unique token for authenticated sessions
 */
function generateSessionToken() {
  return Utilities.getUuid();
}

// ==================================================================================
// PROPERTY SERVICE MANAGEMENT
// ==================================================================================

/**
 * GET ALL USERS
 * Retrieves all users from Property Service
 */
function getAllUsers() {
  const props = PropertiesService.getScriptProperties();
  const usersJson = props.getProperty('USERS');
  return usersJson ? JSON.parse(usersJson) : {};
}

/**
 * SAVE ALL USERS
 * Saves users object to Property Service
 */
function saveAllUsers(users) {
  return withLock('users', function() {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('USERS', JSON.stringify(users));
    return true;
  });
}

/**
 * GET ALL SPREADSHEETS
 * Retrieves all spreadsheets from Property Service
 */
function getAllSpreadsheets() {
  const props = PropertiesService.getScriptProperties();
  const sheetsJson = props.getProperty('SPREADSHEETS');
  return sheetsJson ? JSON.parse(sheetsJson) : {};
}

/**
 * SAVE ALL SPREADSHEETS
 * Saves spreadsheets object to Property Service
 */
function saveAllSpreadsheets(spreadsheets) {
  return withLock('spreadsheets', function() {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SPREADSHEETS', JSON.stringify(spreadsheets));
    return true;
  });
}

/**
 * VIEW CURRENT USERS
 * Helper function to see what users are in the system
 */
function viewCurrentUsers() {
  const users = getAllUsers();
  Logger.log('Current users in system:');
  Logger.log(JSON.stringify(users, null, 2));
  return users;
}

/**
 * CALCULATE STORAGE USAGE
 * Calculates current Property Service storage usage and estimates capacity
 */
function calculateStorageUsage() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  
  // Convert all properties to JSON string to calculate size
  const jsonString = JSON.stringify(allProps);
  
  // Calculate bytes using Utilities.newBlob (Apps Script compatible)
  const bytesUsed = Utilities.newBlob(jsonString).getBytes().length;
  const maxBytes = 524288; // 512 KB = 524,288 bytes
  const bytesRemaining = maxBytes - bytesUsed;
  const percentUsed = (bytesUsed / maxBytes * 100).toFixed(2);
  
  // Get current counts
  const users = getAllUsers();
  const spreadsheets = getAllSpreadsheets();
  
  const userCount = Object.keys(users).length;
  const spreadsheetCount = Object.keys(spreadsheets).length;
  
  // Calculate sizes of each component
  const usersJson = props.getProperty('USERS') || '{}';
  const spreadsheetsJson = props.getProperty('SPREADSHEETS') || '{}';
  const activityLogJson = props.getProperty('ACTIVITY_LOG') || '[]';
  const sessionsJson = props.getProperty('ACTIVE_SESSIONS') || '{}';
  
  const usersBytes = Utilities.newBlob(usersJson).getBytes().length;
  const spreadsheetsBytes = Utilities.newBlob(spreadsheetsJson).getBytes().length;
  const activityLogBytes = Utilities.newBlob(activityLogJson).getBytes().length;
  const sessionsBytes = Utilities.newBlob(sessionsJson).getBytes().length;
  
  // Estimate average sizes (with fallback if no items exist)
  const avgBytesPerUser = userCount > 0 ? Math.ceil(usersBytes / userCount) : 200;
  const avgBytesPerSpreadsheet = spreadsheetCount > 0 ? Math.ceil(spreadsheetsBytes / spreadsheetCount) : 300;
  
  // Calculate overhead (activity log + sessions + property structure overhead)
  // Activity log maintains max 200 entries, estimate ~150 bytes per entry
  const estimatedActivityLogOverhead = 200 * 150; // ~30KB reserved for activity log
  const estimatedSessionsOverhead = 5 * 200; // Assume max 5 concurrent sessions at ~200 bytes each
  const systemOverhead = estimatedActivityLogOverhead + estimatedSessionsOverhead; // ~31KB
  
  // Available space for users and spreadsheets (subtract current usage and system overhead)
  const availableForContent = bytesRemaining - systemOverhead;
  
  // Calculate estimated capacity remaining (account for overhead)
  const estimatedUsersRemaining = availableForContent > 0 ? Math.floor(availableForContent / avgBytesPerUser) : 0;
  const estimatedSpreadsheetsRemaining = availableForContent > 0 ? Math.floor(availableForContent / avgBytesPerSpreadsheet) : 0;
  
  // Format bytes for display
  function formatBytes(bytes) {
    if (bytes < 1024) return bytes.toFixed(0) + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
  }
  
  return {
    bytesUsed: bytesUsed,
    bytesUsedFormatted: formatBytes(bytesUsed),
    maxBytes: maxBytes,
    maxBytesFormatted: formatBytes(maxBytes),
    bytesRemaining: bytesRemaining,
    bytesRemainingFormatted: formatBytes(bytesRemaining),
    percentUsed: parseFloat(percentUsed),
    currentUsers: userCount,
    currentSpreadsheets: spreadsheetCount,
    avgBytesPerUser: avgBytesPerUser,
    avgBytesPerSpreadsheet: avgBytesPerSpreadsheet,
    estimatedUsersRemaining: estimatedUsersRemaining,
    estimatedSpreadsheetsRemaining: estimatedSpreadsheetsRemaining,
    // Additional breakdown for admin visibility
    usersBytes: usersBytes,
    spreadsheetsBytes: spreadsheetsBytes,
    activityLogBytes: activityLogBytes,
    sessionsBytes: sessionsBytes,
    systemOverhead: systemOverhead
  };
}

/**
 * GET ADMIN DASHBOARD DATA
 * Retrieves all data needed for admin dashboard display
 */
function getAdminDashboardData(sessionToken) {
  // Verify session token
  const session = verifySessionToken(sessionToken);
  
  if (!session.valid) {
    return { success: false, message: 'Invalid session' };
  }
  
  if (!session.isAdmin) {
    return { success: false, message: 'Admin access required' };
  }
  
  try {
    const props = PropertiesService.getScriptProperties();  // ← ADD THIS LINE
    const storageData = calculateStorageUsage();
    const users = getAllUsers();
    const spreadsheets = getAllSpreadsheets();
    
    // Format users for display
    const usersList = Object.keys(users).map(function(key) {
      return {
        email: users[key].userID,
        status: users[key].registrationApproval,
        userType: users[key].userType,
        registeredDate: users[key].registeredDate,
        isProtected: isSuperAdmin(users[key].userID)
      };
    });    

    // Format spreadsheets for display
    const spreadsheetsList = Object.keys(spreadsheets).map(function(key) {
      return {
        id: spreadsheets[key].sheetID,
        name: spreadsheets[key].sheetTitle,
        status: spreadsheets[key].admin_hold === 'Yes' ? 'inactive' : 'active',
        permissionLevel: spreadsheets[key].access,
        currentUser: spreadsheets[key].currentUser || null,
        allowedUsers: [],
        group: spreadsheets[key].group
      };
    });

    // Get activity log count for display
    const activityLogs = JSON.parse(props.getProperty('ACTIVITY_LOG') || '[]');
    const activityLogCount = activityLogs.length;
    const maxActivityLogs = MAX_ACTIVITY_LOG_ENTRIES;
    
    // Get lockout status
    const lockoutJson = props.getProperty('LOGIN_LOCKOUT');
    let lockoutStatus = { enabled: false, enabledBy: null, enabledAt: null };
    if (lockoutJson) {
      const lockout = JSON.parse(lockoutJson);
      lockoutStatus = {
        enabled: lockout.enabled || false,
        enabledBy: lockout.enabledBy || null,
        enabledAt: lockout.enabledAt || null
      };
    }
    
    return {
      success: true,
      storage: {
        bytesUsed: storageData.bytesUsed,
        bytesUsedFormatted: storageData.bytesUsedFormatted,
        maxBytes: storageData.maxBytes,
        maxBytesFormatted: storageData.maxBytesFormatted,
        bytesRemaining: storageData.bytesRemaining,
        bytesRemainingFormatted: storageData.bytesRemainingFormatted,
        percentUsed: storageData.percentUsed,
        currentUsers: storageData.currentUsers,
        currentSpreadsheets: storageData.currentSpreadsheets,
        activityLogCount: activityLogCount,
        maxActivityLogs: maxActivityLogs
      },
      users: usersList,
      spreadsheets: spreadsheetsList,
      loginLockout: lockoutStatus
    };
  } catch (e) {
    Logger.log('Error getting admin dashboard data: ' + e.toString());
    return { success: false, message: 'Error loading dashboard data: ' + e.message };
  }
}

/**
 * CLEAR ALL DATA
 * WARNING: Deletes all users and spreadsheets from Property Service
 * Use with caution!
 */
function clearAllData() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('USERS');
  props.deleteProperty('SPREADSHEETS');
  props.deleteProperty('ACTIVE_SESSIONS');
  Logger.log('All data cleared from Property Service');
}

// ==================================================================================
// AUTHENTICATION FUNCTIONS
// ==================================================================================
/**
 * LOGIN USER
 * Authenticates user and returns session token if successful
 */
function loginUser(email, password) {
// Sanitize and normalize inputs
  email = normalizeEmail(sanitizeInput(email, 100));
  password = sanitizeInput(password, 100);
  
  if (!email || !password) {
    logActivity(EVENT_TYPES.LOGIN_FAILED, email, 'Login failed: Missing credentials');
    return { success: false, message: 'Email and password are required' };
  }
  
  // Get all users
  const users = getAllUsers();
  const userKey = 'user_' + email;
  const user = users[userKey];
  
  // Check if user exists
  if (!user) {
    logActivity(EVENT_TYPES.LOGIN_FAILED, email, 'Login failed: User not found');
    return { success: false, message: 'Invalid email or password' };
  }
  
  // Check if user is approved
  if (user.registrationApproval !== 'approved') {
    logActivity(EVENT_TYPES.LOGIN_FAILED, email, 'Login attempt - account not approved');
    return { success: false, message: 'Your account is pending administrator approval' };
  }

  // Check if user has admin hold (case-insensitive check)
  if (user.admin_hold && user.admin_hold.toLowerCase() === 'yes') {
    logActivity(EVENT_TYPES.LOGIN_FAILED, email, 'Login attempt - account on admin hold');
    return { success: false, message: 'An administrative hold has been placed on your account. Please contact the administrator to resolve this issue.' };
  }

  // Verify password
  if (!verifyPassword(password, user.passwordHash)) {
    logActivity(EVENT_TYPES.LOGIN_FAILED, email, 'Login failed: Invalid password');
    return { success: false, message: 'Invalid email or password' };
  }
  
  // Check login lockout for General Users (Admins can always log in)
  if (user.userType !== 'Admin') {
    const lockoutStatus = isLoginLocked();
    if (lockoutStatus.locked) {
      logActivity(EVENT_TYPES.LOGIN_FAILED, email, 'Login blocked: System lockout active', {
        lockedBy: lockoutStatus.enabledBy,
        lockedAt: lockoutStatus.enabledAt
      });
      return { 
        success: false, 
        message: 'System temporarily unavailable for maintenance. Please try again later or contact your administrator.' 
      };
    }
  }
  
  // Generate session token
  const sessionToken = generateSessionToken();

  // Store active session
  return withLock('sessions', function() {
    const props = PropertiesService.getScriptProperties();
    const sessionsJson = props.getProperty('ACTIVE_SESSIONS') || '{}';
    const sessions = JSON.parse(sessionsJson);
    
    sessions[sessionToken] = {
      userID: user.userID,
      userType: user.userType,
      loginTime: new Date().toISOString()
    };
    
    props.setProperty('ACTIVE_SESSIONS', JSON.stringify(sessions));
    
    // Log successful login
    logActivity(EVENT_TYPES.LOGIN_SUCCESS, email, 'User logged in successfully', {
      userType: user.userType
    });
    
    return {
      success: true,
      message: 'Login successful',
      sessionToken: sessionToken,
      userEmail: user.userID,
      userType: user.userType,
      isAdmin: user.userType === 'Admin'
    };
  });
}

/**
 * LOGOUT USER
 * Removes session token and releases any checked-out spreadsheets
 */
function logoutUser(sessionToken) {
  if (!sessionToken) {
    return { success: false, message: 'No session token provided' };
  }
  
  return withLock('logout', function() {
    const props = PropertiesService.getScriptProperties();
    
    // Get current sessions
    const sessionsJson = props.getProperty('ACTIVE_SESSIONS') || '{}';
    const sessions = JSON.parse(sessionsJson);
    
    // Get user info before deleting session
    const session = sessions[sessionToken];
    if (!session) {
      // Session already gone (cleared by admin or expired)
      // This is still a successful logout from user's perspective
      return { 
        success: true, 
        message: 'Session already ended',
        alreadyCleared: true,
        releasedSheets: []
      };
    }

    const userEmail = session.userID;
    
    // Remove session
    delete sessions[sessionToken];
    props.setProperty('ACTIVE_SESSIONS', JSON.stringify(sessions));
    
    // Release any spreadsheets checked out by this user
    const spreadsheets = getAllSpreadsheets();
    let releasedSheets = [];

    Object.keys(spreadsheets).forEach(function(sheetKey) {
      if (spreadsheets[sheetKey].currentUser === userEmail) {
        const sheetTitle = spreadsheets[sheetKey].sheetTitle;
        const sheetID = spreadsheets[sheetKey].sheetID;
        
        // Clear currentUser and checkedOutTime
        spreadsheets[sheetKey].currentUser = '';
        spreadsheets[sheetKey].checkedOutTime = null;
        releasedSheets.push(sheetTitle);
        
        // Remove Google Sheets permissions
        try {
          const ss = SpreadsheetApp.openById(sheetID);
          ss.removeEditor(userEmail);
          ss.removeViewer(userEmail);
          
          // Log permission removal
          logActivity(EVENT_TYPES.PERMISSION_REMOVED, userEmail, 
            'Permissions removed on logout: ' + sheetTitle, 
            { sheetID: sheetID, sheetTitle: sheetTitle }
          );
        } catch (e) {
          Logger.log('Error removing permissions during logout: ' + e.toString());
        }
      }
    });

    saveAllSpreadsheets(spreadsheets);
    
    // Log logout
    logActivity(EVENT_TYPES.LOGOUT, userEmail, 'User logged out', {
      releasedSheets: releasedSheets
    });
    
    return {
      success: true,
      message: 'Logout successful',
      releasedSheets: releasedSheets
    };
  });
}

/**
 * RESET USER PASSWORD
 * Allows logged-in users to change their own password
 * Requires current password verification for security
 */
function resetUserPassword(currentPassword, newPassword, sessionToken) {
  // Sanitize inputs
  currentPassword = sanitizeInput(currentPassword, 100);
  newPassword = sanitizeInput(newPassword, 100);
  
  // Verify session
  const session = verifySessionToken(sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Invalid session. Please log in again.' };
  }
  
  const userEmail = normalizeEmail(session.userID);

  // Validate inputs
  if (!currentPassword || !newPassword) {
    return { success: false, message: 'All fields are required' };
  }
  
  // Validate new password length
  if (newPassword.length < 8) {
    return { success: false, message: 'New password must be at least 8 characters long' };
  }
  
  // Check that new password is different from current
  if (currentPassword === newPassword) {
    return { success: false, message: 'New password must be different from current password' };
  }
  
  try {
    const users = getAllUsers();
    const userKey = 'user_' + userEmail;
    const user = users[userKey];
    
    if (!user) {
      return { success: false, message: 'User not found' };
    }
    
    // Verify current password
    if (!verifyPassword(currentPassword, user.passwordHash)) {
      logActivity(
        EVENT_TYPES.LOGIN_FAILED,
        userEmail,
        'Password reset failed: incorrect current password'
      );
      return { success: false, message: 'Current password is incorrect' };
    }
    
    // Hash and save new password
    users[userKey].passwordHash = hashPassword(newPassword);
    users[userKey].passwordChangedDate = new Date().toISOString();
    
    // Save changes
    saveAllUsers(users);
    
    // Log successful password change
    logActivity(
      EVENT_TYPES.USER_MODIFIED,
      userEmail,
      'User changed their password',
      { method: 'resetUserPassword' }
    );
    
    return { 
      success: true, 
      message: 'Password changed successfully! Please use your new password next time you log in.' 
    };
    
  } catch (error) {
    Logger.log('Error in resetUserPassword: ' + error.toString());
    return { success: false, message: 'Error changing password: ' + error.message };
  }
}

/**
 * VERIFY SESSION TOKEN
 * Checks if a session token is valid
 */
function verifySessionToken(sessionToken) {
  if (!sessionToken) {
    return { valid: false, message: 'No session token provided' };
  }
  
  const props = PropertiesService.getScriptProperties();
  const sessionsJson = props.getProperty('ACTIVE_SESSIONS') || '{}';
  const sessions = JSON.parse(sessionsJson);
  
  const session = sessions[sessionToken];
  
  if (!session) {
    return { valid: false, message: 'Invalid or expired session' };
  }
  
  return {
    valid: true,
    userID: session.userID,
    userType: session.userType,
    isAdmin: session.userType === 'Admin'
  };
}

// ==================================================================================
// ACTIVITY LOGGING SYSTEM
// ==================================================================================

/**
 * LOG ACTIVITY
 * Records system events for admin tracking
 * Maintains a rotating log of the last 200 events
 */
function logActivity(eventType, userEmail, details, additionalData) {
  return withLock('activity_log', function() {
    const props = PropertiesService.getScriptProperties();
    const logJson = props.getProperty('ACTIVITY_LOG') || '[]';
    let activityLog = JSON.parse(logJson);
    
    // Create log entry
    const logEntry = {
      timestamp: new Date().toISOString(),
      eventType: eventType,
      userEmail: userEmail || 'System',
      details: details,
      additionalData: additionalData || null
    };
    
    // Add to beginning of array (most recent first)
    activityLog.unshift(logEntry);
    
    // Keep only last MAX entries to manage storage
    if (activityLog.length > MAX_ACTIVITY_LOG_ENTRIES) {
      activityLog = activityLog.slice(0, MAX_ACTIVITY_LOG_ENTRIES);
    }

    // Save back to properties
    props.setProperty('ACTIVITY_LOG', JSON.stringify(activityLog));
    
    return logEntry;
  });
}

/**
 * GET ACTIVITY LOG
 * Retrieves activity log with optional filtering
 */
function getActivityLog(limit, eventTypeFilter, userEmailFilter) {
  limit = limit || 50;  // Default to last 50 events
  
  const props = PropertiesService.getScriptProperties();
  const logJson = props.getProperty('ACTIVITY_LOG') || '[]';
  let activityLog = JSON.parse(logJson);

  Logger.log('getActivityLog called - found ' + activityLog.length + ' entries'); // ← ADD THIS

  // Apply filters if provided
  if (eventTypeFilter) {
    activityLog = activityLog.filter(function(entry) {
      return entry.eventType === eventTypeFilter;
    });
  }
  
  if (userEmailFilter) {
    activityLog = activityLog.filter(function(entry) {
      return entry.userEmail === userEmailFilter;
    });
  }
  
  // Return limited results
  return activityLog.slice(0, limit);
}

/**
 * CLEAR OLD ACTIVITY LOGS
 * Optional maintenance function to clear logs older than X days
 */
function clearOldActivityLogs(daysToKeep) {
  daysToKeep = daysToKeep || 30;
  
  return withLock('activity_log', function() {
    const props = PropertiesService.getScriptProperties();
    const logJson = props.getProperty('ACTIVITY_LOG') || '[]';
    let activityLog = JSON.parse(logJson);
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    // Filter out old entries
    activityLog = activityLog.filter(function(entry) {
      return new Date(entry.timestamp) > cutoffDate;
    });
    
    props.setProperty('ACTIVITY_LOG', JSON.stringify(activityLog));
    
    return { 
      success: true, 
      message: 'Cleared logs older than ' + daysToKeep + ' days',
      remainingEntries: activityLog.length
    };
  });
}

//==============================================================================
//MIDNIGHT CLEAN UP
//==============================================================================
/**
 * MIDNIGHT CLEANUP - Automated daily reset
 * This function should be triggered daily at midnight via Apps Script trigger
 * Clears all sessions AND releases all spreadsheets with permission removal
 */
function midnightCleanup() {
  try {
    Logger.log('Starting midnight cleanup...');
    
    return withLock('midnight_cleanup', function() {
      const props = PropertiesService.getScriptProperties();
      
      // 1. Get current sessions count
      const sessionsJson = props.getProperty('ACTIVE_SESSIONS') || '{}';
      const sessions = JSON.parse(sessionsJson);
      const sessionCount = Object.keys(sessions).length;
      
      // 2. Clear all sessions
      props.setProperty('ACTIVE_SESSIONS', '{}');
      Logger.log('Cleared ' + sessionCount + ' active sessions');
      
      // 3. Release ALL spreadsheets and remove permissions
      const spreadsheets = getAllSpreadsheets();
      const users = getAllUsers();
      let sheetsReleased = 0;
      let permissionsRemoved = 0;
      const errors = [];
      
      Object.keys(spreadsheets).forEach(function(sheetKey) {
        const sheet = spreadsheets[sheetKey];
        
        if (sheet.currentUser && sheet.currentUser !== '') {
          const userEmail = sheet.currentUser;
          const sheetTitle = sheet.sheetTitle;
          const sheetID = sheet.sheetID;
          
          try {
            // Remove Google Sheets permissions
            const ss = SpreadsheetApp.openById(sheetID);
            ss.removeEditor(userEmail);
            ss.removeViewer(userEmail);
            permissionsRemoved++;
            
            Logger.log('Removed permissions: ' + sheetTitle + ' from ' + userEmail);
          } catch (e) {
            const error = 'Failed to remove permissions for ' + sheetTitle + ': ' + e.message;
            errors.push(error);
            Logger.log('ERROR: ' + error);
          }
          
          // Clear currentUser and checkedOutTime regardless of permission removal success
          sheet.currentUser = '';
          sheet.checkedOutTime = null;
          sheetsReleased++;
        }
      });
      
      // Save updated spreadsheets
      saveAllSpreadsheets(spreadsheets);
      
      // 4. Log the cleanup activity
      logActivity(
        EVENT_TYPES.SYSTEM_CLEANUP,
        'System',
        'Midnight cleanup completed',
        {
          sessionsCleared: sessionCount,
          sheetsReleased: sheetsReleased,
          permissionsRemoved: permissionsRemoved,
          errors: errors.length
        }
      );
      
      Logger.log('Midnight cleanup complete:');
      Logger.log('- Sessions cleared: ' + sessionCount);
      Logger.log('- Sheets released: ' + sheetsReleased);
      Logger.log('- Permissions removed: ' + permissionsRemoved);
      Logger.log('- Errors: ' + errors.length);
      
      // 5. Email admin if there were errors
      if (errors.length > 0) {
        emailAdminMidnightErrors(errors, sessionCount, sheetsReleased, permissionsRemoved);
      }
      
      return {
        success: true,
        sessionsCleared: sessionCount,
        sheetsReleased: sheetsReleased,
        permissionsRemoved: permissionsRemoved,
        errors: errors
      };
    });
    
  } catch (error) {
    Logger.log('CRITICAL ERROR in midnightCleanup: ' + error.toString());
    
    // Log critical error
    logActivity(
      EVENT_TYPES.SYSTEM_CLEANUP,
      'System',
      'Midnight cleanup FAILED: ' + error.message,
      { error: error.toString() }
    );
    
    // Email admin about critical failure
    try {
      MailApp.sendEmail({
        to: ADMIN_CONTACT_EMAIL,
        subject: SYSTEM_NAME + ' - CRITICAL: Midnight Cleanup Failed',
        body: 'The midnight cleanup process failed with error:\n\n' + error.toString() + '\n\nPlease check the system immediately.'
      });
    } catch (e) {
      Logger.log('Failed to send critical error email: ' + e.toString());
    }
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * EMAIL ADMIN ABOUT MIDNIGHT CLEANUP ERRORS
 * Sends detailed error report to admin
 */
function emailAdminMidnightErrors(errors, sessionCount, sheetsReleased, permissionsRemoved) {
  try {
    const subject = SYSTEM_NAME + ' - Midnight Cleanup Errors';
    
    let errorDetails = '';
    errors.forEach(function(err) {
      errorDetails += '• ' + err + '\n';
    });
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #FF9800, #FFA726); padding: 20px; border-radius: 10px 10px 0 0;">
          <h2 style="color: white; margin: 0;">⚠️ Midnight Cleanup Errors</h2>
        </div>
        
        <div style="background-color: #f5f5f5; padding: 25px; border-radius: 0 0 10px 10px;">
          <h3 style="color: #FF9800; margin-top: 0;">Cleanup Summary</h3>
          <p><strong>Timestamp:</strong> ${new Date().toLocaleString()}</p>
          <p><strong>Sessions Cleared:</strong> ${sessionCount}</p>
          <p><strong>Sheets Released:</strong> ${sheetsReleased}</p>
          <p><strong>Permissions Removed:</strong> ${permissionsRemoved}</p>
          <p><strong>Errors:</strong> ${errors.length}</p>
          
          <h3 style="color: #FF9800; margin-top: 25px;">Error Details</h3>
          <div style="background-color: white; padding: 15px; border-left: 4px solid #FF9800; border-radius: 5px;">
            <pre style="white-space: pre-wrap; margin: 0; font-family: monospace;">${errorDetails}</pre>
          </div>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 25px 0;">
          
          <p style="color: #666;">
            The midnight cleanup process completed but encountered some errors while removing permissions.
            Most operations succeeded, but some spreadsheet permissions may need manual review.
          </p>
          
          <p style="font-size: 0.85em; color: #999; margin-top: 20px;">
            This notification was sent from ${SYSTEM_NAME}
          </p>
        </div>
      </div>
    `;
    
    const plainBody = `
Midnight Cleanup Errors - ${SYSTEM_NAME}

Cleanup Summary:
Timestamp: ${new Date().toLocaleString()}
Sessions Cleared: ${sessionCount}
Sheets Released: ${sheetsReleased}
Permissions Removed: ${permissionsRemoved}
Errors: ${errors.length}

Error Details:
${errorDetails}

The midnight cleanup process completed but encountered some errors while removing permissions.
Most operations succeeded, but some spreadsheet permissions may need manual review.
    `;
    
    MailApp.sendEmail({
      to: ADMIN_CONTACT_EMAIL,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });
    
    Logger.log('Error notification email sent to admin');
    
  } catch (e) {
    Logger.log('Failed to send midnight error email: ' + e.toString());
  }
}

// ==================================================================================
// USER REGISTRATION FUNCTIONALITY
// ==================================================================================
/**
 * REGISTER NEW USER
 * Creates a new user account with pending approval status
 */
function registerNewUser(email, password, fullName) {
  // Sanitize and normalize inputs
  email = normalizeEmail(sanitizeInput(email, 100));
  password = sanitizeInput(password, 100);
  fullName = sanitizeInput(fullName, 100);
  
  // Validate inputs
  if (!email || !password || !fullName) {
    return { 
      success: false, 
      message: 'All fields are required' 
    };
  }
  
  // Validate email format
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    return { 
      success: false, 
      message: 'Please enter a valid email address' 
    };
  }
  
  // Validate Gmail only
  if (!email.toLowerCase().endsWith('@gmail.com')) {
    return {
      success: false,
      message: 'Only Gmail accounts (@gmail.com) are allowed to register'
    };
  }
  
  // Validate password strength (minimum 8 characters)
  if (password.length < 8) {
    return {
      success: false,
      message: 'Password must be at least 8 characters long'
    };
  }
  
  return withLock('users', function() {
    const users = getAllUsers();
    const userKey = 'user_' + email;
    
    // Check if user already exists
    if (users[userKey]) {
      return {
        success: false,
        message: 'An account with this email already exists'
      };
    }
    
    // Create new user with updated structure
    users[userKey] = {
      userID: email,
      fullName: fullName,
      passwordHash: hashPassword(password),
      registrationApproval: 'pending',
      userType: '',
      registeredDate: new Date().toISOString(),
      assigned_spreadsheets: [],
      admin_hold: 'No'
    };
    
    // Save users
    saveAllUsers(users);
    
    // Log registration
    logActivity(
      EVENT_TYPES.USER_REGISTERED,
      email,
      'New user registered: ' + fullName,
      { fullName: fullName }
    );
    
    // Send email notification to admin
    try {
      const subject = SYSTEM_NAME + ' - New User Registration';
      
      const htmlBody = `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #2196F3, #4CAF50); padding: 20px; border-radius: 10px 10px 0 0;">
            <h2 style="color: white; margin: 0;">New User Registration</h2>
          </div>
          
          <div style="background-color: #f5f5f5; padding: 25px; border-radius: 0 0 10px 10px;">
            <h3 style="color: #2196F3; margin-top: 0;">User Information</h3>
            <p><strong>Name:</strong> ${fullName}</p>
            <p><strong>Email:</strong> ${email}</p>
            <p><strong>Registration Date:</strong> ${new Date().toLocaleString()}</p>
            <p><strong>Status:</strong> Pending Approval</p>
            
            <hr style="border: none; border-top: 1px solid #ddd; margin: 25px 0;">
            
            <p style="color: #666;">
              Please log in to the admin panel to approve or reject this registration request.
            </p>
            
            <p style="font-size: 0.85em; color: #999; margin-top: 20px;">
              This notification was sent from ${SYSTEM_NAME}
            </p>
          </div>
        </div>
      `;
      
      const plainBody = `
New User Registration - ${SYSTEM_NAME}

User Information:
Name: ${fullName}
Email: ${email}
Registration Date: ${new Date().toLocaleString()}
Status: Pending Approval

Please log in to the admin panel to approve or reject this registration request.
      `;
      
      MailApp.sendEmail({
        to: ADMIN_CONTACT_EMAIL,
        subject: subject,
        body: plainBody,
        htmlBody: htmlBody
      });
    } catch (e) {
      Logger.log('Error sending registration notification email: ' + e.toString());
      // Don't fail registration if email fails
    }
    
    return {
      success: true,
      message: 'Registration successful! Your account is pending approval. You will receive an email notification once approved.'
    };
  });
}

// ==================================================================================
// SPREADSHEET MANAGEMENT FUNCTIONS
// ==================================================================================
/**
 * GET SPREADSHEET DETAILS
 * Retrieves detailed information about a specific spreadsheet for the manage modal
 */
function getSpreadsheetDetails(sheetID, sessionToken) {
  try {
    // Verify admin session
    const session = verifySessionToken(sessionToken);
    if (!session.valid || !session.isAdmin) {
      return { success: false, message: 'Admin access required' };
    }
    
    // Sanitize input
    sheetID = sanitizeInput(sheetID, 100);
    
    if (!sheetID) {
      return { success: false, message: 'Spreadsheet ID is required' };
    }
    
    // Get spreadsheet from Property Service
    const spreadsheets = getAllSpreadsheets();
    const sheetKey = 'sheet_' + sheetID;
    const spreadsheet = spreadsheets[sheetKey];
    
    if (!spreadsheet) {
      return { success: false, message: 'Spreadsheet not found' };
    }
    
    return {
      success: true,
      spreadsheet: spreadsheet
    };
    
  } catch (error) {
    Logger.log('Error in getSpreadsheetDetails: ' + error.toString());
    return { success: false, message: 'Error loading spreadsheet details: ' + error.message };
  }
}

/**
 * UPDATE SPREADSHEET
 * Updates spreadsheet details from the manage modal
 */
function updateSpreadsheet(sheetID, updatedData, sessionToken) {
  try {
    // Verify admin session
    const session = verifySessionToken(sessionToken);
    if (!session.valid || !session.isAdmin) {
      return { success: false, message: 'Admin access required' };
    }
    
    // Sanitize inputs
    sheetID = sanitizeInput(sheetID, 100);
    const sheetTitle = sanitizeInput(updatedData.sheetTitle, 100);
    const access = updatedData.access;
    const adminHold = updatedData.admin_hold;
    const group = sanitizeInput(updatedData.group || "", 10);
    const currentUser = sanitizeInput(updatedData.currentUser || "", 100);
    
    // Validate required fields
    if (!sheetID || !sheetTitle) {
      return { success: false, message: 'Spreadsheet ID and Title are required' };
    }
    
    // Validate access type
    if (access !== 'Viewer' && access !== 'Editor') {
      return { success: false, message: 'Invalid access type' };
    }
    
    // Validate admin hold
    if (adminHold !== 'Yes' && adminHold !== 'No') {
      return { success: false, message: 'Invalid admin hold value' };
    }
    
    return withLock('spreadsheets', function() {
      const spreadsheets = getAllSpreadsheets();
      const sheetKey = 'sheet_' + sheetID;
      const spreadsheet = spreadsheets[sheetKey];
      
      if (!spreadsheet) {
        return { success: false, message: 'Spreadsheet not found' };
      }
      
      // Track if currentUser was cleared (released)
      const userWasReleased = spreadsheet.currentUser && !currentUser;
      const releasedUser = spreadsheet.currentUser;
      
      // Update spreadsheet object
      spreadsheet.sheetTitle = sheetTitle;
      spreadsheet.access = access;
      spreadsheet.admin_hold = adminHold;
      spreadsheet.group = group;
      spreadsheet.currentUser = currentUser;
      spreadsheet.editDate = new Date().toISOString();
      spreadsheet.editedBy = session.userID;
      
      // If currentUser was cleared, also remove checkedOutTime and Google permissions
      if (userWasReleased) {
        spreadsheet.checkedOutTime = null;
        
        // Remove Google Sheets permissions
        try {
          const ss = SpreadsheetApp.openById(sheetID);
          ss.removeEditor(releasedUser);
          ss.removeViewer(releasedUser);
          
          // Log permission removal
          logActivity(EVENT_TYPES.PERMISSION_REMOVED, session.userID, 
            'Admin released spreadsheet "' + sheetTitle + '" from ' + releasedUser, 
            { sheetID: sheetID, sheetTitle: sheetTitle, releasedUser: releasedUser }
          );
        } catch (e) {
          Logger.log('Error removing permissions during release: ' + e.toString());
        }
      }
      
      // Save updated spreadsheets
      spreadsheets[sheetKey] = spreadsheet;
      saveAllSpreadsheets(spreadsheets);
      
      // Log activity
      logActivity(
        EVENT_TYPES.SPREADSHEET_MODIFIED,
        session.userID,
        'Admin edited spreadsheet: ' + sheetTitle + (userWasReleased ? ' (released from ' + releasedUser + ')' : ''),
        { sheetID: sheetID, sheetTitle: sheetTitle }
      );
      
      return {
        success: true,
        message: 'Spreadsheet updated successfully!',
        releasedUser: userWasReleased ? releasedUser : null
      };
    });
    
  } catch (error) {
    Logger.log('Error in updateSpreadsheet: ' + error.toString());
    return { success: false, message: 'Error updating spreadsheet: ' + error.message };
  }
}

function addSpreadsheet(spreadsheetObj, sessionToken) {
  try {
    // Verify admin session using existing function
    const session = verifySessionToken(sessionToken);
    if (!session.valid || !session.isAdmin) {
      return { success: false, message: 'Admin access required' };
    }
    
    // Validate input
    if (!spreadsheetObj || !spreadsheetObj.sheetID || !spreadsheetObj.sheetTitle) {
      return { success: false, message: 'Missing required fields' };
    }
    
    // Sanitize inputs
    const sheetID = sanitizeInput(spreadsheetObj.sheetID, 100);
    const sheetTitle = sanitizeInput(spreadsheetObj.sheetTitle, 100);
    const access = spreadsheetObj.access || 'Viewer';
    const currentUser = sanitizeInput(spreadsheetObj.currentUser || "", 100);
    const adminHold = spreadsheetObj.admin_hold || 'No';
    const group = sanitizeInput(spreadsheetObj.group || "", 10);
    
    // Validate access type
    if (access !== 'Viewer' && access !== 'Editor') {
      return { success: false, message: 'Invalid access type' };
    }
    
    // Validate admin hold
    if (adminHold !== 'Yes' && adminHold !== 'No') {
      return { success: false, message: 'Invalid admin hold value' };
    }
    
    // Verify spreadsheet exists and is accessible
    try {
      const sheet = SpreadsheetApp.openById(sheetID);
      const sheetName = sheet.getName(); // This will throw error if no access
    } catch (e) {
      return { 
        success: false, 
        message: 'Cannot access spreadsheet. Please ensure the spreadsheet ID is correct and this script has permission to access it.' 
      };
    }
    
    return withLock('spreadsheets', function() {
      // Use existing getAllSpreadsheets() function
      const spreadsheets = getAllSpreadsheets();
      const sheetKey = 'sheet_' + sheetID;
      
      // Check if spreadsheet already exists
      if (spreadsheets[sheetKey]) {
        return { success: false, message: 'This spreadsheet is already in the system' };
      }
      
      // Create spreadsheet object
      spreadsheets[sheetKey] = {
        sheetID: sheetID,
        sheetTitle: sheetTitle,
        access: access,
        currentUser: currentUser,
        admin_hold: adminHold,
        group: group,
        addedDate: new Date().toISOString(),
        addedBy: session.userID
      };
      
      // Save using existing function
      saveAllSpreadsheets(spreadsheets);
      
      // Log activity
      logActivity(
        EVENT_TYPES.SPREADSHEET_ADDED,
        session.userID,
        'Added spreadsheet: ' + sheetTitle + ' (ID: ' + sheetID + ')'
      );
      
      return { 
        success: true, 
        message: 'Spreadsheet "' + sheetTitle + '" added successfully!' 
      };
    });
    
  } catch (error) {
    Logger.log('Error in addSpreadsheet: ' + error.toString());
    return { success: false, message: 'Error adding spreadsheet: ' + error.message };
  }
}

/**
 * DELETE SPREADSHEET
 * Removes a spreadsheet from the system
 */
function deleteSpreadsheet(sheetID, sessionToken) {
  try {
    // Verify admin session
    const session = verifySessionToken(sessionToken);
    if (!session.valid || !session.isAdmin) {
      return { success: false, message: 'Admin access required' };
    }
    
    // Sanitize input
    sheetID = sanitizeInput(sheetID, 100);
    
    if (!sheetID) {
      return { success: false, message: 'Spreadsheet ID is required' };
    }
    
    return withLock('spreadsheets', function() {
      const spreadsheets = getAllSpreadsheets();
      const sheetKey = 'sheet_' + sheetID;
      const spreadsheet = spreadsheets[sheetKey];
      
      if (!spreadsheet) {
        return { success: false, message: 'Spreadsheet not found' };
      }
      
      const sheetTitle = spreadsheet.sheetTitle;
      const currentUser = spreadsheet.currentUser;
      
      // If spreadsheet is checked out, remove Google permissions
      if (currentUser) {
        try {
          const ss = SpreadsheetApp.openById(sheetID);
          ss.removeEditor(currentUser);
          ss.removeViewer(currentUser);
          
          // Log permission removal
          logActivity(EVENT_TYPES.PERMISSION_REMOVED, session.userID, 
            'Permissions removed during delete: ' + sheetTitle + ' from ' + currentUser, 
            { sheetID: sheetID, sheetTitle: sheetTitle, removedFrom: currentUser }
          );
        } catch (e) {
          Logger.log('Error removing permissions during delete: ' + e.toString());
          // Continue with deletion even if permission removal fails
        }
      }
      
      // Delete spreadsheet from system
      delete spreadsheets[sheetKey];
      saveAllSpreadsheets(spreadsheets);
      
      // Log activity
      logActivity(
        EVENT_TYPES.SPREADSHEET_DELETED,
        session.userID,
        'Admin deleted spreadsheet: ' + sheetTitle + ' (ID: ' + sheetID + ')',
        { sheetID: sheetID, sheetTitle: sheetTitle, wasCheckedOut: currentUser ? true : false }
      );
      
      return {
        success: true,
        message: 'Spreadsheet "' + sheetTitle + '" removed from system successfully!'
      };
    });
    
  } catch (error) {
    Logger.log('Error in deleteSpreadsheet: ' + error.toString());
    return { success: false, message: 'Error deleting spreadsheet: ' + error.message };
  }
}

/**
 * Get spreadsheet object by ID
 * @param {string} sheetID - Spreadsheet ID
 * @return {object|null} Spreadsheet object or null if not found
 */
function getSpreadsheetByID(sheetID) {
  const spreadsheets = getAllSpreadsheets();
  const key = 'sheet_' + sheetID;
  return spreadsheets[key] || null;
}

/**
 * GET USER AVAILABLE SPREADSHEETS
 * Returns list of spreadsheets the user can access
 */
function getUserAvailableSpreadsheets(sessionToken) {
  try {
    const session = verifySessionToken(sessionToken);
    if (!session.valid) {
      return { success: false, message: 'Invalid session' };
    }
    
    const userEmail = session.userID;
    const users = getAllUsers();
    const userKey = 'user_' + userEmail;
    const user = users[userKey];
    
    if (!user) {
      return { success: false, message: 'User not found' };
    }
    
    const spreadsheets = getAllSpreadsheets();
    const availableSheets = [];
    
    Object.keys(spreadsheets).forEach(function(key) {
      const sheet = spreadsheets[key];
      
      // Skip if on admin hold
      if (sheet.admin_hold === 'Yes') {
        return;
      }
      
      // Check if user is assigned (by ID or group)
      const isAssigned = user.assigned_spreadsheets && user.assigned_spreadsheets.some(function(assigned) {
        return assigned === sheet.sheetID || (sheet.group && assigned === sheet.group);
      });
      
      if (isAssigned) {
        // FIX: Look up the full name of the current user
        let currentUserDisplay = null;
        if (sheet.currentUser) {
          const currentUserKey = 'user_' + sheet.currentUser;
          const currentUserData = users[currentUserKey];
          // Use full name if available, otherwise fall back to email
          currentUserDisplay = currentUserData && currentUserData.fullName ? 
            currentUserData.fullName : 
            sheet.currentUser;
        }
        
        availableSheets.push({
          id: sheet.sheetID,
          name: sheet.sheetTitle,
          accessType: sheet.access,
          inUse: sheet.currentUser ? true : false,
          currentUser: sheet.currentUser || null,
          currentUserDisplay: currentUserDisplay  // NEW: Full name for display
        });
      }
    });
    
    return {
      success: true,
      spreadsheets: availableSheets
    };
    
  } catch (error) {
    Logger.log('Error in getUserAvailableSpreadsheets: ' + error.toString());
    return { success: false, message: 'Error loading spreadsheets: ' + error.message };
  }
}

/**
 * Get ALL currently checked-out spreadsheets (system-wide view)
 * This is for the "Currently In Use" display only
 * @param {string} sessionToken - User's session token
 * @return {Object} Result with array of in-use spreadsheets
 */
function getAllCheckedOutSpreadsheets(sessionToken) {
  try {
    // Verify user is logged in
    const session = verifySessionToken(sessionToken);
    if (!session.valid) {
      return {
        success: false,
        message: 'Invalid session. Please log in again.'
      };
    }    
    const spreadsheets = getAllSpreadsheets();
    const users = getAllUsers();
    const inUseSheets = [];
    
    // Loop through all spreadsheets
    for (const key in spreadsheets) {
      if (key.startsWith('sheet_')) {
        const sheet = spreadsheets[key];
        
        // Only include if someone is currently using it
        if (sheet.currentUser && sheet.currentUser !== '') {
          // Look up the full name of the current user
          const currentUserKey = 'user_' + sheet.currentUser;
          const currentUserData = users[currentUserKey];
          const currentUserDisplay = currentUserData && currentUserData.fullName ? 
            currentUserData.fullName : 
            sheet.currentUser;
          
          inUseSheets.push({
            id: sheet.sheetID,
            name: sheet.sheetTitle,
            currentUser: sheet.currentUser,
            currentUserDisplay: currentUserDisplay,
            accessType: sheet.access
          });
        }
      }
    }
    
    // Sort by sheet name for consistent display
    inUseSheets.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });
    
    return {
      success: true,
      spreadsheets: inUseSheets
    };
    
  } catch (error) {
    Logger.log('Error in getAllCheckedOutSpreadsheets: ' + error.toString());
    return {
      success: false,
      message: 'Error retrieving checked-out spreadsheets: ' + error.message
    };
  }
}

/**
 * RELEASE ALL USER SPREADSHEETS
 * Releases all spreadsheets currently checked out by a specific user
 * Removes Google Sheets permissions and clears checkout status
 */
function releaseUserSpreadsheets(userEmail, sessionToken) {
  // Normalize email
  userEmail = normalizeEmail(userEmail);
  
  // Verify admin session
  const session = verifySessionToken(sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Invalid session. Please log in again.' };
  }
  
  if (!session.isAdmin) {
    return { success: false, message: 'Admin access required' };
  }
  
  if (!userEmail) {
    return { success: false, message: 'User email is required' };
  }
  
  try {
    return withLock('spreadsheets', function() {
      const spreadsheets = getAllSpreadsheets();
      const releasedSheets = [];
      const errors = [];
      
      // Find all spreadsheets checked out by this user
      Object.keys(spreadsheets).forEach(function(sheetKey) {
        const sheet = spreadsheets[sheetKey];
        
        if (sheet.currentUser === userEmail) {
          // Try to remove Google Sheets permissions
          try {
            const ss = SpreadsheetApp.openById(sheet.sheetID);
            ss.removeEditor(userEmail);
            ss.removeViewer(userEmail);
          } catch (e) {
            errors.push({
              sheetTitle: sheet.sheetTitle,
              error: e.message
            });
            Logger.log('Error removing permissions for ' + sheet.sheetTitle + ': ' + e.toString());
          }
          
          // Clear checkout status regardless of permission removal success
          sheet.currentUser = '';
          sheet.checkedOutTime = null;
          
          releasedSheets.push({
            sheetID: sheet.sheetID,
            sheetTitle: sheet.sheetTitle
          });
        }
      });
      
      // Save updated spreadsheets
      if (releasedSheets.length > 0) {
        saveAllSpreadsheets(spreadsheets);
        
        // Log activity
        const sheetNames = releasedSheets.map(function(s) { return s.sheetTitle; }).join(', ');
        logActivity(
          EVENT_TYPES.PERMISSION_REMOVED,
          session.userID,
          'Admin released ' + releasedSheets.length + ' spreadsheet(s) from ' + userEmail + ': ' + sheetNames,
          { 
            targetUser: userEmail, 
            sheetsReleased: releasedSheets.length,
            sheets: releasedSheets
          }
        );
      }
      
      // Build response message
      let message = '';
      if (releasedSheets.length === 0) {
        message = 'No spreadsheets were checked out by this user.';
      } else {
        message = 'Released ' + releasedSheets.length + ' spreadsheet(s):\n\n';
        releasedSheets.forEach(function(sheet) {
          message += '• ' + sheet.sheetTitle + '\n';
        });
        
        if (errors.length > 0) {
          message += '\n⚠️ ' + errors.length + ' permission error(s) occurred.\nSpreadsheets were released but some Google permissions may need manual review.';
        }
      }
      
      return {
        success: true,
        message: message,
        releasedCount: releasedSheets.length,
        releasedSheets: releasedSheets,
        errors: errors
      };
    });
    
  } catch (error) {
    Logger.log('Error in releaseUserSpreadsheets: ' + error.toString());
    return { success: false, message: 'Error releasing spreadsheets: ' + error.message };
  }
}

/**
 * GET USER CHECKED OUT SPREADSHEETS COUNT
 * Returns count of spreadsheets currently checked out by a user
 */
function getUserCheckedOutCount(userEmail) {
  userEmail = normalizeEmail(userEmail);
  const spreadsheets = getAllSpreadsheets();
  let count = 0;
  
  Object.keys(spreadsheets).forEach(function(key) {
    if (spreadsheets[key].currentUser === userEmail) {
      count++;
    }
  });
  
  return count;
}

// ==================================================================================
// USER SECTION FUNCTIONS
// ==================================================================================
/**
 * REQUEST SPREADSHEET ACCESS
 * Allows a user to request access to a spreadsheet they've been assigned
 * @param {string} sheetID - The spreadsheet ID to request access to
 * @param {string} sessionToken - User's session token
 * @return {Object} Result with success status, message, and spreadsheet URL
 */
function requestSpreadsheetAccess(sheetID, sessionToken) {
  try {
    // 1. Verify user session
    const session = verifySessionToken(sessionToken);
    if (!session.valid) {
      return { success: false, message: 'Invalid session. Please login again.' };
    }
    
    const userEmail = session.userID;
    
    // 2. Validate spreadsheet selection
    if (!sheetID) {
      return { success: false, message: 'Please select a spreadsheet' };
    }
    
    // Sanitize input
    sheetID = sanitizeInput(sheetID, 100);
    
    return withLock('spreadsheet_access', function() {
      // 3. Find spreadsheet in system
      const spreadsheets = getAllSpreadsheets();
      const sheetKey = 'sheet_' + sheetID;
      const spreadsheet = spreadsheets[sheetKey];
      
      if (!spreadsheet) {
        logActivity(EVENT_TYPES.SPREADSHEET_ACCESS_DENIED, userEmail, 
          'Access denied: Spreadsheet not found (ID: ' + sheetID + ')');
        return { success: false, message: 'Spreadsheet not found in system' };
      }
      
      // 4. Check spreadsheet admin_hold
      if (spreadsheet.admin_hold === 'Yes') {
        logActivity(EVENT_TYPES.SPREADSHEET_ACCESS_DENIED, userEmail, 
          'Access denied: Spreadsheet on admin hold - ' + spreadsheet.sheetTitle);
        return { 
          success: false, 
          message: 'This spreadsheet is temporarily unavailable. Please contact the administrator.' 
        };
      }
      
      // 5. Check if spreadsheet is already checked out by another user
      if (spreadsheet.currentUser && spreadsheet.currentUser !== userEmail) {
        logActivity(EVENT_TYPES.SPREADSHEET_ACCESS_DENIED, userEmail, 
          'Access denied: Spreadsheet in use by ' + spreadsheet.currentUser + ' - ' + spreadsheet.sheetTitle);
        return { 
          success: false, 
          message: 'This spreadsheet is currently in use by ' + spreadsheet.currentUser + '. Please try again later.' 
        };
      }
      
      // 6. Check if user already has access
      if (spreadsheet.currentUser === userEmail) {
        const spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.sheetID;
        return { 
          success: true, 
          message: 'You already have access to this spreadsheet!',
          spreadsheetUrl: spreadsheetUrl,
          accessType: spreadsheet.access,
          alreadyHasAccess: true
        };
      }
      
      // 7. Get user's assigned spreadsheets
      const users = getAllUsers();
      const userKey = 'user_' + userEmail;
      const user = users[userKey];
      
      if (!user) {
        return { success: false, message: 'User not found' };
      }
      
      // 8. Check if user has any assigned spreadsheets
      if (!user.assigned_spreadsheets || user.assigned_spreadsheets.length === 0) {
        logActivity(EVENT_TYPES.SPREADSHEET_ACCESS_DENIED, userEmail, 
          'Access denied: No spreadsheets assigned to user - requested ' + spreadsheet.sheetTitle);
        return { 
          success: false, 
          message: 'You have no assigned spreadsheets. Please contact the administrator to request access.' 
        };
      }
      
      // 9. Check if user is assigned to this spreadsheet (by ID or group)
      const hasAccess = user.assigned_spreadsheets.some(function(assigned) {
        // Check if assigned matches spreadsheet ID
        if (assigned === sheetID) {
          return true;
        }
        // Check if assigned matches spreadsheet group
        if (spreadsheet.group && assigned === spreadsheet.group) {
          return true;
        }
        return false;
      });
      
      if (!hasAccess) {
        logActivity(EVENT_TYPES.SPREADSHEET_ACCESS_DENIED, userEmail, 
          'Access denied: User not assigned to spreadsheet - ' + spreadsheet.sheetTitle);
        return { 
          success: false, 
          message: 'This spreadsheet has not been assigned to you. Please contact the administrator if this is in error.' 
        };
      }
      
      // 10. Grant Google Sheets permission
      try {
        const ss = SpreadsheetApp.openById(sheetID);
        
        // Grant appropriate access based on spreadsheet settings
        if (spreadsheet.access === 'Editor') {
          ss.addEditor(userEmail);
        } else {
          ss.addViewer(userEmail);
        }
        
        // 11. Update spreadsheet object
        spreadsheet.currentUser = userEmail;
        spreadsheet.checkedOutTime = new Date().toISOString();
        
        // Save updated spreadsheets
        spreadsheets[sheetKey] = spreadsheet;
        saveAllSpreadsheets(spreadsheets);
        
        // 12. Log activity
        logActivity(EVENT_TYPES.SPREADSHEET_CHECKOUT, userEmail, 
          'Checked out spreadsheet: ' + spreadsheet.sheetTitle + ' (' + spreadsheet.access + ' access)',
          { 
            sheetID: sheetID, 
            sheetTitle: spreadsheet.sheetTitle,
            accessType: spreadsheet.access 
          }
        );
        
        // 13. Return success with URL
        const spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/' + sheetID;
        
        return {
          success: true,
          message: 'Access granted! Opening spreadsheet...',
          spreadsheetUrl: spreadsheetUrl,
          accessType: spreadsheet.access,
          sheetTitle: spreadsheet.sheetTitle
        };
        
      } catch (e) {
        Logger.log('Error granting spreadsheet access: ' + e.toString());
        logActivity(EVENT_TYPES.SPREADSHEET_ACCESS_DENIED, userEmail, 
          'Access denied: Error granting permissions - ' + spreadsheet.sheetTitle + ' - ' + e.toString());
        return { 
          success: false, 
          message: 'Error granting access to spreadsheet. The spreadsheet may not exist or permissions may be restricted.' 
        };
      }
    });
    
  } catch (error) {
    Logger.log('Error in requestSpreadsheetAccess: ' + error.toString());
    return { 
      success: false, 
      message: 'Error processing request: ' + error.message 
    };
  }
}

/**
 * SEND CONTACT MESSAGE TO ADMIN
 * Sends an email to the admin with user's message
 */
function sendContactMessage(userEmail, userName, message, sessionToken) {
  // Sanitize inputs
  userEmail = sanitizeInput(userEmail, 100);
  userName = sanitizeInput(userName, 100);
  message = sanitizeInput(message, 2000);
  
  // Verify session if provided (optional - allows logged in or non-logged in users)
  let isAuthenticated = false;
  if (sessionToken) {
    const session = verifySessionToken(sessionToken);
    isAuthenticated = session.valid;
  }
  
  // Validate inputs
  if (!userEmail || !userName || !message) {
    return { 
      success: false, 
      message: 'All fields are required' 
    };
  }
  
  // Validate email format
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(userEmail)) {
    return { 
      success: false, 
      message: 'Please enter a valid email address' 
    };
  }
  
  try {
    // Compose email
    const subject = SYSTEM_NAME + ' - User Contact Request from ' + userName;
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #2196F3, #4CAF50); padding: 20px; border-radius: 10px 10px 0 0;">
          <h2 style="color: white; margin: 0;">Contact Request</h2>
        </div>
        
        <div style="background-color: #f5f5f5; padding: 25px; border-radius: 0 0 10px 10px;">
          <h3 style="color: #2196F3; margin-top: 0;">User Information</h3>
          <p><strong>Name:</strong> ${userName}</p>
          <p><strong>Email:</strong> ${userEmail}</p>
          <p><strong>Status:</strong> ${isAuthenticated ? 'Authenticated User' : 'Not Logged In'}</p>
          <p><strong>Timestamp:</strong> ${new Date().toLocaleString()}</p>
          
          <h3 style="color: #2196F3; margin-top: 25px;">Message</h3>
          <div style="background-color: white; padding: 15px; border-left: 4px solid #4CAF50; border-radius: 5px;">
            <p style="white-space: pre-wrap; margin: 0;">${message}</p>
          </div>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 25px 0;">
          
          <p style="font-size: 0.9em; color: #666;">
            <strong>Reply to:</strong> <a href="mailto:${userEmail}">${userEmail}</a>
          </p>
          
          <p style="font-size: 0.85em; color: #999; margin-top: 20px;">
            This message was sent from ${SYSTEM_NAME}
          </p>
        </div>
      </div>
    `;
    
    const plainBody = `
Contact Request from ${SYSTEM_NAME}

User Information:
Name: ${userName}
Email: ${userEmail}
Status: ${isAuthenticated ? 'Authenticated User' : 'Not Logged In'}
Timestamp: ${new Date().toLocaleString()}

Message:
${message}

---
Reply to: ${userEmail}
    `;
    
    // Send email
    MailApp.sendEmail({
      to: ADMIN_CONTACT_EMAIL,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      replyTo: userEmail
    });
    
    // Log the contact attempt
    logActivity(
      EVENT_TYPES.CONTACT_ADMIN, 
      userEmail, 
      'User contacted admin: ' + userName,
      { userName: userName, messageLength: message.length }
    );
    
    return { 
      success: true, 
      message: 'Your message has been sent to the administrator. They will respond via email.' 
    };
    
  } catch (e) {
    Logger.log('Error sending contact email: ' + e.toString());
    return { 
      success: false, 
      message: 'Error sending message. Please try again or contact the administrator directly at ' + ADMIN_CONTACT_EMAIL 
    };
  }
}

// ==================================================================================
// ADMIN SECTION FUNCTIONS
// ==================================================================================
/**
 * CHECK IF CURRENT SESSION IS SUPER ADMIN
 * Returns true if the session belongs to the super admin
 */
function isSessionSuperAdmin(sessionToken) {
  const session = verifySessionToken(sessionToken);
  if (!session.valid) return false;
  return isSuperAdmin(session.userID);
}

/**
 * CHECK IF USER IS SUPER ADMIN
 * Returns true if the email matches the super admin
 */
function isSuperAdmin(email) {
  if (!email) return false;
  return email.toLowerCase() === SUPER_ADMIN_EMAIL.toLowerCase();
}

/**
 * GET USER DETAILS FOR MANAGE USER MODAL
 */
function getUserDetails(userEmail, sessionToken) {
  // Normalize email for consistent lookup
  userEmail = normalizeEmail(userEmail);

  // Verify session
  const session = verifySessionToken(sessionToken);
  if (!session.valid) {
    throw new Error('Invalid session. Please log in again.');
  }
  
  // Check if requester is admin
  const requesterEmail = session.userID;
  const users = getAllUsers();
  const requesterKey = 'user_' + requesterEmail;
  const requester = users[requesterKey];
  
  if (!requester || requester.userType !== 'Admin') {
    throw new Error('Admin access required');
  }
  
  // SUPER ADMIN PROTECTION - Allow super admin to access their own profile
  if (isSuperAdmin(userEmail) && !isSuperAdmin(requesterEmail)) {
    throw new Error('Access denied: This account is protected and cannot be modified');
  }
  
  // Get user details
  const userKey = 'user_' + userEmail;
  const user = users[userKey];
  
  if (!user) {
    throw new Error('User not found');
  }
  
  return {
    userID: userEmail,
    fullName: user.fullName || '',
    userType: user.userType || 'General User',
    registrationApproval: user.registrationApproval || 'pending',
    registeredDate: user.registeredDate || null,
    assigned_spreadsheets: user.assigned_spreadsheets || [],
    admin_hold: user.admin_hold || 'No',
    editDate: user.editDate || null,
    editedBy: user.editedBy || null
  };
}

/**
 * Grant permissions for newly assigned spreadsheets
 */
function grantPermissionsForAssignments(userEmail, assignedSpreadsheets) {
  if (!assignedSpreadsheets || assignedSpreadsheets.length === 0) {
    return { granted: 0, errors: [] };
  }
  
  // Use the correct property name and helper function
  const spreadsheets = getAllSpreadsheets();
  
  if (!spreadsheets || Object.keys(spreadsheets).length === 0) {
    return { granted: 0, errors: ['No spreadsheets in system'] };
  }
  
  let granted = 0;
  const errors = [];
  
  assignedSpreadsheets.forEach(function(assignment) {
    // Check if it's a group number (1-2 digit number)
    if (/^\d{1,2}$/.test(assignment)) {
      // It's a group - grant access to all sheets in that group
      Object.keys(spreadsheets).forEach(function(sheetKey) {
        const sheet = spreadsheets[sheetKey];
        if (sheet.group === assignment) {
          try {
            const ss = SpreadsheetApp.openById(sheet.sheetID);
            // Grant appropriate permission level
            if (sheet.access === 'Editor') {
              ss.addEditor(userEmail);
            } else {
              ss.addViewer(userEmail);
            }
            granted++;
            
            // Log permission grant
            logActivity(
              EVENT_TYPES.PERMISSION_GRANTED,
              userEmail,
              'Granted ' + sheet.access + ' access to: ' + sheet.sheetTitle,
              { sheetID: sheet.sheetID, accessType: sheet.access, viaGroup: assignment }
            );
          } catch (e) {
            errors.push('Failed to grant access to ' + sheet.sheetTitle + ': ' + e.message);
            Logger.log('Permission error for ' + sheet.sheetTitle + ': ' + e.toString());
          }
        }
      });
    } else {
      // It's an individual sheet ID
      const sheetKey = 'sheet_' + assignment;
      if (spreadsheets[sheetKey]) {
        const sheet = spreadsheets[sheetKey];
        try {
          const ss = SpreadsheetApp.openById(sheet.sheetID);
          // Grant appropriate permission level
          if (sheet.access === 'Editor') {
            ss.addEditor(userEmail);
          } else {
            ss.addViewer(userEmail);
          }
          granted++;
          
          // Log permission grant
          logActivity(
            EVENT_TYPES.PERMISSION_GRANTED,
            userEmail,
            'Granted ' + sheet.access + ' access to: ' + sheet.sheetTitle,
            { sheetID: sheet.sheetID, accessType: sheet.access }
          );
        } catch (e) {
          errors.push('Failed to grant access to ' + sheet.sheetTitle + ': ' + e.message);
          Logger.log('Permission error for ' + sheet.sheetTitle + ': ' + e.toString());
        }
      } else {
        errors.push('Sheet ID not found: ' + assignment);
      }
    }
  });
  
  return { granted: granted, errors: errors };
}

/**
 * Update user details and assignments
 * @param {object} userData - Updated user object
 * @param {array} originalAssignments - Original assignments before edits
 * @param {string} sessionToken - Admin session token
 * @return {object} Success response
 */
function updateUser(userData, originalAssignments, sessionToken) {
  // Verify session
  const session = verifySessionToken(sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Invalid session. Please log in again.' };
  }
  
  // Check if requester is admin
  const requesterEmail = session.userID;
  const users = getAllUsers();
  const requesterKey = 'user_' + requesterEmail;
  const requester = users[requesterKey];
  
  if (!requester || requester.userType !== 'Admin') {
    return { success: false, message: 'Admin access required' };
  }
  
  const userEmail = normalizeEmail(userData.email || userData.userID);

  // SUPER ADMIN PROTECTION - Allow super admin to modify their own profile
  if (isSuperAdmin(userEmail) && !isSuperAdmin(requesterEmail)) {
    return { success: false, message: 'Access denied: The super admin account cannot be modified by other administrators' };
  }
  
  const userKey = 'user_' + userEmail;
  
  if (!users[userKey]) {
    return { success: false, message: 'User not found' };
  }
  
  // Determine what was removed
  const currentAssignments = userData.assigned_spreadsheets || [];
  const removedAssignments = (originalAssignments || []).filter(function(a) {
    return !currentAssignments.includes(a);
  });
  
  // Initialize results
  let revocationResult = { 
    sheetsRevoked: 0, 
    sheetsReleased: 0, 
    details: [], 
    errors: [] 
  };
  
  // Only revoke if assignments were removed
  if (removedAssignments.length > 0) {
    revocationResult = revokeActivePermissions(userEmail, removedAssignments);
  }
  
  // NOTE: We do NOT grant permissions here - permissions are only granted
  // when user actually requests access via requestSpreadsheetAccess()
  const grantResult = { granted: 0, errors: [] };
  
  // Update user object fields
  users[userKey].fullName = userData.fullName;
  users[userKey].userType = userData.userType;
  users[userKey].admin_hold = userData.admin_hold;
  users[userKey].assigned_spreadsheets = currentAssignments;
  users[userKey].registrationApproval = userData.registrationApproval || users[userKey].registrationApproval;
  
  // Track edit timestamp and editor
  users[userKey].editDate = new Date().toISOString();
  users[userKey].editedBy = session.userID;
  
  // Save changes
  saveAllUsers(users);
  
  // Build detailed response message
  const message = buildUpdateMessage(grantResult, revocationResult, removedAssignments, currentAssignments.length);

  // Log comprehensive activity
  logDetailedUserUpdate(session.userID, userEmail, grantResult, revocationResult, removedAssignments, currentAssignments.length);
  
  // Email admin if there were errors
  if (revocationResult.errors.length > 0) {
    emailAdminAboutErrors(userEmail, revocationResult.errors);
  }
  
  return { success: true, message: message };
}

/**
 * Delete user from system and revoke all permissions
 * @param {string} userEmail - Email of user to delete
 * @param {string} sessionToken - Admin session token
 * @return {object} Success response
 */
function deleteUser(userEmail, sessionToken) {
  // Normalize email for consistent lookup
  userEmail = normalizeEmail(userEmail);

  // Verify admin access
  const session = verifySessionToken(sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Invalid session' };
  }
  
  if (!session.isAdmin) {
    return { success: false, message: 'Unauthorized: Admin access required' };
  }
  
  // SUPER ADMIN PROTECTION - Nobody can delete the super admin (not even themselves)
  if (isSuperAdmin(userEmail)) {
    return { success: false, message: 'The super admin account cannot be deleted' };
  }
  
  const users = getAllUsers();
  const userKey = 'user_' + userEmail;
  
  if (!users[userKey]) {
    return { success: false, message: 'User not found' };
  }
  
  // Check if user has any checked-out spreadsheets
  const spreadsheets = getAllSpreadsheets();
  const checkedOutSheets = [];
  
  Object.keys(spreadsheets).forEach(function(key) {
    const sheet = spreadsheets[key];
    if (sheet.currentUser === userEmail) {
      checkedOutSheets.push({
        sheetID: sheet.sheetID,
        sheetTitle: sheet.sheetTitle
      });
    }
  });
  
  // Release all checked-out spreadsheets
  checkedOutSheets.forEach(function(sheet) {
    try {
      // Remove Google Sheets permissions
      const ss = SpreadsheetApp.openById(sheet.sheetID);
      ss.removeEditor(userEmail);
      ss.removeViewer(userEmail);
      
      // Clear currentUser field
      spreadsheets['sheet_' + sheet.sheetID].currentUser = '';
      spreadsheets['sheet_' + sheet.sheetID].checkedOutTime = null;
      
      // Log permission removal
      logActivity(
        EVENT_TYPES.PERMISSION_REMOVED,
        session.userID,
        'Permissions removed from "' + sheet.sheetTitle + '" for user ' + userEmail + ' (user deleted)',
        { sheetID: sheet.sheetID, sheetTitle: sheet.sheetTitle, deletedUser: userEmail }
      );
    } catch (e) {
      Logger.log('Error removing permissions for ' + sheet.sheetID + ': ' + e.message);
      // Continue with deletion even if permission removal fails
    }
  });
  
  // Save updated spreadsheets
  saveAllSpreadsheets(spreadsheets);
  
  // Delete user from system
  delete users[userKey];
  saveAllUsers(users);
  
  // Log user deletion
  logActivity(
    EVENT_TYPES.USER_DELETED,
    session.userID,
    'User deleted: ' + userEmail + (checkedOutSheets.length > 0 ? ' (had ' + checkedOutSheets.length + ' checked-out sheets)' : ''),
    { deletedUser: userEmail, sheetsReleased: checkedOutSheets.length }
  );
  
  return { 
    success: true, 
    message: 'User deleted successfully',
    sheetsReleased: checkedOutSheets.length
  };
}

// ==================================================================================
// GENERAL USER LOGIN LOCKOUT SYSTEM
// ==================================================================================

/**
 * GET LOGIN LOCKOUT STATUS
 * Returns the current lockout state and metadata
 * @param {string} sessionToken - Admin session token
 * @return {Object} {success, enabled, enabledBy, enabledAt, message}
 */
function getLoginLockoutStatus(sessionToken) {
  // Verify admin session
  const session = verifySessionToken(sessionToken);
  if (!session.valid || !session.isAdmin) {
    return { success: false, message: 'Admin access required' };
  }
  
  try {
    const props = PropertiesService.getScriptProperties();
    const lockoutJson = props.getProperty('LOGIN_LOCKOUT');
    
    if (!lockoutJson) {
      return {
        success: true,
        enabled: false,
        enabledBy: null,
        enabledAt: null
      };
    }
    
    const lockout = JSON.parse(lockoutJson);
    return {
      success: true,
      enabled: lockout.enabled || false,
      enabledBy: lockout.enabledBy || null,
      enabledAt: lockout.enabledAt || null
    };
    
  } catch (error) {
    Logger.log('Error getting lockout status: ' + error.toString());
    return { success: false, message: 'Error retrieving lockout status' };
  }
}

/**
 * SET LOGIN LOCKOUT STATUS
 * Enables or disables the general user login lockout
 * @param {boolean} enabled - True to lock, false to unlock
 * @param {string} sessionToken - Admin session token
 * @return {Object} {success, message, enabled}
 */
function setLoginLockoutStatus(enabled, sessionToken) {
  // Verify admin session
  const session = verifySessionToken(sessionToken);
  if (!session.valid || !session.isAdmin) {
    return { success: false, message: 'Admin access required' };
  }
  
  try {
    const props = PropertiesService.getScriptProperties();
    const adminEmail = session.userID;
    
    if (enabled) {
      // Enable lockout
      const lockoutData = {
        enabled: true,
        enabledBy: adminEmail,
        enabledAt: new Date().toISOString()
      };
      
      props.setProperty('LOGIN_LOCKOUT', JSON.stringify(lockoutData));
      
      // Log the action
      logActivity(
        EVENT_TYPES.LOGIN_LOCKOUT_ENABLED,
        adminEmail,
        'General user login lockout ENABLED',
        { enabledBy: adminEmail }
      );
      
      return {
        success: true,
        message: 'General user login has been LOCKED.\n\nGeneral users will not be able to log in until this is disabled.',
        enabled: true,
        enabledBy: adminEmail,
        enabledAt: lockoutData.enabledAt
      };
      
    } else {
      // Disable lockout
      props.deleteProperty('LOGIN_LOCKOUT');
      
      // Log the action
      logActivity(
        EVENT_TYPES.LOGIN_LOCKOUT_DISABLED,
        adminEmail,
        'General user login lockout DISABLED',
        { disabledBy: adminEmail }
      );
      
      return {
        success: true,
        message: 'General user login has been UNLOCKED.\n\nGeneral users can now log in normally.',
        enabled: false
      };
    }
    
  } catch (error) {
    Logger.log('Error setting lockout status: ' + error.toString());
    return { success: false, message: 'Error updating lockout status: ' + error.message };
  }
}

/**
 * CHECK IF LOGIN IS LOCKED (Internal helper)
 * Used by loginUser() to check lockout status without requiring a session
 * @return {Object} {locked, enabledBy, enabledAt}
 */
function isLoginLocked() {
  try {
    const props = PropertiesService.getScriptProperties();
    const lockoutJson = props.getProperty('LOGIN_LOCKOUT');
    
    if (!lockoutJson) {
      return { locked: false };
    }
    
    const lockout = JSON.parse(lockoutJson);
    return {
      locked: lockout.enabled || false,
      enabledBy: lockout.enabledBy || null,
      enabledAt: lockout.enabledAt || null
    };
    
  } catch (error) {
    Logger.log('Error checking lockout: ' + error.toString());
    return { locked: false }; // Fail open to prevent accidental lockout
  }
}

/**
 * CLEAR ALL ACTIVE SESSIONS (Manual Admin Function)
 * Admin function to manually clear all active sessions EXCEPT the current admin
 * Also releases all spreadsheets and removes permissions
 */
function clearAllActiveSessions(sessionToken) {
  // Verify admin session
  const session = verifySessionToken(sessionToken);
  
  if (!session.valid) {
    return { success: false, message: 'Invalid session' };
  }
  
  if (!session.isAdmin) {
    return { success: false, message: 'Admin access required' };
  }
  
  try {
    return withLock('sessions', function() {
      const props = PropertiesService.getScriptProperties();
      
      // Get current sessions
      const sessionsJson = props.getProperty('ACTIVE_SESSIONS') || '{}';
      const sessions = JSON.parse(sessionsJson);
      const sessionCount = Object.keys(sessions).length;
      
      // Count sessions to clear (exclude current admin)
      let sessionsToCleared = sessionCount - 1; // Subtract the admin's session
      if (sessionsToCleared < 0) sessionsToCleared = 0;
      
      // Clear all sessions EXCEPT the current admin's session
      const newSessions = {};
      newSessions[sessionToken] = sessions[sessionToken]; // Keep admin's session
      props.setProperty('ACTIVE_SESSIONS', JSON.stringify(newSessions));
      
      // Release ALL spreadsheets and remove permissions
      const spreadsheets = getAllSpreadsheets();
      const users = getAllUsers();
      let sheetsReleased = 0;
      let permissionsRemoved = 0;
      const errors = [];
      
      Object.keys(spreadsheets).forEach(function(sheetKey) {
        const sheet = spreadsheets[sheetKey];
        
        if (sheet.currentUser && sheet.currentUser !== '') {
          const userEmail = sheet.currentUser;
          const sheetTitle = sheet.sheetTitle;
          const sheetID = sheet.sheetID;
          
          try {
            // Remove Google Sheets permissions
            const ss = SpreadsheetApp.openById(sheetID);
            ss.removeEditor(userEmail);
            ss.removeViewer(userEmail);
            permissionsRemoved++;
            
            Logger.log('Removed permissions: ' + sheetTitle + ' from ' + userEmail);
          } catch (e) {
            errors.push('Failed to remove permissions for ' + sheetTitle + ': ' + e.message);
            Logger.log('Error removing permissions: ' + e.toString());
          }
          
          // Clear currentUser and checkedOutTime
          sheet.currentUser = '';
          sheet.checkedOutTime = null;
          sheetsReleased++;
        }
      });
      
      // Save updated spreadsheets
      saveAllSpreadsheets(spreadsheets);
      
      // Log the action
      logActivity(
        EVENT_TYPES.SYSTEM_CLEANUP,
        session.userID,
        'Admin manually cleared all active sessions (except own)',
        {
          sessionsCleared: sessionsToCleared,
          sheetsReleased: sheetsReleased,
          permissionsRemoved: permissionsRemoved,
          errors: errors.length
        }
      );
      
      let message = 'Successfully cleared ' + sessionsToCleared + ' session(s) and released ' + sheetsReleased + ' spreadsheet(s)';
      
      if (permissionsRemoved > 0) {
        message += '\n\nPermissions removed: ' + permissionsRemoved;
      }
      
      if (errors.length > 0) {
        message += '\n\nWarning: ' + errors.length + ' error(s) occurred while removing permissions. Check activity log for details.';
      }
      
      // Add note that admin session was preserved
      if (sessionCount > 0) {
        message += '\n\n(Your admin session was preserved)';
      }
      
      return {
        success: true,
        message: message,
        sessionsCleared: sessionsToCleared,
        sheetsReleased: sheetsReleased,
        permissionsRemoved: permissionsRemoved,
        errors: errors
      };
    });
  } catch (error) {
    Logger.log('Error in clearAllActiveSessions: ' + error.toString());
    return {
      success: false,
      message: 'Error clearing sessions: ' + error.message
    };
  }
}

/**
 * Revoke permissions only from sheets where user is currentUser
 * This respects the checkout/checkin model
 */
function revokeActivePermissions(userEmail, removedAssignments) {
  const spreadsheets = getAllSpreadsheets();
  let sheetsRevoked = 0;
  let sheetsReleased = 0;
  const details = [];
  const errors = [];
  
  // Helper function to process revocation
  function processRevocation(sheet, userEmail) {
    try {
      const ss = SpreadsheetApp.openById(sheet.sheetID);
      
      // Remove permissions
      ss.removeEditor(userEmail);
      ss.removeViewer(userEmail);
      
      // Force release
      sheet.currentUser = '';
      sheet.checkedOutTime = null;
      
      sheetsRevoked++;
      sheetsReleased++;
      details.push({
        sheetTitle: sheet.sheetTitle,
        sheetID: sheet.sheetID
      });
      
      Logger.log('Revoked and released: ' + sheet.sheetTitle + ' from ' + userEmail);
      
    } catch (e) {
      const error = {
        sheetTitle: sheet.sheetTitle,
        sheetID: sheet.sheetID,
        error: e.message
      };
      errors.push(error);
      Logger.log('ERROR revoking ' + sheet.sheetTitle + ': ' + e.toString());
    }
  }
  
  // Process each removed assignment
  removedAssignments.forEach(function(assignment) {
    // Check if it's a group number (1-2 digit number)
    if (/^\d{1,2}$/.test(assignment)) {
      // It's a group - process all sheets in that group where user is currentUser
      Object.keys(spreadsheets).forEach(function(sheetKey) {
        const sheet = spreadsheets[sheetKey];
        if (sheet.group === assignment && sheet.currentUser === userEmail) {
          processRevocation(sheet, userEmail);
        }
      });
    } else {
      // It's an individual sheet ID
      const sheetKey = 'sheet_' + assignment;
      const sheet = spreadsheets[sheetKey];
      if (sheet && sheet.currentUser === userEmail) {
        processRevocation(sheet, userEmail);
      }
    }
  });
  
  // Save spreadsheet changes
  if (sheetsReleased > 0) {
    saveAllSpreadsheets(spreadsheets);
  }
  
  return {
    sheetsRevoked: sheetsRevoked,
    sheetsReleased: sheetsReleased,
    details: details,
    errors: errors,
    removedAssignments: removedAssignments
  };
}

/**
 * GET USERS BACKUP (JSON)
 * Returns all users data as formatted JSON for backup purposes
 */
function getUsersBackup(sessionToken) {
  // Verify admin session
  const session = verifySessionToken(sessionToken);
  
  if (!session.valid) {
    return { success: false, message: 'Invalid session' };
  }
  
  if (!session.isAdmin) {
    return { success: false, message: 'Admin access required' };
  }
  
  try {
    // Get all users
    const users = getAllUsers();
    
    // Get user count
    const userCount = Object.keys(users).length;
    
    // Create timestamp
    const timestamp = new Date().toISOString();
    const dateString = timestamp.split('T')[0]; // YYYY-MM-DD
    
    // Create filename
    const filename = 'users_backup_' + dateString + '.txt';
    
    // Convert to pretty JSON
    const jsonData = JSON.stringify(users, null, 2);
    
    // Log the backup activity
    logActivity(
      EVENT_TYPES.SYSTEM_CLEANUP,
      session.userID,
      'Admin created users backup',
      { userCount: userCount, filename: filename }
    );
    
    return {
      success: true,
      data: jsonData,
      timestamp: timestamp,
      filename: filename,
      recordCount: userCount
    };
    
  } catch (error) {
    Logger.log('Error in getUsersBackup: ' + error.toString());
    return {
      success: false,
      message: 'Error creating backup: ' + error.message
    };
  }
}

/**
 * GET SPREADSHEETS BACKUP (JSON)
 * Returns all spreadsheets data as formatted JSON for backup purposes
 */
function getSpreadsheetsBackup(sessionToken) {
  // Verify admin session
  const session = verifySessionToken(sessionToken);
  
  if (!session.valid) {
    return { success: false, message: 'Invalid session' };
  }
  
  if (!session.isAdmin) {
    return { success: false, message: 'Admin access required' };
  }
  
  try {
    // Get all spreadsheets
    const spreadsheets = getAllSpreadsheets();
    
    // Get spreadsheet count
    const sheetCount = Object.keys(spreadsheets).length;
    
    // Create timestamp
    const timestamp = new Date().toISOString();
    const dateString = timestamp.split('T')[0]; // YYYY-MM-DD
    
    // Create filename
    const filename = 'spreadsheets_backup_' + dateString + '.txt';
    
    // Convert to pretty JSON
    const jsonData = JSON.stringify(spreadsheets, null, 2);
    
    // Log the backup activity
    logActivity(
      EVENT_TYPES.SYSTEM_CLEANUP,
      session.userID,
      'Admin created spreadsheets backup',
      { sheetCount: sheetCount, filename: filename }
    );
    
    return {
      success: true,
      data: jsonData,
      timestamp: timestamp,
      filename: filename,
      recordCount: sheetCount
    };
    
  } catch (error) {
    Logger.log('Error in getSpreadsheetsBackup: ' + error.toString());
    return {
      success: false,
      message: 'Error creating backup: ' + error.message
    };
  }
}

/**
 * Build detailed notification message for admin
 */
function buildUpdateMessage(grantResult, revocationResult, removedAssignments, assignedCount) {
  let message = 'User updated successfully\n\n';
  
  // Show assigned spreadsheets count (not "granted")
  if (assignedCount > 0) {
    message += '📋 Currently assigned: ' + assignedCount + ' spreadsheet(s)\n';
  }

  // Revoked permissions
  if (revocationResult.sheetsRevoked > 0) {
    message += '\n🔒 Revoked permissions: ' + revocationResult.sheetsRevoked + ' spreadsheet(s)\n';
    revocationResult.details.forEach(function(detail) {
      message += '  • ' + detail.sheetTitle + '\n';
    });
    
    if (revocationResult.sheetsReleased > 0) {
      message += '\n📤 Force-released: ' + revocationResult.sheetsReleased + ' spreadsheet(s)\n';
    }
  }
  
  // Removed assignments (even if no active checkouts)
  if (removedAssignments.length > 0 && revocationResult.sheetsRevoked === 0) {
    message += '\n📋 Removed assignments: ' + removedAssignments.length + '\n';
    message += '(No active checkouts - no permissions to revoke)\n';
  }
  
  // Errors
  if (revocationResult.errors.length > 0) {
    message += '\n⚠️ Errors during revocation: ' + revocationResult.errors.length + '\n';
    revocationResult.errors.forEach(function(err) {
      message += '  • ' + err.sheetTitle + ': ' + err.error + '\n';
    });
    message += '\nAdmin has been notified via email.\n';
  }
  
  // Grant errors
  if (grantResult.errors && grantResult.errors.length > 0) {
    message += '\n⚠️ Errors granting access: ' + grantResult.errors.length + '\n';
    grantResult.errors.forEach(function(err) {
      message += '  • ' + err + '\n';
    });
  }
  
  return message;
}

/**
 * Email admin about permission revocation errors
 */
function emailAdminAboutErrors(userEmail, errors) {
  if (!errors || errors.length === 0) return;
  
  try {
    const subject = SYSTEM_NAME + ' - Permission Revocation Errors';
    
    let errorDetails = '';
    errors.forEach(function(err) {
      errorDetails += '• ' + err.sheetTitle + ' (ID: ' + err.sheetID + ')\n  Error: ' + err.error + '\n\n';
    });
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #F44336, #EF5350); padding: 20px; border-radius: 10px 10px 0 0;">
          <h2 style="color: white; margin: 0;">⚠️ Permission Revocation Errors</h2>
        </div>
        
        <div style="background-color: #f5f5f5; padding: 25px; border-radius: 0 0 10px 10px;">
          <h3 style="color: #F44336; margin-top: 0;">Error Summary</h3>
          <p><strong>User:</strong> ${userEmail}</p>
          <p><strong>Timestamp:</strong> ${new Date().toLocaleString()}</p>
          <p><strong>Errors:</strong> ${errors.length}</p>
          
          <h3 style="color: #F44336; margin-top: 25px;">Details</h3>
          <div style="background-color: white; padding: 15px; border-left: 4px solid #F44336; border-radius: 5px;">
            <pre style="white-space: pre-wrap; margin: 0; font-family: monospace;">${errorDetails}</pre>
          </div>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 25px 0;">
          
          <p style="color: #666;">
            This occurred during an admin update to remove spreadsheet assignments.
            The user record was still updated, but some permissions could not be revoked.
          </p>
          
          <p style="color: #666;">
            <strong>Action Required:</strong> You may need to manually check these spreadsheet permissions.
          </p>
          
          <p style="font-size: 0.85em; color: #999; margin-top: 20px;">
            This notification was sent from ${SYSTEM_NAME}
          </p>
        </div>
      </div>
    `;
    
    const plainBody = `
Permission Revocation Errors

User: ${userEmail}
Timestamp: ${new Date().toLocaleString()}
Errors: ${errors.length}

Details:
${errorDetails}

This occurred during an admin update to remove spreadsheet assignments.
The user record was still updated, but some permissions could not be revoked.

You may need to manually check these spreadsheet permissions.
    `;
    
    MailApp.sendEmail({
      to: ADMIN_CONTACT_EMAIL,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });
    
    Logger.log('Error notification email sent to admin');
    
  } catch (e) {
    Logger.log('Failed to send error email: ' + e.toString());
  }
}

/**
 * Log detailed user update activity
 */
function logDetailedUserUpdate(adminEmail, userEmail, grantResult, revocationResult, removedAssignments, assignedCount) {
  let details = 'Admin updated user: ' + userEmail;
  
  const parts = [];
  
  if (assignedCount > 0) {
    parts.push('Assigned: ' + assignedCount + ' sheet(s)');
  }

  if (revocationResult.sheetsRevoked > 0) {
    parts.push('Revoked: ' + revocationResult.sheetsRevoked + ' sheet(s)');
    parts.push('Released: ' + revocationResult.sheetsReleased + ' sheet(s)');
    const sheetNames = revocationResult.details.map(function(d) { return d.sheetTitle; }).join(', ');
    parts.push('Sheets: ' + sheetNames);
  }
  
  // FIX: Convert removed assignment IDs/groups to readable names
  if (removedAssignments.length > 0) {
    const spreadsheets = getAllSpreadsheets();
    const readableNames = removedAssignments.map(function(assignment) {
      // Check if it's a group number (1-2 digit number)
      if (/^\d{1,2}$/.test(assignment)) {
        return 'Group ' + assignment;
      } else {
        // It's a sheet ID - look up the title
        const sheetKey = 'sheet_' + assignment;
        const sheet = spreadsheets[sheetKey];
        if (sheet && sheet.sheetTitle) {
          return sheet.sheetTitle;
        }
        // Fallback to shortened ID if sheet not found
        return assignment.substring(0, 10) + '...';
      }
    });
    parts.push('Removed assignments: ' + readableNames.join(', '));
  }
  
  if (parts.length > 0) {
    details += ' | ' + parts.join(' | ');
  }
  
  logActivity(
    EVENT_TYPES.USER_MODIFIED,
    adminEmail,
    details,
    {
      targetUser: userEmail,
      granted: grantResult.granted,
      revoked: revocationResult.sheetsRevoked,
      released: revocationResult.sheetsReleased,
      errors: revocationResult.errors.length,
      removedAssignments: removedAssignments
    }
  );
}

/**
 * ADMIN RESET USER PASSWORD
 * Allows administrators to reset any user's password
 * Optionally sends email notification to user
 */
function adminResetUserPassword(userEmail, newPassword, sendNotification, sessionToken) {
// Sanitize and normalize inputs
  userEmail = normalizeEmail(sanitizeInput(userEmail, 100));
  newPassword = sanitizeInput(newPassword, 100);
  
  // Verify admin session
  const session = verifySessionToken(sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Invalid session. Please log in again.' };
  }
  
  // Check if user is admin
  const adminEmail = session.userID;
  const users = getAllUsers();
  const adminKey = 'user_' + adminEmail;
  const admin = users[adminKey];
  
  if (!admin || admin.userType !== 'Admin') {
    return { success: false, message: 'Admin access required' };
  }
  
  // SUPER ADMIN PROTECTION - Allow super admin to reset their own password only
  if (isSuperAdmin(userEmail) && !isSuperAdmin(adminEmail)) {
    return { success: false, message: 'Access denied: The super admin password cannot be reset by other administrators' };
  }
  
  // Validate inputs
  if (!userEmail || !newPassword) {
    return { success: false, message: 'User email and new password are required' };
  }
  
  // Validate new password length
  if (newPassword.length < 8) {
    return { success: false, message: 'Password must be at least 8 characters long' };
  }
  
  // Find the target user
  const userKey = 'user_' + userEmail;
  const user = users[userKey];
  
  if (!user) {
    return { success: false, message: 'User not found: ' + userEmail };
  }
  
  try {
    // Hash and save new password
    users[userKey].passwordHash = hashPassword(newPassword);
    users[userKey].passwordChangedDate = new Date().toISOString();
    users[userKey].passwordChangedBy = adminEmail;
    
    // Save changes
    saveAllUsers(users);
    
    // Log the password reset
    logActivity(
      EVENT_TYPES.USER_MODIFIED,
      adminEmail,
      'Admin reset password for user: ' + userEmail,
      { targetUser: userEmail, method: 'adminResetUserPassword' }
    );
    
    // Send email notification if requested
    if (sendNotification) {
      try {
        const subject = 'Your Password Has Been Reset - ' + SYSTEM_NAME;
        const body = 'Hello ' + (user.fullName || userEmail) + ',\n\n' +
          'Your password for the ' + SYSTEM_NAME + ' has been reset by an administrator.\n\n' +
          'Your new temporary password is: ' + newPassword + '\n\n' +
          'Please log in and change your password as soon as possible using the "Reset Password" option in the Users tab.\n\n' +
          'If you did not request this password reset, please contact your administrator immediately.\n\n' +
          'Best regards,\n' +
          SYSTEM_NAME;

        MailApp.sendEmail({
          to: userEmail,
          subject: subject,
          body: body
        });
        
        logActivity(
          EVENT_TYPES.USER_MODIFIED,
          adminEmail,
          'Password reset notification email sent to: ' + userEmail
        );
        
        return { 
          success: true, 
          message: 'Password reset successfully!\n\nNotification email sent to ' + userEmail,
          emailSent: true
        };
      } catch (emailError) {
        Logger.log('Error sending password reset email: ' + emailError.toString());
        return { 
          success: true, 
          message: 'Password reset successfully!\n\n⚠️ Warning: Could not send notification email. Please inform the user manually.\n\nError: ' + emailError.message,
          emailSent: false
        };
      }
    }
    
    return { 
      success: true, 
      message: 'Password reset successfully!\n\nRemember to inform the user of their new temporary password.',
      emailSent: false
    };
    
  } catch (error) {
    Logger.log('Error in adminResetUserPassword: ' + error.toString());
    return { success: false, message: 'Error resetting password: ' + error.message };
  }
}


/**
 * FUNCTIONS BELOW
 * createSuperAdmin
 * resetSuperAdminPassword (backup method in case Property Service problem )
 * createMidnightTrigger
 * testEmailPermissions
 * testActivityLog
 * 
 */

// ==================================================================================
// SUPER ADMIN RECOVERY/MAINTENANCE FUNCTIONS
// Run these from the Apps Script editor when needed
// ==================================================================================
/**
 * CREATE OR RESTORE SUPER ADMIN ACCOUNT
 * Run this from script editor if the super admin account doesn't exist or needs to be recreated
 * 
 * INSTRUCTIONS:
 * 1. Change 'YOUR_PASSWORD_HERE' below to your desired password
 * 2. Run this function from the script editor
 * 3. IMPORTANT: Change the password back to 'YOUR_PASSWORD_HERE' after running
 */
function createSuperAdmin() {
  // ⚠️ CHANGE THIS TO YOUR PASSWORD, THEN RUN, THEN CHANGE BACK
  const password = 'default1234';
  
  // Safety check - don't run with placeholder password
  if (password === 'YOUR_PASSWORD_HERE') {
    Logger.log('❌ Please change the password in the function before running!');
    Logger.log('Edit the password variable, run the function, then change it back.');
    return {
      success: false,
      message: 'Please set a real password in the function first'
    };
  }
  
  // Validate password length
  if (password.length < 8) {
    Logger.log('❌ Password must be at least 8 characters long');
    return {
      success: false,
      message: 'Password must be at least 8 characters'
    };
  }
  
  try {
    const users = getAllUsers();
    const userKey = 'user_' + SUPER_ADMIN_EMAIL;
    
    // Check if already exists
    const exists = users[userKey] ? true : false;
    
    // Create or update super admin account
    users[userKey] = {
      userID: SUPER_ADMIN_EMAIL,
      fullName: 'Super Administrator',
      passwordHash: hashPassword(password),
      registrationApproval: 'approved',
      userType: 'Admin',
      registeredDate: users[userKey]?.registeredDate || new Date().toISOString(),
      assigned_spreadsheets: users[userKey]?.assigned_spreadsheets || [],
      admin_hold: 'No'
    };
    
    // Save changes
    saveAllUsers(users);
    
    // Log the action
    logActivity(
      exists ? EVENT_TYPES.USER_MODIFIED : EVENT_TYPES.USER_REGISTERED,
      'System',
      (exists ? 'Super admin account restored' : 'Super admin account created') + ': ' + SUPER_ADMIN_EMAIL,
      { method: 'createSuperAdmin()' }
    );
    
    Logger.log('✅ Super admin account ' + (exists ? 'restored' : 'created') + ' successfully!');
    Logger.log('Email: ' + SUPER_ADMIN_EMAIL);
    Logger.log('Status: Approved');
    Logger.log('User Type: Admin');
    Logger.log('');
    Logger.log('⚠️ IMPORTANT: Now change the password in this function back to');
    Logger.log('   "YOUR_PASSWORD_HERE" so your real password is not saved in the code!');
    
    return {
      success: true,
      message: 'Super admin account ' + (exists ? 'restored' : 'created') + ' successfully'
    };
    
  } catch (error) {
    Logger.log('❌ Error creating super admin: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/**
 * RESET SUPER ADMIN PASSWORD
 * Run this from the script editor to change the super admin password
 * 
 * INSTRUCTIONS:
 * 1. Change 'YOUR_NEW_PASSWORD_HERE' below to your desired password
 * 2. Run this function from the script editor
 * 3. IMPORTANT: Change the password back to 'YOUR_NEW_PASSWORD_HERE' after running
 *    (so your real password isn't saved in the code)
 */
function resetSuperAdminPassword() {
  // ⚠️ CHANGE THIS TO YOUR NEW PASSWORD, THEN RUN, THEN CHANGE BACK
  const newPassword = 'default1234';
  
  // Safety check - don't run with placeholder password
  if (newPassword === 'YOUR-YOUR_NEW_PASSWORD_HERE') {
    Logger.log('❌ Please change the password in the function before running!');
    Logger.log('Edit the newPassword variable, run the function, then change it back.');
    return {
      success: false,
      message: 'Please set a real password in the function first'
    };
  }
  
  // Validate password length
  if (newPassword.length < 8) {
    Logger.log('❌ Password must be at least 8 characters long');
    return {
      success: false,
      message: 'Password must be at least 8 characters'
    };
  }
  
  try {
    const users = getAllUsers();
    const userKey = 'user_' + SUPER_ADMIN_EMAIL;
    
    // Check if super admin exists
    if (!users[userKey]) {
      Logger.log('❌ Super admin account not found: ' + SUPER_ADMIN_EMAIL);
      Logger.log('You may need to register this account first through the web app.');
      return {
        success: false,
        message: 'Super admin account not found. Please register first.'
      };
    }
    
    // Hash and save new password
    users[userKey].passwordHash = hashPassword(newPassword);
    
    // Save changes
    saveAllUsers(users);
    
    // Log the action (don't log the actual password!)
    logActivity(
      EVENT_TYPES.USER_MODIFIED,
      'System',
      'Super admin password reset via script: ' + SUPER_ADMIN_EMAIL,
      { method: 'resetSuperAdminPassword()' }
    );
    
    Logger.log('✅ Super admin password reset successfully!');
    Logger.log('Email: ' + SUPER_ADMIN_EMAIL);
    Logger.log('');
    Logger.log('⚠️ IMPORTANT: Now change the password in this function back to');
    Logger.log('   "YOUR_NEW_PASSWORD_HERE" so your real password is not saved in the code!');
    
    return {
      success: true,
      message: 'Password reset successfully'
    };
    
  } catch (error) {
    Logger.log('❌ Error resetting password: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

// ==================================================================================
// TEST / SETUP / HELPER - FUNCTIONS
// ==================================================================================
/**
 * CREATE MIDNIGHT TRIGGER
 * Run this function ONCE to set up the automated midnight cleanup
 */
function createMidnightTrigger() {
  // Delete any existing midnight cleanup triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'midnightCleanup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger for midnight (between midnight and 1 AM)
  ScriptApp.newTrigger('midnightCleanup')
    .timeBased()
    .atHour(0) // Midnight
    .everyDays(1) // Every day
    .create();
  
  Logger.log('✓ Midnight cleanup trigger created successfully!');
  Logger.log('The midnightCleanup() function will run daily at midnight.');
  
  return {
    success: true,
    message: 'Midnight trigger created - will run daily at midnight'
  };
}

/**
 * TEST EMAIL PERMISSIONS
 * Run this function once to authorize email sending permissions
 * This will trigger the authorization dialog
 */
function testEmailPermissions() {
  try {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: 'Google Sheets Access Manager - Email Test',
      body: 'Email permissions are working! You can now receive contact messages from users.'
    });
    Logger.log('✓ Email sent successfully! Permissions are configured.');
    Logger.log('Check your inbox for the test email.');
    return 'Success! Check your email.';
  } catch (e) {
    Logger.log('✗ Error: ' + e.toString());
    Logger.log('Please authorize the script to send emails.');
    return 'Error: ' + e.toString();
  }
}

/**
 * TEST ACTIVITY LOG
 * Run this to verify activity log is working
 */
function testActivityLog() {
  // Add test entry
  logActivity(
    EVENT_TYPES.SYSTEM_CLEANUP,
    'test@gmail.com',
    'Test activity log entry',
    { test: true }
  );
  
  // Retrieve logs
  const logs = getActivityLog(10);
  
  Logger.log('Activity log test results:');
  Logger.log('Total entries: ' + logs.length);
  Logger.log('Latest entry: ' + JSON.stringify(logs[0], null, 2));
  
  return {
    success: true,
    totalEntries: logs.length,
    latestEntry: logs[0]
  };
}


