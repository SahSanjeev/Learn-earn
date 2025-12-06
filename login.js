// Login and User Management System
let currentOTP = '';
let currentMobileNumber = '';
let currentUserName = '';

// Initialize login system
document.addEventListener('DOMContentLoaded', function() {
    checkLoginStatus();
    setupLoginModal();
    setupForms();
    setupAdminExport();
});

// Check if user is logged in
function checkLoginStatus() {
    const userData = localStorage.getItem('currentUser');
    if (userData) {
        const user = JSON.parse(userData);
        showLoggedInState(user);
    }
}

// Show logged in state
function showLoggedInState(user) {
    document.getElementById('loginBtn').style.display = 'none';
    document.getElementById('logoutBtn').style.display = 'inline-block';
    document.getElementById('adminBtn').style.display = 'inline-block';
}

// Setup login modal
function setupLoginModal() {
    const modal = document.getElementById('loginModal');
    const loginBtn = document.getElementById('loginBtn');
    const logoutBtn = document.getElementById('logoutBtn');
    const closeBtn = document.querySelector('.modal-close');

    // Open modal
    loginBtn.addEventListener('click', function() {
        modal.style.display = 'block';
        document.getElementById('loginStep1').style.display = 'block';
        document.getElementById('loginStep2').style.display = 'none';
        document.getElementById('mobileForm').reset();
        document.getElementById('otpForm').reset();
    });

    // Close modal
    closeBtn.addEventListener('click', function() {
        modal.style.display = 'none';
    });

    // Close modal when clicking outside
    window.addEventListener('click', function(event) {
        if (event.target === modal) {
            modal.style.display = 'none';
        }
    });

    // Logout
    logoutBtn.addEventListener('click', function() {
        localStorage.removeItem('currentUser');
        document.getElementById('loginBtn').style.display = 'inline-block';
        document.getElementById('logoutBtn').style.display = 'none';
        document.getElementById('adminBtn').style.display = 'none';
        alert('Logged out successfully!');
    });
}

// Setup forms
function setupForms() {
    const mobileForm = document.getElementById('mobileForm');
    const otpForm = document.getElementById('otpForm');
    const resendOTP = document.getElementById('resendOTP');

    // Mobile number form submission
    mobileForm.addEventListener('submit', function(e) {
        e.preventDefault();
        const mobileNumber = document.getElementById('mobileNumber').value;
        const userName = document.getElementById('userName').value || 'Guest User';

        if (!validateMobileNumber(mobileNumber)) {
            showError('Please enter a valid mobile number (10-15 digits)');
            return;
        }

        currentMobileNumber = mobileNumber;
        currentUserName = userName;

        // Generate OTP (6 digits)
        currentOTP = generateOTP();
        
        // Display OTP (in production, this would be sent via SMS)
        document.getElementById('displayOTP').textContent = currentOTP;
        
        // Show OTP step
        document.getElementById('loginStep1').style.display = 'none';
        document.getElementById('loginStep2').style.display = 'block';
        
        // Clear any previous errors
        clearMessages();
    });

    // OTP form submission
    otpForm.addEventListener('submit', function(e) {
        e.preventDefault();
        const enteredOTP = document.getElementById('otpCode').value;

        if (enteredOTP === currentOTP) {
            // Login successful
            handleSuccessfulLogin();
        } else {
            showError('Invalid OTP. Please try again.');
        }
    });

    // Resend OTP
    resendOTP.addEventListener('click', function() {
        currentOTP = generateOTP();
        document.getElementById('displayOTP').textContent = currentOTP;
        document.getElementById('otpCode').value = '';
        showSuccess('New OTP has been generated!');
    });
}

// Generate 6-digit OTP
function generateOTP() {
    return Math.floor(100000 + Math.random() * 900000).toString();
}

// Validate mobile number
function validateMobileNumber(mobile) {
    const cleaned = mobile.replace(/[^0-9+]/g, '');
    return cleaned.length >= 10 && cleaned.length <= 15;
}

// Handle successful login
function handleSuccessfulLogin() {
    const loginData = {
        mobileNumber: currentMobileNumber,
        userName: currentUserName,
        loginTime: new Date().toISOString(),
        loginDate: new Date().toLocaleDateString(),
        loginTimestamp: Date.now()
    };

    // Store current user
    localStorage.setItem('currentUser', JSON.stringify(loginData));

    // Save to user database (localStorage)
    saveUserToDatabase(loginData);

    // Show success and close modal
    showSuccess('Login successful!');
    
    setTimeout(() => {
        document.getElementById('loginModal').style.display = 'none';
        showLoggedInState(loginData);
        clearMessages();
    }, 1000);
}

// Save user to database (localStorage)
function saveUserToDatabase(userData) {
    let users = JSON.parse(localStorage.getItem('userDatabase') || '[]');
    
    // Check if user already exists
    const existingUserIndex = users.findIndex(u => u.mobileNumber === userData.mobileNumber);
    
    if (existingUserIndex !== -1) {
        // Update existing user
        users[existingUserIndex].lastLoginTime = userData.loginTime;
        users[existingUserIndex].lastLoginDate = userData.loginDate;
        users[existingUserIndex].loginCount = (users[existingUserIndex].loginCount || 0) + 1;
        users[existingUserIndex].userName = userData.userName;
    } else {
        // Add new user
        const newUser = {
            id: Date.now().toString(),
            mobileNumber: userData.mobileNumber,
            userName: userData.userName,
            firstLoginTime: userData.loginTime,
            firstLoginDate: userData.loginDate,
            lastLoginTime: userData.loginTime,
            lastLoginDate: userData.loginDate,
            loginCount: 1,
            totalEarnings: 0,
            coursesCompleted: 0,
            status: 'Active'
        };
        users.push(newUser);
    }
    
    localStorage.setItem('userDatabase', JSON.stringify(users));
}

// Setup admin export
function setupAdminExport() {
    const adminBtn = document.getElementById('adminBtn');
    
    adminBtn.addEventListener('click', function() {
        exportToExcel();
    });
}

// Export data to Excel
function exportToExcel() {
    const users = JSON.parse(localStorage.getItem('userDatabase') || '[]');
    
    if (users.length === 0) {
        alert('No user data available to export.');
        return;
    }

    // Prepare data for Excel
    const excelData = users.map(user => ({
        'ID': user.id,
        'Mobile Number': user.mobileNumber,
        'User Name': user.userName,
        'First Login Date': user.firstLoginDate,
        'First Login Time': user.firstLoginTime,
        'Last Login Date': user.lastLoginDate,
        'Last Login Time': user.lastLoginTime,
        'Login Count': user.loginCount,
        'Total Earnings': user.totalEarnings || 0,
        'Courses Completed': user.coursesCompleted || 0,
        'Status': user.status
    }));

    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData);

    // Set column widths
    const colWidths = [
        { wch: 15 }, // ID
        { wch: 18 }, // Mobile Number
        { wch: 20 }, // User Name
        { wch: 18 }, // First Login Date
        { wch: 20 }, // First Login Time
        { wch: 18 }, // Last Login Date
        { wch: 20 }, // Last Login Time
        { wch: 12 }, // Login Count
        { wch: 15 }, // Total Earnings
        { wch: 18 }, // Courses Completed
        { wch: 12 }  // Status
    ];
    ws['!cols'] = colWidths;

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'User Data');

    // Generate summary sheet
    const summaryData = [{
        'Total Users': users.length,
        'Active Users': users.filter(u => u.status === 'Active').length,
        'Total Logins': users.reduce((sum, u) => sum + (u.loginCount || 0), 0),
        'Total Earnings': users.reduce((sum, u) => sum + (u.totalEarnings || 0), 0),
        'Total Courses Completed': users.reduce((sum, u) => sum + (u.coursesCompleted || 0), 0),
        'Export Date': new Date().toLocaleDateString(),
        'Export Time': new Date().toLocaleTimeString()
    }];

    const summaryWs = XLSX.utils.json_to_sheet(summaryData);
    summaryWs['!cols'] = [{ wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    const filename = `Learn_Earn_Users_${timestamp}.xlsx`;

    // Download file
    XLSX.writeFile(wb, filename);
    
    alert(`Data exported successfully! File: ${filename}`);
}

// Show success message
function showSuccess(message) {
    clearMessages();
    const successDiv = document.createElement('div');
    successDiv.className = 'success-message';
    successDiv.textContent = message;
    const form = document.querySelector('#loginStep2 form') || document.querySelector('#loginStep1 form');
    form.insertBefore(successDiv, form.firstChild);
}

// Show error message
function showError(message) {
    clearMessages();
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error-message';
    errorDiv.textContent = message;
    const form = document.querySelector('#loginStep2 form') || document.querySelector('#loginStep1 form');
    form.insertBefore(errorDiv, form.firstChild);
}

// Clear messages
function clearMessages() {
    const messages = document.querySelectorAll('.success-message, .error-message');
    messages.forEach(msg => msg.remove());
}

