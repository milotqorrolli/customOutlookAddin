// ============================================
// IDMS Mail Archivator - Frontend JavaScript
// ============================================

// Office.js Initialization
Office.onReady(function(info) {
    console.log('IDMS Mail Archivator loaded');
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);
    
    // Initialize user email
    initializeUserInfo();
});

// ============================================
// Tab Switching
// ============================================

function switchTab(tabName) {
    // Hide all tab panes
    var tabPanes = document.getElementsByClassName('tab-pane');
    for (var i = 0; i < tabPanes.length; i++) {
        tabPanes[i].classList.remove('active');
    }
    
    // Remove active class from all tabs
    var tabs = document.getElementsByClassName('tab');
    for (var i = 0; i < tabs.length; i++) {
        tabs[i].classList.remove('active');
    }
    
    // Show selected tab pane
    document.getElementById(tabName).classList.add('active');
    
    // Add active class to clicked tab
    event.target.classList.add('active');
}

// ============================================
// Initialize User Info
// ============================================

function initializeUserInfo() {
    try {
        var userEmail = Office.context.mailbox.userProfile.emailAddress;
        document.getElementById('myMailAddress').value = userEmail;
    } catch (error) {
        console.error('Error getting user email:', error);
    }
}

// ============================================
// Archive Mails Function
// ============================================

function archiveMails() {
    var item = Office.context.mailbox.item;
    
    if (!item) {
        alert('No email selected');
        return;
    }
    
    // Show processing message
    var archiveBtn = document.querySelector('.archive-btn');
    archiveBtn.textContent = 'Processing...';
    archiveBtn.disabled = true;
    
    // Get sender
    item.from.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var email = result.value.emailAddress;
            var domain = extractDomain(email);
            
            // For now, just show success message
            // Later this will call backend API
            setTimeout(function() {
                alert('Email from ' + email + ' will be archived');
                archiveBtn.textContent = 'Archive Mails';
                archiveBtn.disabled = false;
            }, 1000);
        } else {
            archiveBtn.textContent = 'Archive Mails';
            archiveBtn.disabled = false;
            alert('Error getting email sender');
        }
    });
}

// ============================================
// Helper Functions
// ============================================

function extractDomain(email) {
    var parts = email.split('@');
    return parts.length > 1 ? parts[1] : '';
}

function clearField(fieldId) {
    document.getElementById(fieldId).value = '';
}

// ============================================
// Title Bar Button Handlers
// ============================================

document.addEventListener('DOMContentLoaded', function() {
    // These are just visual - Office.js controls actual window behavior
    var minimizeBtn = document.querySelector('.minimize');
    var maximizeBtn = document.querySelector('.maximize');
    var closeBtn = document.querySelector('.close');
    
    if (minimizeBtn) {
        minimizeBtn.addEventListener('click', function() {
            console.log('Minimize clicked');
        });
    }
    
    if (maximizeBtn) {
        maximizeBtn.addEventListener('click', function() {
            console.log('Maximize clicked');
        });
    }
    
    if (closeBtn) {
        closeBtn.addEventListener('click', function() {
            console.log('Close clicked');
            // In actual add-in, this would close the task pane
        });
    }
});
