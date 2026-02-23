// ============================================
// IDMS Mail Archivator - Frontend
// ============================================

// Configuration
const CONFIG = {
    // Backend API URL (change this when you deploy backend)
    BACKEND_URL: 'http://localhost:5000',  // For now, localhost
    
    // Customer domains (hardcoded for now, will come from backend later)
    CUSTOMER_DOMAINS: [
        'acme.com',
        'globex.com',
        'initech.com'
    ]
};

// ============================================
// Office.js Initialization
// ============================================

Office.onReady(function(info) {
    console.log('Office.js ready');
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);
    
    // Update status
    updateStatus('Active', true);
    
    // Load current email info
    loadCurrentEmail();
    
    // Load statistics (placeholder for now)
    loadStatistics();
    
    // Set up periodic heartbeat (every 5 minutes)
    setInterval(sendHeartbeat, 5 * 60 * 1000);
    
    // Send initial heartbeat
    sendHeartbeat();
});

// ============================================
// Main Functions
// ============================================

/**
 * Load current email information
 */
function loadCurrentEmail() {
    var item = Office.context.mailbox.item;
    
    if (!item) {
        document.getElementById('emailSubject').textContent = 'No email selected';
        return;
    }
    
    // Get subject
    var subject = item.subject || '(No subject)';
    document.getElementById('emailSubject').textContent = subject;
    
    // Get sender (async)
    item.from.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var from = result.value;
            var email = from.emailAddress;
            
            // Display sender
            document.getElementById('emailFrom').textContent = 
                from.displayName + ' <' + email + '>';
            
            // Extract domain
            var domain = extractDomain(email);
            document.getElementById('emailDomain').textContent = '@' + domain;
        }
    });
}

/**
 * Classify current email
 */
function classifyCurrentEmail() {
    var item = Office.context.mailbox.item;
    
    if (!item) {
        alert('No email selected');
        return;
    }
    
    var classifyBtn = document.getElementById('classifyBtn');
    classifyBtn.disabled = true;
    classifyBtn.textContent = 'Classifying...';
    
    // Get sender
    item.from.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var email = result.value.emailAddress;
            var domain = extractDomain(email);
            
            // For now, classify locally (later will call backend)
            var classification = classifyEmailLocally(domain);
            
            // Display result
            displayClassification(classification);
            
            // Re-enable button
            classifyBtn.disabled = false;
            classifyBtn.textContent = 'Classify Email';
            
            // Update last update time
            updateLastUpdate();
        } else {
            classifyBtn.disabled = false;
            classifyBtn.textContent = 'Classify Email';
            alert('Error getting email sender');
        }
    });
}

/**
 * Classify email locally (temporary - will use backend API later)
 */
function classifyEmailLocally(domain) {
    // Check if domain is in customer list
    if (CONFIG.CUSTOMER_DOMAINS.indexOf(domain) !== -1) {
        return {
            result: 'Archive',
            folder: domain.split('.')[0].toUpperCase() + '_12345',
            reason: 'Customer domain found'
        };
    } else {
        return {
            result: 'Ignore',
            reason: 'Not a customer domain'
        };
    }
}

/**
 * Display classification result
 */
function displayClassification(classification) {
    var classElement = document.getElementById('classification');
    
    // Remove all classification classes
    classElement.className = 'value';
    
    if (classification.result === 'Archive') {
        classElement.className = 'value classification-archive';
        classElement.textContent = '✓ Archive to ' + classification.folder;
    } else if (classification.result === 'Pending') {
        classElement.className = 'value classification-pending';
        classElement.textContent = '⏳ Pending - ' + classification.reason;
    } else {
        classElement.className = 'value classification-ignore';
        classElement.textContent = '✗ Ignore - ' + classification.reason;
    }
}

/**
 * Load statistics (placeholder - will call backend API later)
 */
function loadStatistics() {
    // For now, show placeholder values
    document.getElementById('statArchived').textContent = '-';
    document.getElementById('statPending').textContent = '-';
    document.getElementById('statIgnored').textContent = '-';
}

/**
 * Refresh statistics
 */
function refreshStats() {
    loadStatistics();
    loadCurrentEmail();
    updateLastUpdate();
}

/**
 * Send heartbeat to backend (placeholder for now)
 */
function sendHeartbeat() {
    console.log('Heartbeat sent');
    // Will implement backend call later
    
    /*
    var userEmail = Office.context.mailbox.userProfile.emailAddress;
    
    fetch(CONFIG.BACKEND_URL + '/api/heartbeat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            userEmail: userEmail,
            timestamp: new Date().toISOString()
        })
    }).catch(function(error) {
        console.error('Heartbeat failed:', error);
    });
    */
}

// ============================================
// Helper Functions
// ============================================

/**
 * Extract domain from email address
 */
function extractDomain(email) {
    var parts = email.split('@');
    return parts.length > 1 ? parts[1] : '';
}

/**
 * Update status indicator
 */
function updateStatus(text, active) {
    var statusText = document.getElementById('statusText');
    var statusDot = document.getElementById('statusDot');
    
    statusText.textContent = text;
    
    if (active) {
        statusDot.className = 'status-dot status-active';
    } else {
        statusDot.className = 'status-dot';
    }
}

/**
 * Update last update timestamp
 */
function updateLastUpdate() {
    var now = new Date();
    var timeStr = now.toLocaleTimeString();
    document.getElementById('lastUpdate').textContent = 'Last update: ' + timeStr;
}
