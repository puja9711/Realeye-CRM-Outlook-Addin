/**
 * taskpane.js
 * Comprehensive logic for RealEye AI CRM Outlook Add-in
 * Implements Routing, Login, REST Contacts fetching, and AI details parsing.
 */

const BASE_URL = 'https://api.realeye.live';
let currentRestToken = null;
let currentRestUrl = null;

// --- Pagination state for inbox list ---
let inboxPageSkip = 0;
const INBOX_PAGE_SIZE = 20;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initApp();
        setupEventListeners();
    }
});

// --- Initialization & UI Routing ---

function initApp() {
    const token = Office.context.roamingSettings.get("CRM_TOKEN");
    const entId = Office.context.roamingSettings.get("ENTERPRISE_ID");

    if (!token || !entId) {
        showView('login');
    } else {
        // Since we are always in an email context (due to MessageReadCommandSurface)
        // Show the current email details FIRST
        showView('detail');
        loadCurrentContactData();
        
        // Also update the footer email
        const userEmail = Office.context.roamingSettings.get("CRM_USER_EMAIL") || 'CRM User';
        const userName = Office.context.roamingSettings.get("CRM_USER_NAME") || userEmail.split('@')[0];
        document.getElementById('footer-email').innerText = userEmail;
        document.getElementById('footer-username').innerText = userName;
    }
}

function showView(viewId) {
    document.getElementById('login-view').classList.remove('active');
    document.getElementById('list-view').classList.remove('active');
    document.getElementById('detail-view').classList.remove('active');

    document.getElementById(viewId + '-view').classList.add('active');

    // Header & Footer visibility
    if (viewId === 'login') {
        document.getElementById('app-header').style.display = 'none';
        document.getElementById('app-footer').style.display = 'none';
    } else {
        document.getElementById('app-header').style.display = 'flex';
        document.getElementById('app-footer').style.display = 'block';
    }
}

function setupEventListeners() {
    // Auth
    document.getElementById('login-btn').onclick = handleLogin;
    document.getElementById('menu-logout').onclick = handleLogout;
    document.getElementById('footer-logout-btn').onclick = handleLogout;
    
    // Global Header Dropdown
    document.getElementById('menu-toggle').onclick = (e) => {
        document.getElementById('dropdown-menu').classList.toggle('hidden');
        e.stopPropagation();
    };
    document.addEventListener('click', () => {
        document.getElementById('dropdown-menu').classList.add('hidden');
    });

    // Navigation
    document.getElementById('back-to-list-btn').onclick = () => {
        showView('list');
        loadRecentContacts();
    };
    
    // Details Togglers
    document.getElementById('toggle-details-btn').onclick = toggleDetails;
    document.getElementById('create-contact-btn').onclick = handleCreateContact;
    document.getElementById('open-crm-btn').onclick = () => window.open('https://realeye.live/crm/all-contacts', '_blank');

    // List view events
    document.getElementById('refresh-list-btn').onclick = () => {
        inboxPageSkip = 0;
        loadRecentContacts(true);
    };
    document.getElementById('load-more-btn').onclick = () => {
        inboxPageSkip += INBOX_PAGE_SIZE;
        loadRecentContacts(false);
    };
}

// --- Auth Logic ---

async function handleLogin() {
    const email = document.getElementById('login-email').value;
    const password = document.getElementById('login-password').value;
    const errorDiv = document.getElementById('login-error');
    const btn = document.getElementById('login-btn');
    
    btn.innerText = 'Logging in...';
    btn.disabled = true;

    try {
        const response = await fetch(`${BASE_URL}/v1/enterprises/login`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, password })
        });

        const body = await response.json();
        if (body.isSuccess && body.data.accessToken) {
            Office.context.roamingSettings.set("CRM_TOKEN", body.data.accessToken);
            Office.context.roamingSettings.set("ENTERPRISE_ID", body.data.enterpriseId);
            Office.context.roamingSettings.set("CRM_USER_ID", body.data.enterpriseUserId);
            Office.context.roamingSettings.set("CRM_USER_EMAIL", email);
            Office.context.roamingSettings.set("CRM_USER_NAME", body.data.firstName || body.data.name || email.split('@')[0]);
            Office.context.roamingSettings.saveAsync(() => initApp());
        } else {
            errorDiv.innerText = body.message || "Login failed";
            errorDiv.classList.remove('hidden');
            btn.innerText = 'Login';
            btn.disabled = false;
        }
    } catch (err) {
        errorDiv.innerText = "Connection error";
        errorDiv.classList.remove('hidden');
        btn.innerText = 'Login';
        btn.disabled = false;
    }
}

function handleLogout(e) {
    if(e) e.preventDefault();
    Office.context.roamingSettings.remove("CRM_TOKEN");
    Office.context.roamingSettings.remove("ENTERPRISE_ID");
    Office.context.roamingSettings.remove("CRM_USER_ID");
    Office.context.roamingSettings.remove("CRM_USER_EMAIL");
    Office.context.roamingSettings.remove("CRM_USER_NAME");
    Office.context.roamingSettings.saveAsync(() => {
        document.getElementById('login-email').value = '';
        document.getElementById('login-password').value = '';
        document.getElementById('login-btn').innerText = 'Login';
        document.getElementById('login-btn').disabled = false;
        document.getElementById('login-error').classList.add('hidden');
        showView('login');
    });
}

async function handleConnectMailbox(e) {
    if (e) e.preventDefault();
    const token = Office.context.roamingSettings.get("CRM_TOKEN");
    const enterpriseId = Office.context.roamingSettings.get("ENTERPRISE_ID");

    if (!token || !enterpriseId) {
        alert("You must be logged in to trace emails.");
        return;
    }

    try {
        const response = await fetch(`${BASE_URL}/v1/email-mailbox-imap/oauth/authorization-url`, {
            method: 'POST',
            headers: { 
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({ 
                enterpriseId: enterpriseId,
                provider: "Outlook" 
            })
        });

        if (response.status === 401) {
            alert("Unauthorized. Your session might have expired. Please login again.");
            handleLogout();
            return;
        }

        const result = await response.json();
        // Fallback checks for the URL depending on full API response structure
        const authUrl = result.data || result.url || result.authorizationUrl || (typeof result === 'string' ? result : null);
        
        if (authUrl) {
            // Open the OAuth URL to allow the user to consent.
            window.open(authUrl, '_blank');
        } else {
            console.error("No authorization URL returned:", result);
            alert("Failed to initiate mailbox connection.");
        }
    } catch (err) {
        console.error("Error connecting mailbox:", err);
        alert("A connection error occurred.");
    }
}

// --- Main Detail View Logic ---

function loadCurrentContactData() {
    const item = Office.context.mailbox.item; 
    const email = item.from.emailAddress;
    const name = item.from.displayName;
    
    document.getElementById('back-to-list-btn').style.display = 'block';

    renderDetailHeader(email, name);
    fetchCrmStatus(email, name, item);
}

function renderDetailHeader(email, name) {
    document.getElementById('ui-email').innerText = email;
    document.getElementById('ui-name').innerText = name || email;
    
    let avatarChar = '?';
    if (name && name.length > 0) avatarChar = name[0].toUpperCase();
    else if (email && email.length > 0) avatarChar = email[0].toUpperCase();
    document.getElementById('detail-avatar').innerText = avatarChar;
}

function toggleDetails() {
    const hiddenDiv = document.getElementById('hidden-details');
    const btn = document.getElementById('toggle-details-btn');
    if (hiddenDiv.classList.contains('hidden')) {
        hiddenDiv.classList.remove('hidden');
        btn.innerText = 'Show less ⏶';
    } else {
        hiddenDiv.classList.add('hidden');
        btn.innerText = 'Show more ⏷';
    }
}

async function fetchCrmStatus(email, name, itemFallback = null) {
    const token = Office.context.roamingSettings.get("CRM_TOKEN");
    const enterpriseId = Office.context.roamingSettings.get("ENTERPRISE_ID");

    const statusEl = document.getElementById('ui-status');
    const createBtn = document.getElementById('create-contact-btn');
    
    statusEl.innerText = "Checking...";
    statusEl.style.color = "#5f6368";

    try {
        const response = await fetch(`${BASE_URL}/v1/contact-leads/get-with-filters`, {
            method: 'POST',
            headers: { 
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({ enterpriseId, emailFilter: email, pageNumber: 1, pageSize: 1 })
        });

        const result = await response.json();
        const items = result.data || result.items || [];
        
        if (items.length > 0) {
            const contact = items[0];
            statusEl.innerText = formatContactStatus(contact.statusId);
            statusEl.style.color = "#188038"; // Success Green
            
            document.getElementById('ui-email-ext').innerText = email || '-';
            document.getElementById('ui-name-ext').innerText = name || email || '-';
            document.getElementById('ui-phone').innerText = contact.phone || '-';
            document.getElementById('ui-company').innerText = contact.companyName || '-';
            document.getElementById('ui-location').innerText = `${contact.address || ''} ${contact.postalCode || ''}`.trim() || '-';
            
            createBtn.style.display = 'none';

            // Also replace CRM link if possible
            const contactId = contact.id || contact.contactLeadId;
            if(contactId) {
                document.getElementById('open-crm-btn').onclick = () => window.open(`https://realeye.live/contacts/${contactId}`, '_blank');
            }
        } else {
            // Not in CRM -> Extract AI details from email
            statusEl.innerText = "Not in CRM";
            statusEl.style.color = "#d93025"; // Red
            createBtn.style.display = 'block';
            createBtn.innerText = '➕ Create Contact';
            createBtn.disabled = false;
            
            if (itemFallback) {
                // If we are looking at the current active item, extract body
                itemFallback.body.getAsync(Office.CoercionType.Text, function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        applyExtractedDetails(extractContactDetails(result.value, email, name));
                    }
                });
            } else {
                applyExtractedDetails({ email: email, firstName: name || '', lastName: '', phone: '-', company: '-', address: '-', postalCode: '-' });
            }
        }
    } catch (e) {
        console.error("CRM Fetch Error:", e);
        statusEl.innerText = "Error checking CRM";
    }
}

function applyExtractedDetails(details) {
    document.getElementById('ui-email-ext').innerText = details.email || '-';
    let fullName = (details.firstName + " " + details.lastName).trim();
    document.getElementById('ui-name-ext').innerText = fullName || '-';
    document.getElementById('ui-phone').innerText = details.phone || '-';
    document.getElementById('ui-company').innerText = details.company || '-';
    document.getElementById('ui-location').innerText = `${details.address || ''} ${details.postalCode || ''}`.trim() || '-';
}

function handleCreateContact() {
    const item = Office.context.mailbox.item;
    const token = Office.context.roamingSettings.get("CRM_TOKEN");
    const enterpriseId = Office.context.roamingSettings.get("ENTERPRISE_ID");
    const ownerId = Office.context.roamingSettings.get("CRM_USER_ID");

    const btn = document.getElementById('create-contact-btn');
    btn.innerHTML = '⏳ Extracting...';
    btn.disabled = true;

    item.body.getAsync(Office.CoercionType.Text, async function (result) {
        try {
            let bodyText = result.status === Office.AsyncResultStatus.Succeeded ? result.value : '';
            const details = extractContactDetails(bodyText, item.from.emailAddress, item.from.displayName);

            const payload = {
                enterpriseId: enterpriseId,
                email: details.email,
                firstName: details.firstName,
                lastName: details.lastName,
                phone: details.phone,
                company: details.company,
                address: details.address,
                postalCode: details.postalCode,
                countryId: details.countryId,
                industryId: details.industryId,
                contactLeadOwnerUserId: ownerId,
                contactCreatedByStatusId: 'MANUAL_USER_ENTRY',
                marketingCampaignStatusId: 'MARKETING_CONTACT'
            };

            const response = await fetch(`${BASE_URL}/v1/contact-leads/create`, {
                method: 'POST',
                headers: { 
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${token}`
                },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                btn.innerHTML = '✅ Created';
                setTimeout(() => { loadCurrentContactData(); }, 1500);
            } else {
                btn.innerHTML = '❌ Error';
                setTimeout(() => { btn.innerHTML = '➕ Create Contact'; btn.disabled = false; }, 3000);
            }
        } catch(e) {
            btn.innerHTML = '❌ Error';
            setTimeout(() => { btn.innerHTML = '➕ Create Contact'; btn.disabled = false; }, 3000);
        }
    });
}

// --- Updated List View Logic (uses Outlook REST API directly) ---

function loadRecentContacts(isRefresh = true) {
    const listContainer = document.getElementById('recent-contacts-list');
    const loadMoreBtn = document.getElementById('load-more-btn');

    if (isRefresh) {
        inboxPageSkip = 0;
        listContainer.innerHTML = '<p class="placeholder-text">📬 Loading inbox...</p>';
    } else {
        loadMoreBtn.innerText = '⏳ Loading...';
        loadMoreBtn.disabled = true;
    }

    // Step 1: Check if the REST URL is available.
    // This will be null if the account is Gmail/IMAP connected via Outlook (not Exchange/O365).
    const restUrl = Office.context.mailbox.restUrl;
    if (!restUrl) {
        listContainer.innerHTML = `
            <div style="padding:15px; text-align:center;">
                <p style="font-size:13px; font-weight:600; color:#d93025; margin-bottom:8px;">⚠️ Account Not Supported</p>
                <p style="color:#5f6368; font-size:11px; text-align:left; line-height:1.6;">
                    Your inbox (<strong>${Office.context.mailbox.userProfile.emailAddress}</strong>) is a 
                    <strong>Gmail / IMAP</strong> account connected to Outlook.<br><br>
                    The Outlook REST API only works with <strong>Microsoft 365 / Exchange</strong> accounts.<br><br>
                    ✅ To fix this: Open Outlook and sign in with a <strong>Microsoft / Office 365</strong> account instead of Gmail.
                </p>
            </div>
        `;
        loadMoreBtn.style.display = 'none';
        return;
    }

    
    // Step 2: Get OAuth callback token from Office runtime
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(tokenResult) {
       
        if (tokenResult.status !== Office.AsyncResultStatus.Succeeded) {
            const errCode = tokenResult.error ? tokenResult.error.code : 'unknown';
            
            // Check for the 9018 error to trigger the manual fallback
            if (errCode === 9018 || errCode === 13007) {
                showManualLoginPrompt(listContainer);
            } else {
                listContainer.innerHTML = `<p style="color:red;">Auth Error: ${errCode}</p>`;
            }
            return;
        }

        const accessToken = tokenResult.value;

        // NEW: Logic to use manual token if it exists, otherwise use the automatic one
        const authHeader = window.outlookManualToken 
            ? `Bearer ${window.outlookManualToken}` 
            : `Bearer ${accessToken}`;

        // Build Outlook REST API URL to fetch inbox messages
        const apiUrl = `${restUrl}/v2.0/me/MailFolders/Inbox/messages` +
            `?$top=${INBOX_PAGE_SIZE}` +
            `&$skip=${inboxPageSkip}` +
            `&$select=from,receivedDateTime,subject,isRead` +
            `&$orderby=receivedDateTime desc`;

        // NEW: Fetch using the authHeader variable
        fetch(apiUrl, {
            headers: { 'Authorization': authHeader }
        })
        .then(res => {
            if (!res.ok) throw new Error(`Outlook API error: ${res.status}`);
            return res.json();
        })
        .then(data => {
            const messages = data.value || [];

            // Deduplicate by sender email address
            const seen = new Set();
            const uniqueSenders = [];
            messages.forEach(msg => {
                const senderEmail = msg.from && msg.from.emailAddress ? msg.from.emailAddress.address : null;
                if (senderEmail && !seen.has(senderEmail)) {
                    seen.add(senderEmail);
                    uniqueSenders.push({
                        email: senderEmail,
                        name: msg.from.emailAddress.name || senderEmail,
                        date: msg.receivedDateTime,
                        subject: msg.subject || '(no subject)',
                        isRead: msg.isRead
                    });
                }
            });

            if (isRefresh) {
                renderUniqueSendersList(uniqueSenders, listContainer);
            } else {
                appendSendersList(uniqueSenders, listContainer);
            }

            // Show/hide Load More based on whether we got a full page
            loadMoreBtn.style.display = messages.length < INBOX_PAGE_SIZE ? 'none' : 'block';
            loadMoreBtn.innerText = '📥 Load More';
            loadMoreBtn.disabled = false;
        })
        .catch(err => {
            console.error('Error fetching inbox:', err);
            if (isRefresh) {
                listContainer.innerHTML = `
                    <div style="padding:15px; text-align:center;">
                        <p style="color:#d93025; font-size:12px; margin-bottom:6px;">❌ Failed to load inbox.</p>
                        <p style="color:#5f6368; font-size:11px;">${err.message}</p>
                    </div>
                `;
            }
            loadMoreBtn.innerText = '📥 Load More';
            loadMoreBtn.disabled = false;
        });
    });
}

function buildSenderRow(s) {
    const div = document.createElement('div');
    div.className = 'list-item';
    const initial = s.name ? s.name[0].toUpperCase() : (s.email ? s.email[0].toUpperCase() : '?');
    const dateObj = new Date(s.date);
    const dateStr = dateObj.toLocaleDateString(undefined, { month: 'short', day: 'numeric'});
    div.innerHTML = `
        <div class="avatar">${initial}</div>
        <div style="flex-grow:1; display:flex; justify-content:space-between; align-items:center;">
            <div style="overflow:hidden;">
                <div style="font-weight:${s.isRead ? '400' : '700'}; color:#202124; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${s.name || s.email}</div>
                <div style="font-size:11px; color:#5f6368; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${s.email}</div>
                <div style="font-size:10px; color:#9aa0a6; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${s.subject}</div>
            </div>
            <div style="font-size:10px; color:#9aa0a6; white-space:nowrap; margin-left:6px;">${dateStr}</div>
        </div>
    `;
    div.onclick = () => {
        showView('detail');
        renderDetailHeader(s.email, s.name);
        fetchCrmStatus(s.email, s.name, null);
        document.getElementById('create-contact-btn').style.display = 'none';
    };
    return div;
}

function renderUniqueSendersList(uniqueSenders, container) {
    container.innerHTML = '';
    if (!uniqueSenders || uniqueSenders.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No emails found in inbox.</p>';
        return;
    }
    uniqueSenders.forEach(s => container.appendChild(buildSenderRow(s)));
}

function appendSendersList(uniqueSenders, container) {
    if (!uniqueSenders || uniqueSenders.length === 0) return;
    uniqueSenders.forEach(s => container.appendChild(buildSenderRow(s)));
}

// --- Utility Extract Logic (from Gmail port) ---

function extractContactDetails(body, email, displayName) {
    const details = { 
        email: email, firstName: '', lastName: '', phone: '', company: '', address: '', postalCode: '', countryId: '', industryId: '21'
    };

    if (displayName) {
        const parts = displayName.split(/\s+/);
        details.firstName = parts[0] || '';
        details.lastName = parts.slice(1).join(' ') || '';
    }

    try {
        if (!body) return details;

        const lines = body.split(/\r?\n/);
        const signatureBlock = lines.slice(Math.max(0, lines.length - 60)).join('\n');

        const phoneMatch = signatureBlock.match(/(?:phone|ph|tel|mobile|cell|mob|m|t|p)[:\s]*([+\d\s\-().]{7,20})/i)
          || signatureBlock.match(/(\+?\d[\d\s\-().]{7,18}\d)/);
        if (phoneMatch) details.phone = phoneMatch[1].trim();

        const companyMatch = signatureBlock.match(/(?:^|\n)\s*([A-Z][A-Za-z0-9 &.,]+(?:Inc|LLC|Ltd|Corp|Co|Group|Solutions|Technologies|Tech|Software|Services|Consulting|Partners)\.?)\s*$/im);
        if (companyMatch) details.company = companyMatch[1].trim();

    } catch (e) {
        console.error('Extract err:', e);
    }
    return details;
}

function formatContactStatus(statusId) {
  const statuses = {
    'NONE': '-', 'NEW': 'New', 'OPEN': 'Open', 'IN_PROGRESS': 'In Progress', 
    'OPEN_DEAL': 'Open Deal', 'UNQUALIFIED': 'Unqualified', 'CONNECTED': 'Connected'
  };
  return statuses[statusId] || statusId || '-';
}

function showManualLoginPrompt(container) {
    container.innerHTML = `
        <div style="padding:20px; text-align:center;">
            <p style="font-size:13px; color:#5f6368; margin-bottom:15px;">
                Outlook requires a manual sign-in to sync your inbox.
            </p>
            <button id="manual-auth-btn" style="background:#1a73e8; color:white; border:none; padding:10px 20px; border-radius:4px; cursor:pointer;">
                Sign in to Outlook
            </button>
        </div>
    `;
    
    document.getElementById('manual-auth-btn').onclick = () => {
        const clientID = "b47febaf-cc5b-4be3-b480-63702e916d44"; 
        const redirectUri = "https://wonderful-flower-067571610.7.azurestaticapps.net/taskpane.html";
        const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientID}&response_type=token&scope=openid%20profile%20User.Read%20Mail.Read&redirect_uri=${redirectUri}`;

        // This version includes the event handler to catch the token from the popup
        Office.context.ui.displayDialogAsync(authUrl, { height: 60, width: 40 }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const dialog = asyncResult.value;
            
                // This listens for the "shout" from the script in taskpane.html
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                    // 1. Close the popup
                    dialog.close();
    
                    // 2. Extract token from the hash string (#access_token=...)
                    const hash = arg.message;
                    const tokenMatch = hash.match(/access_token=([^&]+)/);
    
                    if (tokenMatch && tokenMatch[1]) {
                        // Save this token globally so the fetch function can find it
                        window.outlookManualToken = tokenMatch[1];
        
                        // 3. Refresh the contact list
                        loadRecentContacts(true); 
                    }
                });
            }
        });
    };
}
