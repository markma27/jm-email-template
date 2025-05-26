// SharePoint configuration
const SHAREPOINT_CONFIG = {
    siteUrl: 'https://jaquillardminns.sharepoint.com/sites/JaquillardMinns411',
    listTitle: 'JM Email Template Library',
    get apiUrl() {
        // Use GetByTitle and encode the list name properly
        return `${this.siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(this.listTitle)}')/items`;
    }
};

// Global variables
let templates = [];
let selectedTemplate = null;

// DOM elements
const elements = {
    loading: document.getElementById('loading'),
    error: document.getElementById('error'),
    errorMessage: document.getElementById('errorMessage'),
    mainContent: document.getElementById('mainContent'),
    templateSelect: document.getElementById('templateSelect'),
    emailTitleSection: document.getElementById('emailTitleSection'),
    emailTitle: document.getElementById('emailTitle'),
    previewPane: document.getElementById('previewPane'),
    copyTitleBtn: document.getElementById('copyTitleBtn'),
    copyBodyBtn: document.getElementById('copyBodyBtn'),
    refreshBtn: document.getElementById('refreshBtn'),
    retryBtn: document.getElementById('retryBtn')
};

// Initialize the extension
document.addEventListener('DOMContentLoaded', async () => {
    setupEventListeners();
    await loadTemplates();
});

// Event listeners
function setupEventListeners() {
    elements.templateSelect.addEventListener('change', handleTemplateSelection);
    elements.copyTitleBtn.addEventListener('click', handleCopyTitle);
    elements.copyBodyBtn.addEventListener('click', handleCopyBody);
    elements.refreshBtn.addEventListener('click', handleRefresh);
    elements.retryBtn.addEventListener('click', handleRetry);
}

// Check SharePoint authentication
async function checkSharePointAuth() {
    try {
        const response = await fetch(`${SHAREPOINT_CONFIG.siteUrl}/_api/web`, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose'
            },
            credentials: 'include'
        });

        if (!response.ok) {
            // If not authenticated, open SharePoint in a new tab
            window.open(SHAREPOINT_CONFIG.siteUrl, '_blank');
            throw new Error('Please sign in to SharePoint and try again.');
        }

        return true;
    } catch (error) {
        console.error('Auth check error:', error);
        throw new Error('Authentication required. Please sign in to SharePoint and try again.');
    }
}

// Get list fields to verify column names
async function getListFields() {
    try {
        const response = await fetch(`${SHAREPOINT_CONFIG.apiUrl}/fields`, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose'
            },
            credentials: 'include'
        });

        if (!response.ok) {
            throw new Error(`Failed to get list fields: ${response.status}`);
        }

        const data = await response.json();
        console.log('Available fields:', data.d.results.map(f => ({
            title: f.Title,
            internalName: f.InternalName,
            staticName: f.StaticName
        })));
        return data.d.results;
    } catch (error) {
        console.error('Error getting fields:', error);
        throw error;
    }
}

// Load templates from SharePoint
async function loadTemplates() {
    showLoading();
    
    try {
        // First try to verify the list exists and log debug info
        console.log('Attempting to access SharePoint site:', SHAREPOINT_CONFIG.siteUrl);
        
        const listInfoUrl = `${SHAREPOINT_CONFIG.siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(SHAREPOINT_CONFIG.listTitle)}')?$select=Title,ItemCount`;
        console.log('Checking list URL:', listInfoUrl);

        const listResponse = await fetch(listInfoUrl, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose'
            },
            credentials: 'include'
        });

        if (!listResponse.ok) {
            console.error('List check failed:', listResponse.status, listResponse.statusText);
            const errorText = await listResponse.text();
            console.error('Error details:', errorText);
            throw new Error(`Failed to verify list: ${listResponse.status}`);
        }

        const listInfo = await listResponse.json();
        console.log('List info:', listInfo);

        // Now fetch the items
        const itemsUrl = `${SHAREPOINT_CONFIG.apiUrl}?$select=ID,Title,EmailCategory,EmailTitle,EmailBody&$orderby=Title`;
        console.log('Fetching items URL:', itemsUrl);

        const response = await fetch(itemsUrl, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose'
            },
            credentials: 'include'
        });

        if (!response.ok) {
            console.error('Items fetch failed:', response.status, response.statusText);
            const errorText = await response.text();
            console.error('Error details:', errorText);
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const data = await response.json();
        console.log('SharePoint response:', data);
        
        if (data.d && data.d.results) {
            templates = data.d.results.map(item => ({
                id: item.ID || item.Id,
                title: item.Title || 'Untitled Template',
                emailTitle: item.EmailTitle || item.Title || 'Untitled Template',
                category: item.EmailCategory || 'Uncategorized',
                body: item.EmailBody || ''
            }));
            
            populateUI();
            showMainContent();
        } else {
            throw new Error('Invalid response format from SharePoint');
        }
    } catch (error) {
        console.error('Error loading templates:', error);
        showError(getErrorMessage(error));
    }
}

// Get user-friendly error message
function getErrorMessage(error) {
    if (error.message.includes('401')) {
        return 'Authentication required. Please sign in to SharePoint and try again.';
    } else if (error.message.includes('403')) {
        return 'Access denied. You may not have permission to access this SharePoint list.';
    } else if (error.message.includes('404') || error.message.includes('List not found')) {
        return 'SharePoint list not found. Please check the list name and URL.';
    } else if (error.message.includes('400')) {
        return 'Bad request. Please check SharePoint configuration and try again.';
    } else if (error.message.includes('Failed to fetch')) {
        return 'Network error. Please check your internet connection and try again.';
    } else {
        return `Error: ${error.message}`;
    }
}

// Populate UI with templates
function populateUI() {
    populateTemplateSelect();
}

// Populate template dropdown
function populateTemplateSelect() {
    elements.templateSelect.innerHTML = '<option value="">Choose a template...</option>';
    
    templates.forEach(template => {
        const option = document.createElement('option');
        option.value = template.id;
        option.textContent = template.title;
        elements.templateSelect.appendChild(option);
    });
}

// Handle template selection
function handleTemplateSelection(event) {
    const templateId = event.target.value;
    
    if (templateId) {
        selectedTemplate = templates.find(t => t.id == templateId);
        if (selectedTemplate) {
            showEmailTitle(selectedTemplate.emailTitle || selectedTemplate.title);
            showPreview(selectedTemplate.body);
            elements.copyTitleBtn.disabled = false;
            elements.copyBodyBtn.disabled = false;
        }
    } else {
        selectedTemplate = null;
        hideEmailTitle();
        showPlaceholder();
        elements.copyTitleBtn.disabled = true;
        elements.copyBodyBtn.disabled = true;
    }
}

// Show template preview
function showPreview(htmlContent) {
    if (htmlContent && htmlContent.trim()) {
        // Sanitize and display HTML content
        elements.previewPane.innerHTML = sanitizeHtml(htmlContent);
    } else {
        elements.previewPane.innerHTML = '<p class="placeholder">No content available for this template</p>';
    }
}

// Show placeholder in preview
function showPlaceholder() {
    elements.previewPane.innerHTML = '<p class="placeholder">Select a template to see preview</p>';
}

// Show email title
function showEmailTitle(title) {
    elements.emailTitle.textContent = title || 'No title available';
    elements.emailTitleSection.style.display = 'block';
}

// Hide email title
function hideEmailTitle() {
    elements.emailTitle.textContent = '';
    elements.emailTitleSection.style.display = 'none';
}

// Basic HTML sanitization (remove script tags and dangerous attributes)
function sanitizeHtml(html) {
    // Create a temporary div to parse HTML
    const temp = document.createElement('div');
    temp.innerHTML = html;
    
    // Remove script tags
    const scripts = temp.querySelectorAll('script');
    scripts.forEach(script => script.remove());
    
    // Remove dangerous attributes
    const allElements = temp.querySelectorAll('*');
    allElements.forEach(element => {
        const dangerousAttrs = ['onload', 'onerror', 'onclick', 'onmouseover', 'onfocus', 'onblur'];
        dangerousAttrs.forEach(attr => {
            if (element.hasAttribute(attr)) {
                element.removeAttribute(attr);
            }
        });
    });
    
    return temp.innerHTML;
}

// Convert HTML to formatted plain text for email body (legacy function - kept for compatibility)
function htmlToFormattedText(html) {
    return convertHtmlToFormattedPlainText(html);
}

// Convert HTML to formatted plain text with bold markers and proper paragraph spacing
function convertHtmlToFormattedPlainText(html) {
    console.log('Original HTML:', html); // Debug log
    
    // Create a temporary div to parse HTML
    const temp = document.createElement('div');
    temp.innerHTML = html;
    
    // First, handle nested formatting by processing the DOM elements directly
    const processElement = (element) => {
        if (element.nodeType === Node.TEXT_NODE) {
            return element.textContent;
        }
        
        if (element.nodeType === Node.ELEMENT_NODE) {
            let content = '';
            for (let child of element.childNodes) {
                content += processElement(child);
            }
            
            const tagName = element.tagName.toLowerCase();
            
                         // Apply formatting based on tag
             switch (tagName) {
                 case 'b':
                 case 'strong':
                     return `**${content}**`;
                 case 'i':
                 case 'em':
                     return `*${content}*`;
                 case 'u':
                     return `_${content}_`;
                 case 'p':
                     // For "No Spacing" style: content + line break + blank line
                     return content.trim() + '\n\n';
                 case 'div':
                     // Treat divs like paragraphs for consistent spacing
                     return content.trim() + '\n\n';
                 case 'br':
                     // Single line break
                     return '\n';
                 case 'li':
                     return '• ' + content.trim() + '\n';
                 case 'ul':
                 case 'ol':
                     return '\n' + content + '\n';
                 case 'h1':
                 case 'h2':
                 case 'h3':
                 case 'h4':
                 case 'h5':
                 case 'h6':
                     return `\n**${content.trim()}**\n\n`;
                 default:
                     return content;
             }
        }
        
        return '';
    };
    
    // Process the HTML structure
    let formattedText = '';
    for (let child of temp.childNodes) {
        formattedText += processElement(child);
    }
    
         // Clean up for "No Spacing" style formatting
     // Remove any trailing spaces from lines
     formattedText = formattedText.replace(/[ \t]+\n/g, '\n');
     // Ensure consistent double line breaks between paragraphs (blank line)
     formattedText = formattedText.replace(/\n\s*\n\s*\n+/g, '\n\n');
     // Clean up any spaces at start/end
     formattedText = formattedText.replace(/^\s+|\s+$/g, '');
     // Replace multiple spaces/tabs with single space within lines
     formattedText = formattedText.replace(/[ \t]+/g, ' ');
     // Ensure we end with proper line breaks if content exists
     if (formattedText && !formattedText.endsWith('\n\n')) {
         formattedText = formattedText.replace(/\n*$/, '');
     }
    
    console.log('Formatted text:', formattedText); // Debug log
    
    return formattedText;
}

// Handle copy title to clipboard
async function handleCopyTitle() {
    if (!selectedTemplate) {
        alert('Please select a template first.');
        return;
    }
    
    try {
        const title = selectedTemplate.emailTitle || selectedTemplate.title;
        await navigator.clipboard.writeText(title);
        
        // Show success feedback
        const originalText = elements.copyTitleBtn.textContent;
        elements.copyTitleBtn.textContent = 'Copied!';
        elements.copyTitleBtn.style.backgroundColor = '#107c10';
        
        setTimeout(() => {
            elements.copyTitleBtn.textContent = originalText;
            elements.copyTitleBtn.style.backgroundColor = '#0078d4';
        }, 2000);
        
    } catch (error) {
        console.error('Error copying title:', error);
        alert(`Failed to copy title. Please manually copy: "${selectedTemplate.emailTitle || selectedTemplate.title}"`);
    }
}

// Handle copy body to clipboard
async function handleCopyBody() {
    if (!selectedTemplate) {
        alert('Please select a template first.');
        return;
    }
    
    try {
        // Get the exact HTML content from the preview pane
        const previewContent = elements.previewPane.innerHTML;
        
        // Copy HTML content to clipboard using the Clipboard API
        if (navigator.clipboard && navigator.clipboard.write) {
            // Create clipboard items with both HTML and plain text
            const htmlBlob = new Blob([previewContent], { type: 'text/html' });
            const textBlob = new Blob([elements.previewPane.textContent || elements.previewPane.innerText], { type: 'text/plain' });
            
            const clipboardItem = new ClipboardItem({
                'text/html': htmlBlob,
                'text/plain': textBlob
            });
            
            await navigator.clipboard.write([clipboardItem]);
            
            // Show success feedback
            const originalText = elements.copyBodyBtn.textContent;
            elements.copyBodyBtn.textContent = 'Copied!';
            elements.copyBodyBtn.style.backgroundColor = '#107c10';
            
            setTimeout(() => {
                elements.copyBodyBtn.textContent = originalText;
                elements.copyBodyBtn.style.backgroundColor = '#0078d4';
            }, 2000);
            
        } else {
            // Fallback for older browsers
            const textContent = elements.previewPane.textContent || elements.previewPane.innerText;
            await navigator.clipboard.writeText(textContent);
        }
        
    } catch (error) {
        console.error('Error copying body:', error);
        
        // Manual copy fallback
        try {
            // Select the preview content for manual copying
            const range = document.createRange();
            range.selectNodeContents(elements.previewPane);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            
            alert(`Unable to copy automatically. The email body is now selected.\n\nPlease press Ctrl+C to copy the selected content.`);
            
        } catch (selectError) {
            console.error('Error selecting content:', selectError);
            alert(`Please manually copy the content from the preview pane.`);
        }
    }
}

// Handle refresh
async function handleRefresh() {
    elements.refreshBtn.disabled = true;
    await loadTemplates();
    elements.refreshBtn.disabled = false;
}

// Handle retry
async function handleRetry() {
    await loadTemplates();
}

// UI state management
function showLoading() {
    elements.loading.style.display = 'block';
    elements.error.style.display = 'none';
    elements.mainContent.style.display = 'none';
}

function showError(message) {
    elements.loading.style.display = 'none';
    elements.error.style.display = 'block';
    elements.mainContent.style.display = 'none';
    elements.errorMessage.textContent = message;
}

function showMainContent() {
    elements.loading.style.display = 'none';
    elements.error.style.display = 'none';
    elements.mainContent.style.display = 'block';
}

// Utility function to escape HTML for safe display
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
} 