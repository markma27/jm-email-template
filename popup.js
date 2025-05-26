// SharePoint configuration
const SHAREPOINT_CONFIG = {
    siteUrl: 'https://jaquillardminns.sharepoint.com/sites/JaquillardMinns411',
    listTitle: 'JM Email Template Library',
    apiUrl: 'https://jaquillardminns.sharepoint.com/sites/JaquillardMinns411/_api/web/lists/getbytitle(\'JM%20Email%20Template%20Library\')/items'
};

// Global variables
let templates = [];
let filteredTemplates = [];
let selectedTemplate = null;

// DOM elements
const elements = {
    loading: document.getElementById('loading'),
    error: document.getElementById('error'),
    errorMessage: document.getElementById('errorMessage'),
    mainContent: document.getElementById('mainContent'),
    templateSelect: document.getElementById('templateSelect'),
    categoryFilter: document.getElementById('categoryFilter'),
    previewPane: document.getElementById('previewPane'),
    launchBtn: document.getElementById('launchBtn'),
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
    elements.categoryFilter.addEventListener('change', handleCategoryFilter);
    elements.launchBtn.addEventListener('click', handleLaunchEmail);
    elements.refreshBtn.addEventListener('click', handleRefresh);
    elements.retryBtn.addEventListener('click', handleRetry);
}

// Load templates from SharePoint
async function loadTemplates() {
    showLoading();
    
    try {
        const response = await fetch(SHAREPOINT_CONFIG.apiUrl + '?$select=Id,Title,Email_x0020_Category,Email_x0020_Body&$orderby=Title', {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose'
            },
            credentials: 'include'
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const data = await response.json();
        
        if (data.d && data.d.results) {
            templates = data.d.results.map(item => ({
                id: item.Id,
                title: item.Title || 'Untitled Template',
                category: item.Email_x0020_Category || 'Uncategorized',
                body: item.Email_x0020_Body || ''
            }));
            
            filteredTemplates = [...templates];
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
    } else if (error.message.includes('404')) {
        return 'SharePoint list not found. Please check the list name and URL.';
    } else if (error.message.includes('Failed to fetch')) {
        return 'Network error. Please check your internet connection and try again.';
    } else {
        return `Error: ${error.message}`;
    }
}

// Populate UI with templates and categories
function populateUI() {
    populateTemplateSelect();
    populateCategoryFilter();
}

// Populate template dropdown
function populateTemplateSelect() {
    elements.templateSelect.innerHTML = '<option value="">Choose a template...</option>';
    
    filteredTemplates.forEach(template => {
        const option = document.createElement('option');
        option.value = template.id;
        option.textContent = `${template.title} (${template.category})`;
        elements.templateSelect.appendChild(option);
    });
}

// Populate category filter
function populateCategoryFilter() {
    const categories = [...new Set(templates.map(t => t.category))].sort();
    
    elements.categoryFilter.innerHTML = '<option value="">All Categories</option>';
    
    categories.forEach(category => {
        const option = document.createElement('option');
        option.value = category;
        option.textContent = category;
        elements.categoryFilter.appendChild(option);
    });
}

// Handle template selection
function handleTemplateSelection(event) {
    const templateId = event.target.value;
    
    if (templateId) {
        selectedTemplate = templates.find(t => t.id == templateId);
        if (selectedTemplate) {
            showPreview(selectedTemplate.body);
            elements.launchBtn.disabled = false;
        }
    } else {
        selectedTemplate = null;
        showPlaceholder();
        elements.launchBtn.disabled = true;
    }
}

// Handle category filter
function handleCategoryFilter(event) {
    const selectedCategory = event.target.value;
    
    if (selectedCategory) {
        filteredTemplates = templates.filter(t => t.category === selectedCategory);
    } else {
        filteredTemplates = [...templates];
    }
    
    populateTemplateSelect();
    
    // Reset selection
    selectedTemplate = null;
    showPlaceholder();
    elements.launchBtn.disabled = true;
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

// Handle email launch
async function handleLaunchEmail() {
    if (!selectedTemplate) {
        alert('Please select a template first.');
        return;
    }
    
    try {
        // Prepare email content
        const emailBody = encodeURIComponent(selectedTemplate.body);
        const subject = encodeURIComponent(selectedTemplate.title);
        
        // Try Outlook Web first
        const outlookUrl = `https://outlook.office.com/mail/deeplink/compose?subject=${subject}&body=${emailBody}`;
        
        // Open in new tab
        const newTab = window.open(outlookUrl, '_blank');
        
        // Fallback to mailto if Outlook Web fails
        if (!newTab) {
            const mailtoUrl = `mailto:?subject=${subject}&body=${emailBody}`;
            window.location.href = mailtoUrl;
        }
        
        // Close the popup after a short delay
        setTimeout(() => {
            window.close();
        }, 500);
        
    } catch (error) {
        console.error('Error launching email:', error);
        alert('Failed to launch email. Please try again.');
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