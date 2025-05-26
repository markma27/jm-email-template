// Get the site URL from the current page
const siteUrl = _spPageContextInfo.webAbsoluteUrl;
const listTitle = "JM Email Template Library";

// Construct the API URL
const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/fields`;

// Make the request
fetch(apiUrl, {
    method: 'GET',
    headers: {
        'Accept': 'application/json;odata=verbose'
    },
    credentials: 'include'
})
.then(response => response.json())
.then(data => {
    // Filter out system fields and log the relevant ones
    const fields = data.d.results
        .filter(f => !f.Hidden && !f.InternalName.startsWith('_'))
        .map(f => ({
            Title: f.Title,
            InternalName: f.InternalName,
            Type: f.TypeDisplayName
        }));
    
    console.table(fields);
})
.catch(error => console.error('Error:', error)); 