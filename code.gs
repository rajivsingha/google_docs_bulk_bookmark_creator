/**
 * Google Apps Script to add bookmarks to headings in a Google Doc
 * Creates a custom menu and dialog for selecting heading levels
 * Created by Rajiv Singha - https://www.linkedin.com/in/rajivsingha/
 */

/**
 * Runs when the document is opened
 * Creates a custom menu in the Google Docs UI
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Bookmark Manager')
    .addItem('Add Bookmarks to Headings', 'showHeadingSelector')
    .addItem('Remove All Bookmarks', 'removeAllBookmarks')
    .addToUi();
}

/**
 * Shows a dialog with checkboxes for selecting heading levels
 */
function showHeadingSelector() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            margin: 0;
          }
          h2 {
            color: #1a73e8;
            font-size: 18px;
            margin-bottom: 15px;
          }
          .checkbox-group {
            margin: 10px 0;
          }
          .checkbox-item {
            margin: 8px 0;
            display: flex;
            align-items: center;
          }
          input[type="checkbox"] {
            margin-right: 10px;
            width: 18px;
            height: 18px;
            cursor: pointer;
          }
          label {
            cursor: pointer;
            font-size: 14px;
          }
          .button-group {
            margin-top: 20px;
            display: flex;
            gap: 10px;
          }
          button {
            padding: 10px 20px;
            font-size: 14px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 500;
          }
          #addButton {
            background-color: #1a73e8;
            color: white;
          }
          #addButton:hover {
            background-color: #1557b0;
          }
          #addButton:disabled {
            background-color: #ccc;
            cursor: not-allowed;
          }
          #selectAllButton {
            background-color: #f1f3f4;
            color: #202124;
          }
          #selectAllButton:hover {
            background-color: #e8eaed;
          }
          .status {
            margin-top: 15px;
            padding: 10px;
            border-radius: 4px;
            display: none;
          }
          .status.success {
            background-color: #e6f4ea;
            color: #137333;
            display: block;
          }
          .status.error {
            background-color: #fce8e6;
            color: #c5221f;
            display: block;
          }
        </style>
      </head>
      <body>
        <h2>Select Heading Levels to Bookmark</h2>
        <div class="checkbox-group">
          <div class="checkbox-item">
            <input type="checkbox" id="title" value="TITLE">
            <label for="title">Title</label>
          </div>
          <div class="checkbox-item">
            <input type="checkbox" id="h1" value="HEADING1">
            <label for="h1">Heading 1</label>
          </div>
          <div class="checkbox-item">
            <input type="checkbox" id="h2" value="HEADING2">
            <label for="h2">Heading 2</label>
          </div>
          <div class="checkbox-item">
            <input type="checkbox" id="h3" value="HEADING3">
            <label for="h3">Heading 3</label>
          </div>
          <div class="checkbox-item">
            <input type="checkbox" id="h4" value="HEADING4">
            <label for="h4">Heading 4</label>
          </div>
          <div class="checkbox-item">
            <input type="checkbox" id="h5" value="HEADING5">
            <label for="h5">Heading 5</label>
          </div>
          <div class="checkbox-item">
            <input type="checkbox" id="h6" value="HEADING6">
            <label for="h6">Heading 6</label>
          </div>
        </div>
        <div class="button-group">
          <button id="addButton" onclick="addBookmarks()">Add Bookmarks</button>
          <button id="selectAllButton" onclick="toggleSelectAll()">Select All</button>
        </div>
        <div id="status" class="status"></div>
        
        <script>
          let allSelected = false;
          
          function toggleSelectAll() {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]');
            allSelected = !allSelected;
            checkboxes.forEach(cb => cb.checked = allSelected);
            document.getElementById('selectAllButton').textContent = 
              allSelected ? 'Deselect All' : 'Select All';
          }
          
          function addBookmarks() {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
            const selectedHeadings = Array.from(checkboxes).map(cb => cb.value);
            
            if (selectedHeadings.length === 0) {
              showStatus('Please select at least one heading level.', 'error');
              return;
            }
            
            document.getElementById('addButton').disabled = true;
            document.getElementById('addButton').textContent = 'Processing...';
            
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .addBookmarksToHeadings(selectedHeadings);
          }
          
          function onSuccess(result) {
            showStatus(result, 'success');
            document.getElementById('addButton').disabled = false;
            document.getElementById('addButton').textContent = 'Add Bookmarks';
          }
          
          function onFailure(error) {
            showStatus('Error: ' + error.message, 'error');
            document.getElementById('addButton').disabled = false;
            document.getElementById('addButton').textContent = 'Add Bookmarks';
          }
          
          function showStatus(message, type) {
            const statusDiv = document.getElementById('status');
            statusDiv.textContent = message;
            statusDiv.className = 'status ' + type;
          }
        </script>
      </body>
    </html>
  `)
    .setWidth(400)
    .setHeight(450);
  
  DocumentApp.getUi().showModalDialog(html, 'Add Bookmarks to Headings');
}

/**
 * Adds bookmarks to headings based on selected heading levels
 * @param {Array<string>} selectedHeadings - Array of heading types to bookmark
 * @return {string} Success message
 */
function addBookmarksToHeadings(selectedHeadings) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let bookmarksAdded = 0;
  
  // Convert heading types to ParagraphHeading enum values
  const headingTypes = selectedHeadings.map(heading => 
    DocumentApp.ParagraphHeading[heading]
  );
  
  // Get all paragraphs in the document
  const paragraphs = body.getParagraphs();
  
  // Track existing bookmarks to avoid duplicates
  const existingBookmarks = new Set();
  const bookmarks = doc.getBookmarks();
  bookmarks.forEach(bookmark => {
    const position = bookmark.getPosition();
    if (position) {
      const element = position.getElement();
      existingBookmarks.add(element.asText().getText());
    }
  });
  
  // Iterate through paragraphs and add bookmarks to matching headings
  paragraphs.forEach((paragraph, index) => {
    const heading = paragraph.getHeading();
    
    if (headingTypes.includes(heading)) {
      const text = paragraph.getText().trim();
      
      // Only add bookmark if there's text and no existing bookmark
      if (text && !existingBookmarks.has(text)) {
        // Create a unique bookmark ID based on the heading text
        const bookmarkId = createBookmarkId(text, index);
        
        // Add bookmark at the start of the paragraph
        const position = doc.newPosition(paragraph, 0);
        doc.addBookmark(position);
        
        bookmarksAdded++;
        existingBookmarks.add(text);
      }
    }
  });
  
  return `Successfully added ${bookmarksAdded} bookmark(s) to the selected heading levels.`;
}

/**
 * Creates a clean bookmark ID from heading text
 * (Currently unused, but kept for future extension where bookmark IDs may be used)
 * 
 * @param {string} text - The heading text
 * @param {number} index - The paragraph index
 * @return {string} A clean bookmark ID
 */
function createBookmarkId(text, index) {
  const cleanText = text.toLowerCase()
    .replace(/[^a-z0-9]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
  
  return cleanText + '-' + index;
}

/**
 * Removes all bookmarks from the document
 * Shows a summary alert to the user
 */
function removeAllBookmarks() {
  const doc = DocumentApp.getActiveDocument();
  const bookmarks = doc.getBookmarks();
  const count = bookmarks.length;
  
  bookmarks.forEach(bookmark => {
    bookmark.remove();
  });
  
  DocumentApp.getUi().alert(`Removed ${count} bookmark(s) from the document.`);
}