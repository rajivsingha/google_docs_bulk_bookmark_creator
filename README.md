# Google Docs Bulk Bookmark Creator

A Google Apps Script for Google Docs that adds **bookmarks to headings** 
with a simple UI to select which heading levels to include.

You can choose to add bookmarks to:

- Title
- Heading 1 (H1)
- Heading 2 (H2)
- Heading 3 (H3)
- Heading 4 (H4)
- Heading 5 (H5)
- Heading 6 (H6)

The script adds a custom menu and a modal dialog in Google Docs to make 
this easy and non-technical for end users.

---

## Features

- Bookmarks for headings – Adds bookmarks to all headings of the selected 
levels.
- Per-level selection – Choose any combination of Title, H1–H6.
- Select All / Deselect All toggle.
- Avoids duplicates – Does not re-bookmark the same heading text.
- One-click cleanup – Remove all bookmarks in the document from the same 
menu.
- No coding knowledge required for end‑users once installed.

---

## How It Works

The script:

1. Adds a **“Bookmark Manager”** menu to your Google Docs UI.
2. Opens a modal dialog with checkboxes for each heading level.
3. Scans the document for paragraphs styled as Title, Heading 1–6.
4. Adds a bookmark at the start of each matching heading paragraph (if not 
already bookmarked).

---

## Installation

### 1. Open or create a Google Doc

Open the Google Doc where you want to use this functionality.

### 2. Open the Apps Script editor

In the Google Docs menu bar:

- Go to **Extensions → Apps Script**

This opens the Google Apps Script editor in a new tab.

### 3. Add the script code

1. In the Apps Script editor, ensure you have a file named `Code.gs`.
2. Replace its contents with the code from this repo’s 
[`Code.gs`](./Code.gs).
3. Click **Save** and give the project a name, for example: `Heading 
Bookmark Manager`.

### 4. Authorize the script (first-time use)

The first time you run it:

1. In the Apps Script editor, select the function `onOpen` from the 
dropdown and click the ► **Run** button.
2. A permission dialog will appear:
   - Choose your Google account.
   - Click **Advanced** → **Go to *project name***.
   - Click **Allow** to grant the necessary permissions (read and modify 
the current document).

---

## Usage

### 1. Refresh the document

Back in your Google Doc, refresh the page.  
You should now see a new menu in the top bar:

**`Bookmark Manager`**

### 2. Add bookmarks to headings

1. In the Google Doc, go to:

   **Bookmark Manager → Add Bookmarks to Headings**

2. In the dialog that appears:
   - Check the heading levels you want to bookmark:
     - Title, Heading 1, …, Heading 6
   - Optionally click **Select All** to quickly select everything.
3. Click **Add Bookmarks**.
4. Wait for the confirmation message in the dialog, for example:

   > Successfully added 12 bookmark(s) to the selected heading levels.

Bookmarks are added at the start of each matching heading paragraph.

### 3. Remove all bookmarks

To quickly remove all bookmarks from the document:

1. In the menu, go to:

   **Bookmark Manager → Remove All Bookmarks**

2. An alert will display how many bookmarks were removed.

---

## Screenshots

Create an `images` folder in the repository and add these two screenshot 
files:

- `images/bookmark-manager-menu.png`
- `images/heading-selector-dialog.png`

Then capture screenshots in your browser and save them with those exact 
names.

### Custom Menu

![Custom "Bookmark Manager" menu in Google 
Docs](images/bookmark-manager-menu.png)

### Heading Selection Dialog

![Dialog to select heading levels for 
bookmarks](images/heading-selector-dialog.png)

---
