# Microsoft Word – Automated Picture Insertion and Resizing Using VBA

This project contains two Microsoft Word VBA modules that automate picture insertion and resizing.

The first module automatically searches for image files matching text placeholders in a Word document and inserts the corresponding pictures.

The second module automatically resizes inserted pictures to fit their designated table cells, eliminating the need to manually adjust image dimensions.

Both modules are installed in the global `Normal.dotm` template, allowing the automation to work with regular `.docx` documents without requiring macro-enabled report files.

## Features

- Automatically inserts pictures based on document placeholders.
- Allows users to select an image folder.
- Searches document stories, including tables and headers.
- Matches image filenames against document placeholders.
- Supports JPG, JPEG, PNG, BMP, GIF, TIF, and TIFF.
- Preserves original image aspect ratios.
- Limits inserted pictures to approximately 6 inches wide and 4 inches high.
- Automatically triggers picture resizing after insertion.
- Fits pictures inside their designated table cells.
- Identifies missing image matches.
- Displays a summary of successfully inserted pictures and missing placeholders.

## Repository Structure

```text
Word-Picture-Automation/
│
├── README.md
│
├── VBA/
│   ├── AutoResizeImage.bas
│   └── InsertPicturesFromFolder.bas
│
└── Screenshots/
```

## Requirements

- Microsoft Word Desktop for Windows
- VBA support enabled
- An image folder containing the pictures to insert

## Installation

Both VBA modules must be installed in the global `Normal.dotm` template.

### 1. Install AutoResizeImage.bas

**Important:** The automatic resizing functionality depends on the `AutoResizeImage.bas` module.

If you have already installed this module in your `Normal.dotm` template, you can skip this step.

Otherwise:

1. Open Microsoft Word.
2. Press `Alt + F11` to open the VBA Editor.
3. Locate `Normal` in the Project Explorer.
4. Right-click `Modules`.
5. Select **Import File**.
6. Locate and select `AutoResizeImage.bas`.

### 2. Install InsertPicturesFromFolder.bas

1. In the VBA Editor, locate `Normal > Modules`.
2. Right-click `Modules`.
3. Select **Import File**.
4. Locate and select `InsertPicturesFromFolder.bas`.

Both modules should now appear under `Normal > Modules`.

### 3. Save Normal.dotm

1. Press `Ctrl + S` inside the VBA Editor.
2. Close the VBA Editor.
3. If Word asks whether to save changes to the global template, select **Save**.

The automation is now ready to use.

## How to Use

### 1. Prepare Your Word Document

Insert text placeholders wherever you want pictures to appear.

Example:

```text
{{SITE PHOTO}}
{{SERVER ROOM}}
{{NETWORK CABINET}}
{{ELECTRICAL PANEL}}
```

### 2. Prepare Your Image Folder

Name your pictures according to their corresponding placeholders.

Example:

```text
Pictures/
│
├── SITE PHOTO.jpg
├── SERVER ROOM.jpg
├── NETWORK CABINET.jpg
└── ELECTRICAL PANEL.jpg
```

The matching function ignores spaces, underscores, hyphens, and periods when comparing filenames with placeholders.

### 3. Run the Automation

1. Open your Word document.
2. Press `Alt + F8`.
3. Select `InsertPicturesFromFolder`.
4. Click **Run**.
5. Select the folder containing your pictures.
6. Wait for the automation to finish.

The macro will:

1. Search the document for picture placeholders.
2. Find matching image files in the selected folder.
3. Replace matching placeholders with their corresponding pictures.
4. Automatically trigger the picture-resizing functionality.
5. Display a summary of inserted pictures and missing matches.

### 4. Review the Results

After processing, Word displays a summary containing:

- Number of pictures successfully inserted.
- Number of placeholders without matching pictures.
- Names of unmatched placeholders.

Unmatched placeholders remain in the document so they can be addressed manually.

## Important Notes

- Both VBA modules must be installed in `Normal.dotm`.
- Install `AutoResizeImage.bas` first if it has not already been installed.
- Save report documents as `.docx`.
- Keep image filenames consistent with their corresponding placeholders.
- Pictures are embedded in the document rather than linked to external files.
- Original image proportions are preserved.
- The insertion macro searches document stories, including tables and headers.

## License

This project is available for personal learning, modification, and use. You may adapt the VBA modules to suit your own document automation requirements.

## Development Note

AI-assisted development. The code was reviewed and tested by the author.
<br>
*

<p align="center">
  <a href="https://github.com/arrianmunsod">
    <strong>← BACK TO ARRIAN'S GITHUB PROFILE</strong>
  </a>
  <br>
  <sub>Explore more projects, scripts, and technical documentation.</sub>
</p>
