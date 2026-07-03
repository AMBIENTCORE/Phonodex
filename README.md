# Phonodex - Music Metadata Manager (with Automatic Metadata Fetching)

![image](assets/full.png)

Currently compatible with **Windows only** (due to clipboard operations).

A desktop application for managing and organizing music file metadata.
By having correct Artist and Title and/or Album, it can automatically fetch the Cover Art, Catalog Number and Year.
Only these three fields are applied automatically; the rest (Genre, Comment, Track, etc.) are assumed to be filled in already or are left for you to edit.
Catalog Number and Year are fetched from the Discogs API (max. 60 calls/minute), while Album Art is retrieved from public cover-art sources.

You can get a Discogs API key by doing the following:
- Log in to [Discogs](https://www.discogs.com/),
- Go to [Settings -> Developer](https://www.discogs.com/settings/developers),
- Click Generate New Token.

## Backstory

I identified the need for an automatic way to fill in certain metadata fields. I used to do all tagging manually using Discogs, so I decided to create a MP3Tag extension to do that for me.
After that was developed, I decided to turn it into an actual application with a GUI that's familiar and easy to use.
All functionality was achieved in around 10 days, and I've recently been working to turn the main.py monolith into separate modules.

**For some context**: I'm ***not*** a developer, I have no coding background except VERY basic Python usage before college, and C in college.
However, I'm naturally curious and have fun with logic and troubleshooting, so I decided to try this 'small' project using different **AI tools**.
This will explain why, for a programmer, the code will definitely have things that could be optimized or might not look logical at first. Hence, I decided to create this project on GitHub.

## Features

- 🎵 Batch edit ID3 tags and metadata for multiple audio formats (MP3, FLAC, M4A, OGG, WMA, WAV)
- 🔍 Automatic metadata fetching (Album Art, Catalog Number, Year)
- 🚦 Automatic API request limiter (max 60 requests/minute)
- ✏️ Inline editing — double-click any cell in the table to edit it directly
- 🏷️ Genre autocomplete with an editable dictionary, plus a "Scan Genres" feature that merges genre tags found in a folder
- 🖼️ Album Art clipboard support (copy / paste / remove via right-click)
- 🔎 Live filtering across all fields and sortable columns (click a header to sort)
- 🎨 Color-coded rows (green = updated, red = failed) for quick status at a glance
- 🖱️ Right-click context menu to play a file or show it in Explorer
- 📁 Exporting files from the imported folders to the final Music Collection folder
- 🔄 Drag-and-drop file support
- 📊 Visual progress tracking for both file processing and API usage

## Supported Audio Formats
- MP3 (ID3 tags)
- FLAC
- M4A/MP4
- OGG Vorbis
- WMA (Windows Media Audio)
- WAV

## Dependencies

- tkinter / tkinterdnd2 - UI framework with drag-and-drop support
- mutagen - Audio metadata handling
- requests - API communication
- Pillow - Album art image processing and display
- pywin32 - Windows clipboard support for Album Art


## Installation

### Option A: Installer (recommended for most users)

Download the latest installer from the [Releases page](https://github.com/AMBIENTCORE/Phonodex/releases/tag/1.0) and run it to install Phonodex like any other desktop app. No Python setup required — it creates Start menu and optional desktop shortcuts and can launch on finish.

### Option B: Run from source (developers)

1. Clone the repository
2. Optional: Setup a virtual environment so dependencies are not installed globally:

```bash
python -m venv venv
venv\Scripts\activate
```

3. Install dependencies

```bash
pip install -r requirements.txt
```

4. Run the application:

```bash
python main.py
```

## Configuration

When running the packaged app, Phonodex stores its settings in `%LOCALAPPDATA%\Phonodex\`:

- `api_key.txt` - your saved Discogs API token
- `folder_format_settings.json` - your custom export folder structure
- `genres.txt` - your editable genre dictionary (one genre per line; lines starting with `#` are comments)

When running from source, these files live in the project root. The default export structure format (that I use) is:

```
D:\Music\Collection\%genre%\%year%\[%catalognumber%] %albumartist% - %album%\%artist% - %title%
```

![image](assets/folder_structure.png)

You can customize this format using the following placeholders:
- %genre%
- %year%
- %catalognumber%
- %albumartist%
- %album%
- %artist%
- %title%

## Usage

1. Launch the application
2. Add your Discogs API token in the top input field, and click SAVE

![image](assets/apikey.png)

3. Load files either by:
   - Clicking the "FILES" button to select individual files
   - Clicking the "FOLDER" button to select a folder
   - Dragging and dropping files into the application
4. Select the files you want to process (CTRL+A selects all files loaded)
5a. Click on the checkboxes for 'art', 'year' and/or 'catalog'. Then, click "PROCESS" to fetch metadata automatically
5b. Edit metadata directly by double-clicking the field on the table **(click the Enter key to save)**, or for selected files on the left pane **(click SAVE METADATA)**
6. Use "EXPORT" to organize files into your collection

### Tips

- Use the filter box above the table to quickly narrow down files, and click any column header to sort by it.
- Right-click a row to play the file or show it in Explorer; right-click the album art to copy, paste, or remove it.
- Start typing in the Genre field to get autocomplete suggestions, or use "SCAN GENRES" to merge genre tags found across a folder into your dictionary.
- Rows turn green when successfully updated and red when a fetch fails.

## Contributing

Contributions are welcome!
Please feel free to submit a Pull Request.

## License

MIT
