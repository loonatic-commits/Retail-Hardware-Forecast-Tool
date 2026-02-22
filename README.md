# SnapVault - Photo Scanner & Album Manager

A Photomyne-style web app for scanning, enhancing, organizing, and exporting photos. Runs entirely in the browser with no backend required -- all data is stored locally using IndexedDB.

## Features

- **Photo Scanning** -- Use your device camera to capture old photos with a guided scanning interface
- **Auto Enhancement** -- Automatic image analysis and correction (brightness, contrast, saturation, warmth, sharpness)
- **Photo Editor** -- Full editing suite with adjustments, 12 preset filters, crop (free/fixed ratios), rotate, and flip
- **Album Management** -- Create, rename, and delete albums; organize photos across albums
- **Gallery View** -- Grid gallery with sort options (newest, oldest, name), search, and multi-select
- **Lightbox Viewer** -- Full-screen photo viewer with keyboard navigation, info panel, download, and delete
- **Multi-Select** -- Batch operations: add to album, download, delete
- **Drag & Drop Upload** -- Drop image files anywhere to import
- **Export** -- Download individual photos or export entire library
- **Offline-First** -- All data persisted in IndexedDB, works without internet
- **Responsive** -- Works on desktop, tablet, and mobile

## Getting Started

Open `index.html` in any modern browser. No build step or server required.

## Project Structure

```
index.html              Main HTML shell
css/
  styles.css            Core layout and component styles
  gallery.css           Photo grid and lightbox styles
  editor.css            Photo editor styles
js/
  db.js                 IndexedDB persistence layer
  imageProcessor.js     Canvas-based image processing engine
  editor.js             Photo editor controller
  app.js                Main application controller
```

## Browser Support

Requires a modern browser with support for IndexedDB, Canvas API, and getUserMedia (for camera scanning).
