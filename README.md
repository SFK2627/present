# Presentation Hub - Website Only Version

This is the static website version of Presentation Hub. You can open `index.html` directly or upload the folder to static hosting.

## What works locally

- Upload multiple PDF and PPTX files
- Search and sort presentation cards
- Open presentations
- PDF rendering with real page layout, images, colors, and thumbnails using PDF.js
- PPTX visual rendering using PPTXjs when online/CDN scripts are available
- Next / Previous / Jump page
- Fullscreen
- Auto-slide timing
- Timer overlay
- Dark mode

## Important PPTX note

The app now attempts to render PowerPoint slides visually, including colors, images, and layouts, through PPTXjs.

However, browser-only PowerPoint rendering is still not as perfect as PowerPoint itself. For the most accurate output, especially for complex PowerPoint files with custom fonts, animations, SmartArt, or unusual layouts, export the PPTX as PDF first and upload the PDF. PDF mode is the most accurate mode.

## Firebase remote setup

Your Firebase config is already placed in `firebase-config.js`.

To make phone remote work:

1. Open Firebase Console.
2. Enable Authentication > Anonymous.
3. Create Firestore Database.
4. Add these Firestore rules:

```txt
rules_version = '2';

service cloud.firestore {
  match /databases/{database}/documents {
    match /presentationHubSessions/{sessionId} {
      allow read, write: if request.auth != null;
    }
  }
}
```

## Best way to use with phone remote

Upload the whole folder to Netlify, Firebase Hosting, GitHub Pages, or any web host. Phone remote works best when both laptop and phone open the same hosted website URL.


## Update notes

This version improves GitHub Pages usage:

- Heavy PowerPoint rendering libraries are lazy-loaded only when a PPTX is opened, so the home screen loads faster.
- Firebase sign-in no longer blocks the dashboard while loading.
- Fullscreen now uses a true presentation mode: sidebar is hidden, the viewer fills the whole screen, and the toolbar appears only when the mouse moves.
- Keyboard shortcuts still work in viewer mode: Arrow Right, Arrow Left, F, and Esc.
- The service worker now uses a fresh cache name and network-first updates so GitHub Pages changes are less likely to stay stuck on an old cached version.

## Latest update notes

This build improves the presentation behavior requested for GitHub Pages:

- Fullscreen now targets the entire monitor area and removes sidebar spacing, borders, rounded corners, and shadows.
- QR and setup modals are moved into the active fullscreen container, so they remain visible while the viewer is fullscreen.
- PDF rendering now uses high-DPI canvas rendering for clearer fullscreen output.
- Phone remote preview now sends a larger, clearer current-page image.
- Host phones can pinch on the preview to zoom the desktop viewer in real time.
- PPTX no longer falls back to a plain text-only slide view when visual rendering fails. It shows a clear warning instead, because exact PPTX rendering needs visual rendering or PDF export.

For PowerPoint files that must look exactly like the original, use PowerPoint/WPS/Canva export to PDF first, then upload the PDF into Presentation Hub. That is the recommended exact mode.

## Timer, autoplay, and phone preview update

This build adds the requested production presentation controls:

- Auto Play now has a clear **Timing Mode**:
  - **Global**: every page/slide uses the same interval.
  - **Per-slide**: manually set the timing for the current page/slide; slides without custom timing fall back to the global interval.
- The slide timer resets every time Auto Play moves to a new page/slide.
- Pause freezes both Auto Play and the visible timer.
- Resume continues from the paused time.
- Stop fully stops Auto Play and resets the visible timer.
- Phone remote layout has been tightened so the preview no longer overlaps or pushes past the control sections.
- The phone remote now has **Fullscreen Preview**. Open it, rotate the phone landscape, then pinch and drag the preview. The desktop viewer receives the same zoom and focus area.

Reminder: for exact PowerPoint colors, pictures, fonts, and layouts, the most reliable workflow is still to export PPTX to PDF first, then upload the PDF.
