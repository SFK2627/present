# Presentation Hub - Static Website Version

Ito yung **website lang** version. Hindi ito React/Vite project, kaya hindi kailangan ng `npm install` or `npm run dev`.

## Paano gamitin

1. I-extract ang ZIP.
2. Buksan ang folder.
3. I-double click ang `index.html`.
4. Upload ng PDF or PPTX.

Pwede rin i-upload ang buong folder sa Netlify, Firebase Hosting, GitHub Pages, Hostinger, cPanel, or kahit anong static hosting.

## Important notes

### PDF
PDF files are fully viewable using PDF.js:

- high quality page rendering
- thumbnails
- next/previous
- jump to page
- fullscreen
- zoom
- auto slide
- timer overlay

### PPTX
PPTX is supported in **static fallback mode**:

- app reads PPTX slide count
- app extracts slide text
- app displays simplified slide previews

But for exact PowerPoint layout, images, animations, fonts, and formatting, convert the PPTX to PDF first, then upload the PDF. This is a browser limitation. Accurate PowerPoint rendering normally needs a conversion backend like LibreOffice or Microsoft Office rendering.

## Phone remote control

The app includes QR remote support, but phone-to-desktop syncing needs Firebase because two different devices need a realtime backend.

To enable:

1. Create a Firebase project.
2. Create a Web App.
3. Enable Firestore Database.
4. Enable Anonymous Authentication.
5. Edit `firebase-config.js` and paste your Firebase web config.
6. Upload the website folder to a static host.
7. Open a presentation, click **Remote QR**, then scan from phone.

Example `firebase-config.js`:

```js
window.PRESENTATION_HUB_FIREBASE_CONFIG = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_PROJECT.firebaseapp.com",
  projectId: "YOUR_PROJECT_ID",
  storageBucket: "YOUR_PROJECT.appspot.com",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID"
};
```

Recommended Firestore rules for testing:

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

For production, tighten the rules with session ownership or access tokens.

## Files

- `index.html` - main website
- `styles.css` - full responsive UI design
- `app.js` - presentation logic
- `firebase-config.js` - optional Firebase config
- `manifest.webmanifest` - PWA metadata
- `sw.js` - offline shell caching
- `icon.svg` - app icon

## Keyboard shortcuts

- Arrow Right: next page/slide
- Arrow Left: previous page/slide
- F: fullscreen
- Esc: exit fullscreen
