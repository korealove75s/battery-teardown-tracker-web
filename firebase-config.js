// Firebase connection settings for the shared (cross-device) database.
// Paste the values from Firebase console > Project settings > Your apps > Web app.
// While apiKey is empty, the app runs in browser-only mode (data stays in this browser).
window.BTT_FIREBASE_CONFIG = {
  apiKey: '',
  authDomain: '',
  projectId: '',
  storageBucket: '',
  messagingSenderId: '',
  appId: '',
};

// The single shared team account (Firebase Authentication > Users).
// Its password is the "team password" people type on the sign-in screen.
window.BTT_TEAM_EMAIL = 'team@example.com';
