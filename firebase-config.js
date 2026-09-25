// Firebase connection settings for the shared (cross-device) database.
// Paste the values from Firebase console > Project settings > Your apps > Web app.
// While apiKey is empty, the app runs in browser-only mode (data stays in this browser).
window.BTT_FIREBASE_CONFIG = {
  apiKey: 'AIzaSyCmWTs31VacL3kdkit0SqAkBXoGGj3siGc',
  authDomain: 'love75s.firebaseapp.com',
  projectId: 'love75s',
  storageBucket: 'love75s.firebasestorage.app',
  messagingSenderId: '969722345286',
  appId: '1:969722345286:web:404d69162340cb28a95747',
};

// Users sign in with their own ID (email) and password (Firebase Authentication > Users).
// Only this account sees the Reset button that deletes shared data.
window.BTT_ADMIN_EMAIL = 'korealove75s@gmail.com';