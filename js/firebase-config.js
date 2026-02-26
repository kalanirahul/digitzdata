// Centralized Firebase configuration
// API key is restricted in Google Cloud Console to digitzdata.com domain only
const firebaseConfig = {
  apiKey: "AIzaSyDyZwl08LB5BkYQUckJLmjh9ErGehZxJ8o",
  authDomain: "dd-consulting-llc.firebaseapp.com",
  projectId: "dd-consulting-llc",
  storageBucket: "dd-consulting-llc.firebasestorage.app",
  messagingSenderId: "251995082210",
  appId: "1:251995082210:web:e797d8e59a92f337407dbb"
};

firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();
