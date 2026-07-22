
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";


const firebaseConfig = {
  apiKey: "AIzaSyDKRU0gq6ewoAGEsUDaS3leBGGfa0Zgu6A",
  authDomain: "conculeyt.firebaseapp.com",
  projectId: "conculeyt",
  storageBucket: "conculeyt.firebasestorage.app",
  messagingSenderId: "220370367873",
  appId: "1:220370367873:web:2d7b1f94b4a0986fc86480"
};

// Initialize Firebase
export const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);