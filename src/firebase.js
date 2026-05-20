import { initializeApp } from "firebase/app";
import { getDatabase } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyDr6wEkei9LXlXwJbwUm8JQWlJDKMTUnoE",
  authDomain: "twhgai-a31cd.firebaseapp.com",
  databaseURL: "https://twhgai-a31cd-default-rtdb.firebaseio.com",
  projectId: "twhgai-a31cd",
  storageBucket: "twhgai-a31cd.firebasestorage.app",
  messagingSenderId: "847539280749",
  appId: "1:847539280749:web:353c566fa7f0491b335864"
};

const app = initializeApp(firebaseConfig);
export const db = getDatabase(app);
