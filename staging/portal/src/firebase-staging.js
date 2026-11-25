import {initializeApp} from 'firebase/app';
import {getAuth} from 'firebase/auth';
import {getFirestore} from 'firebase/firestore';
const apiKey=import.meta.env.VITE_FIREBASE_API_KEY;
if (!apiKey || typeof apiKey !== 'string') throw new Error('Missing staging Firebase API key');
const app=initializeApp({
 apiKey,
 authDomain:'varga-cantieri-v2-staging.firebaseapp.com',
 projectId:'varga-cantieri-v2-staging',
 appId:'1:394791335182:web:6c31c93ac34e64c881935f',
 messagingSenderId:'394791335182'
});
export const auth=getAuth(app);
export const db=getFirestore(app);
