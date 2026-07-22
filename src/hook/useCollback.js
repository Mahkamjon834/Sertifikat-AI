import { addDoc, collection, onSnapshot, } from 'firebase/firestore';
import { db } from '../firebase/config'

import { useEffect, useState } from 'react'

const useCollback = (collectionName) => {
    const [data, setData] = useState(null);
    useEffect(() => {
        const unsubscribe = onSnapshot(collection(db, collectionName), (querySnapshot) => {
            const data = []
            querySnapshot.forEach((doc) => {
                data.push({ id: doc.id, ...doc.data() })
            });
            setData(data)
        })
        return () => unsubscribe();
    }, [])

    return { data }
}

export default useCollback
