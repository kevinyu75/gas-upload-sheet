function 取得商品DocumentId(customID) {
  try {
    const email = PropertiesService.getScriptProperties().getProperty('firestore_email');
    const key = PropertiesService.getScriptProperties().getProperty('firestore_key');
    const projectId = PropertiesService.getScriptProperties().getProperty('firestore_projectId');
    
    if (!email || !key || !projectId) {
      return null;
    }
    
    const formattedKey = key.replace(/\\n/g, '\n');
    const firestore = FirestoreApp.getFirestore(email, formattedKey, projectId);

    try {
      const products = firestore.query('products')
        .Where('customID', '==', customID)
        .Execute();
      
      if (products.length > 0) {
        return products[0].name.split('/').pop();
      }
      return null;
      
    } catch (queryError) {
      return null;
    }
    
  } catch (error) {
    return null;
  }
}