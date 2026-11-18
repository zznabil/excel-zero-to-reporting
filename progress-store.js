(() => {
  'use strict';
  const name='excel-zero-to-reporting-public-v2';
  let connection=null;
  function open(){
    if(!connection)connection=new Promise((resolve,reject)=>{
      let failed=false;
      const request=indexedDB.open(name,1);
      request.onupgradeneeded=()=>request.result.createObjectStore('records');
      request.onerror=()=>{failed=true;reject(request.error);};
      request.onblocked=()=>{failed=true;reject(new Error('Progress storage is blocked by another browser connection.'));};
      request.onsuccess=()=>{const db=request.result;if(failed){db.close();return;}db.onversionchange=()=>{db.close();connection=null;};resolve(db);};
    });
    return connection;
  }
  async function transact(edit){
    const db=await open();
    return new Promise((resolve,reject)=>{
      const transaction=db.transaction('records',edit?'readwrite':'readonly',edit?{durability:'strict'}:{});
      const store=transaction.objectStore('records'),request=store.get('progress');
      let value,error;
      request.onsuccess=()=>{
        try{value=edit?edit(request.result):request.result;if(edit)store.put(value,'progress');}
        catch(cause){error=cause;transaction.abort();}
      };
      transaction.oncomplete=()=>resolve(value);
      transaction.onabort=()=>reject(error||transaction.error||new Error('Progress transaction was cancelled.'));
      transaction.onerror=()=>{error=error||transaction.error;};
    });
  }
  window.EXCEL_PROGRESS_STORE={
    read:()=>transact(),
    update:edit=>transact(edit),
    close:()=>{const old=connection;connection=null;if(old)old.then(db=>db.close(),()=>{});}
  };
})();
