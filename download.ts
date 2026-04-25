import fs from 'fs';
import https from 'https';

const url = 'https://raw.githubusercontent.com/AI-technician/ware2/477c2575e1749417fd119e74630cc9e3635e4df8/public/logistics.xlsx';

https.get(url, (res) => {
  const path = 'public/logistics.xlsx';
  const filePath = fs.createWriteStream(path);
  res.pipe(filePath);
  filePath.on('finish',() => {
      filePath.close();
      console.log('Download Completed'); 
  })
});
