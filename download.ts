import fs from 'fs';
import https from 'https';

const url = 'https://raw.githubusercontent.com/AI-technician/ware2/6928bd4dcc61a60e72c074fbf2e544ea1b278249/public/logistics.xlsx';

https.get(url, (res) => {
  const path = 'public/logistics.xlsx';
  const filePath = fs.createWriteStream(path);
  res.pipe(filePath);
  filePath.on('finish',() => {
      filePath.close();
      console.log('Download Completed'); 
  })
});
