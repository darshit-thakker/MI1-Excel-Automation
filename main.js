// main.js
const { app, BrowserWindow, ipcMain, clipboard, screen, Notification } = require('electron');
const exceljs = require('exceljs');
const path = require('path');
const os = require('os');
 
let mainWindow;
 
function createWindow() {
  const { width: screenWidth, height: screenHeight } = screen.getPrimaryDisplay().workAreaSize;
  const winWidth = 320;
  const winHeight = 600;
  const x = screenWidth - winWidth;
  const y = (screenHeight - winHeight) / 2;
 
  mainWindow = new BrowserWindow({
    width: winWidth,
    height: winHeight,
    x: x,
    y: y,
    frame: false,
    resizable: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
    alwaysOnTop: true,
  });
 
  mainWindow.loadFile('index.html');
 
  mainWindow.on('ready-to-show', () => {
    mainWindow.show();
  });
 
 
  mainWindow.on('closed', function () {
    mainWindow = null;
  });
 
  ipcMain.on('quit-app', () => {
    app.quit();
  });
 
  ipcMain.handle('clipboard-read-text', () => {
    return clipboard.readText();
  });
 
 
  ipcMain.on('trans-text-copied', (event, {text, userID}) => {
    copyToTransID(text, userID)
      .then(() => {
        console.log('Text copied to Excel');
       
      })
      .catch((error) => {
        console.log('Error copying text to Excel:', error);
      });
  });
 
  // ipcMain.on('EC-text-copied', (event, text) => {
  //   copyToErrorCode(text)
  //     .then(() => {
  //       console.log('Text copied to Excel');
  //     })
  //     .catch((error) => {
  //       console.log('Error copying text to Excel:', error);
  //     });
  // });
 
  // ipcMain.on('ED-text-copied', (event, text) => {
  //   copyToErrorDesc(text)
  //     .then(() => {
  //       console.log('Text copied to Excel');
       
  //     })
  //     .catch((error) => {
  //       console.log('Error copying text to Excel:', error);
  //     });
  // });
 
 
  ipcMain.handle('text-copied-2', async (event, selectedOption) => {
    const username = os.userInfo().username;
    const homeDir = os.homedir(); // Get the user's home directory
    const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
    const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
    const workbook = new exceljs.Workbook();
    try {
        await workbook.xlsx.readFile(excelFilePath);
    } catch (error) {
        console.log('Error reading Excel file:', error.message);
        console.log('Creating a new workbook.');
    }
 
    const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
    let nextRow = 1;
 
    const headersExist = sheet.getCell(1, 3).value !== null;
 
    if (!headersExist) {
        const headers = ['Case Scenarios'];
        headers.forEach((header, columnIndex) => {
            const cell = sheet.getCell(nextRow, columnIndex + 3);
            cell.value = header;
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} };
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
        });
 
        nextRow++;
    } else {
        while (sheet.getCell(nextRow, 3).value) {
            nextRow++;
        }
    }
 
    const targetColumns = ['Action Taken'];
    if (startTimeRow !== null) {
      targetColumns.forEach((columnName, columnIndex) => {
        sheet.getCell(startTimeRow, columnIndex + 3).value = selectedOption;
        let maxLength=0;
        // Calculate maximum length
        maxLength = selectedOption.length > maxLength ? selectedOption.length : maxLength;
        // Set column width
        sheet.getColumn(3).width = maxLength;
        sheet.getColumn(3).alignment = { vertical: 'middle', horizontal: 'center' };
    });
  }
  else {
    console.log(`Start time is not recorded yet. "${selectedOption}" cannot be inserted.`);
    printWithNotification('Error',`Start time is not recorded yet. "${selectedOption}" cannot be inserted.`);
    // Handle the case where start time is not recorded yet
    return;
  }
 
    try {
      await workbook.xlsx.writeFile(excelFilePath);
      printWithNotification('Success', `Text "${selectedOption}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
      //console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    } catch (writeError) {
      printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
    }
  });
 
 
 
ipcMain.handle('text-copied-3', async (event, {selectedOption,quantity}) => {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
  const workbook = new exceljs.Workbook();
 
  try {
      await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
      console.log('Error reading Excel file:', error.message);
      console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 12).value !== null;
 
  if (!headersExist) {
      const headers = ['Status'];
      const headers1 = ['Count of Tracking IDs'];
      headers.forEach((header, columnIndex) => {
          const cell = sheet.getCell(nextRow, columnIndex + 12);
          cell.value = header;
          cell.font = { bold: true };
          cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
 
      });
      let maxLength=0;
      headers1.forEach((header, columnIndex) => {
          const cell = sheet.getCell(nextRow, columnIndex + 5);
          cell.value = header;
          cell.font = { bold: true };
          cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} }; // Red background color
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
          // Calculate maximum length
          maxLength = header.length > maxLength ? header.length : maxLength;
      });
      // Set column width
      sheet.getColumn(5).width = maxLength;
      nextRow++;
  } else {
      while (sheet.getCell(nextRow, 12).value) {
          nextRow++;
      }
  }
 
  const targetColumns = ['Status'];
  if (startTimeRow !== null) {
    targetColumns.forEach((columnName, columnIndex) => {
        sheet.getCell(startTimeRow, columnIndex + 12).value = selectedOption;
        sheet.getCell(startTimeRow, columnIndex + 5).value = Number(quantity);
        sheet.getColumn(12).alignment = { vertical: 'middle', horizontal: 'center' };
        sheet.getColumn(5).alignment = { vertical: 'middle', horizontal: 'center' };
 
        // let maxLength=0;
        // let maxLength1=0;
        // // Calculate maximum length
        // maxLength = selectedOption.length > maxLength ? selectedOption.length : maxLength;
        // maxLength1 = quantity.length > maxLength1 ? quantity.length : maxLength1;
        // // Set column width
        // sheet.getColumn(12).width = maxLength;
        // sheet.getColumn(5).width = maxLength1;
  });
}
else {
  console.log(`Start time is not recorded yet. "${selectedOption}" cannot be inserted.`);
  printWithNotification('Error',`Start time is not recorded yet. "${selectedOption}" cannot be inserted.`);
  // Handle the case where start time is not recorded yet
  return;
}
 
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${selectedOption}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns} for AWB Quantity: ${quantity}`);
   // console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
});
 
ipcMain.handle('insert-remark', async (event, {remark}) => {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'test', excelFileName);
  const workbook = new exceljs.Workbook();
 
  try {
      await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
      console.log('Error reading Excel file:', error.message);
      console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 13).value !== null;
 
  if (!headersExist) {
      const headers = ['Remarks'];
      headers.forEach((header, columnIndex) => {
          const cell = sheet.getCell(nextRow, columnIndex + 13);
          cell.value = header;
          cell.font = { bold: true };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
      });
      nextRow++;
  } else {
      while (sheet.getCell(nextRow, 13).value) {
          nextRow++;
      }
  }
 
  const targetColumns = ['Remarks'];
  if (startTimeRow !== null) {
    targetColumns.forEach((columnName, columnIndex) => {
      sheet.getCell(startTimeRow, columnIndex + 13).value = remark;
      let maxLength=0;
      // Calculate maximum length
      maxLength = remark.length > maxLength ? remark.length : maxLength;
      // Set column width
      sheet.getColumn(13).width = maxLength;
      sheet.getColumn(13).alignment = { vertical: 'middle', horizontal: 'center' };
  });
  } else {
      console.log(`Start time is not recorded yet. "${remark}" cannot be inserted.`);
      printWithNotification('Error',`Start time is not recorded yet. "${remark}" cannot be inserted.`);
      // Handle the case where start time is not recorded yet
      return;
    }
 
    try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${remark}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
  });
 
ipcMain.handle('record-time', () => {
  return getCurrentTimestamp();
});
 
ipcMain.handle('record-time2', async () => {
  const{currentTime2,row}=await getCurrentTimestamp2();
  //console.log(`Time "${currentTime2}" recorded to Excel at Row ..., Column: End Time`);
 
  // Calculate TAT and update the 'Available' column
  await calculateTAT(row, row);
  return currentTime2;
});
 
  // Add the following code to handle the 'text-copied' event
  ipcMain.handle('text-copied', (event, text) => {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
  const workbook = new exceljs.Workbook();
  try {
    workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
  const headersExist = sheet.getCell(1, 1).value !== null;
 
  if (!headersExist) {
    const headers = ['Text'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 1);
      cell.value = header;
      cell.font = { bold: true };
    });
 
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 1).value) {
      nextRow++;
    }
  }
 
  ipcMain.on('insert-text', async (event, rowData) => {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
  const workbook = new exceljs.Workbook();
    try {
        await workbook.xlsx.readFile(excelFilePath);
    } catch (error) {
        console.log('Error reading Excel file:', error.message);
        console.log('Creating a new workbook.');
    }
 
    const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
    // Find the next available row
    let nextRow = sheet.rowCount + 1;
 
    // Insert user input text into the specified column (e.g., 12th column)
    //sheet.getCell(nextRow, 12).value = rowData.text; // Assuming 'text' property contains user input
 
    try {
        await workbook.xlsx.writeFile(excelFilePath);
        console.log(`Text "${rowData.text}" appended to Excel at Row ${nextRow}, Column: 12`);
        return { success: true, message: `Text "${rowData.text}" appended to Excel at Row ${nextRow}, Column: 12`};
    } catch (writeError) {
        console.error('Error writing to Excel file:', writeError.message);
        return { success: false, message: 'Error writing to Excel file: ' + writeError.message };
    }
});
 
  const targetColumns = ['Text'];
  targetColumns.forEach((columnName, columnIndex) => {
    sheet.getCell(nextRow, columnIndex + 1).value = text;
    sheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'center' };
  });
 
  workbook.xlsx.writeFile(excelFilePath);
});
 
  const isAlwaysOnTop = mainWindow.isAlwaysOnTop();
  console.log('Is window always on top?', isAlwaysOnTop);
}
 
app.on('ready', createWindow);
 
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
 
app.on('activate', function () {
  if (mainWindow === null) {
    createWindow();
  }
});
 
function printWithNotification(title, message) {
  const notification = new Notification({
    title: title,
    body: message
  });
  notification.show();
  setTimeout(() => {
    notification.close();
  }, 1500);
}
 
 
 
// async function copyToAwb(text) {
//   const username = os.userInfo().username;
//   const homeDir = os.homedir(); // Get the user's home directory
//   const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
//   const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
//   const workbook = new exceljs.Workbook();
//   try {
//     await workbook.xlsx.readFile(excelFilePath);
//   } catch (error) {
//     console.log('Creating a new workbook.');
//   }
 
//   const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
//   let nextRow = 1;
 
//   const headersExist = sheet.getCell(1, 1).value !== null;
 
//   if (!headersExist) {
//     const headers = ['Sr. No.', 'EDI No.'];
//       headers.forEach((header, columnIndex) => {
//       const cell = sheet.getCell(nextRow, columnIndex + 1);
//       cell.value = header;
//       cell.font = { bold: true };
//     });
//     nextRow++;
//   } else {
//     while (sheet.getCell(nextRow, 2).value) {
//       nextRow++;
//     }
//   }
 
//   const serialNumber = nextRow - 1;
//   // Insert serial number
//   sheet.getCell(nextRow, 1).value = serialNumber;
 
//   const targetColumns = ['EDI No.'];
//   targetColumns.forEach((columnName, columnIndex) => {
//     sheet.getCell(nextRow, columnIndex + 2).value = text;
//   });
 
//   try {
//     await workbook.xlsx.writeFile(excelFilePath);
//     printWithNotification('Success', `Text "${text}" pasted to Excel at Row ${nextRow}, Column: ${targetColumns}`);
//   } catch (writeError) {
//     printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
//   }
// }
// Similarly, modify other copyTo* functions to include Sr. No.
 
async function copyToTransID(text, userID) {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
  const workbook = new exceljs.Workbook();
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Creating a new workbook.');
    printWithNotification('Error', 'Error reading Excel file: ' + error.message);
    printWithNotification('Info', 'Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 3).value !== null;
 
  if (!headersExist) {
    const headers = ['ODC Case #'];
    const headers1 = ['Processed By'];
    const headers2 = ['Sr. No.'];
 
    let maxLength=0;
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 4);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      maxLength = header.length > maxLength ? header.length : maxLength;
    });
    // Set column width
    sheet.getColumn(4).width = maxLength;
 
 
    //let maxLength1=0;
    headers1.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 11);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      //maxLength1 = header.length > maxLength1 ? header.length : maxLength1;
    });
    //sheet.getColumn(11).width = maxLength;
 
 
    headers2.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 1);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FCE4D6'} }; // Red background color
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 4).value) {
      nextRow++;
    }
  }
 
  const serialNumber = nextRow - 1;
 
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
  sheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'center' };
 
  const targetColumns = ['Tracking ID'];
  const targetColumns1 = ['UserID'];
  // Check if startTimeRow is defined and insert the tracking ID into the same row
  if (startTimeRow !== null) {
    targetColumns.forEach((columnName, columnIndex) => {
      sheet.getCell(startTimeRow, columnIndex + 4).value = text;
      sheet.getCell(startTimeRow, columnIndex + 11).value = userID;
      //let maxLength=0;
        let maxLength1=0;
        // Calculate maximum length
        //maxLength = text.length > maxLength ? text.length : maxLength;
        maxLength1 = userID.length > maxLength1 ? userID.length : maxLength1;
        // Set column width
        //sheet.getColumn(4).width = maxLength;
        sheet.getColumn(11).width = maxLength1;
        sheet.getColumn(4).alignment = { vertical: 'middle', horizontal: 'center' };
        sheet.getColumn(11).alignment = { vertical: 'middle', horizontal: 'center' };
    });
  } else {
    console.log("Start time is not recorded yet. Tracking ID cannot be inserted.");
    printWithNotification('Error',"Start time is not recorded yet. Tracking ID cannot be inserted.");
    // Handle the case where start time is not recorded yet
    return;
  }
   try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
}
 
// Similarly, modify copyToErrorCode and copyToErrorDesc functions
 
async function copyToErrorCode(text) {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
  const workbook = new exceljs.Workbook();
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 5).value !== null;
 
  if (!headersExist) {
    const headers = ['Count of Tracking IDs'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 5);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} }; // Red background color
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
 
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 5).value) {
      nextRow++;
    }
  }
 
  const serialNumber = nextRow - 1;
 
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
 
  const targetColumns = ['Case Scenarios'];
  if (startTimeRow !== null) {
    targetColumns.forEach((columnName, columnIndex) => {
      sheet.getCell(startTimeRow, columnIndex + 5).value = text;
      let maxLength=0;
      // Calculate maximum length
      maxLength = text.length > maxLength ? text.length : maxLength;
    // Set column width
      sheet.getColumn(5).width = maxLength;
      sheet.getColumn(5).alignment = { vertical: 'middle', horizontal: 'center' };
    });
 
  } else {
    console.log("Start time is not recorded yet. Error Code cannot be inserted.");
    printWithNotification('Error',"Start time is not recorded yet. Error Code cannot be inserted.");
    // Handle the case where start time is not recorded yet
    return;
  }
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
}
 
let startTimeRow = null; // Variable to store the row where start time is inserted
 
async function getCurrentTimestamp() {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
  const workbook = new exceljs.Workbook();
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Error reading Excel file:', error.message);
    console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 8).value !== null;
 
  if (!headersExist) {
    const headers = ['Start Time'];
    const headers1 = ['Assigned Date'];
    const headers2 = ['Processed By'];
    const headers3 = ['Ageing','TAT','SLA'];
    const headers4 = ['Process Area'];
    const headers5 = ['# of Disputes raised','Dispute #','Dispute Requeued','Case Requeued','Case Forwarded','Was DNF available?'];
    let maxLength1=0;
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 8);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} }; // Red background color
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      maxLength1 = header.length > maxLength1 ? header.length : maxLength1;
    });
    sheet.getColumn(8).width = maxLength1;

    let maxLength=0;
    headers1.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 6);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} }; // Red background color
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      // Calculate maximum length
    maxLength = header.length > maxLength ? header.length : maxLength;
    });
    // Set column width
    sheet.getColumn(6).width = maxLength;

    headers2.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 11);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} }; // Red background color
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    headers3.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 14);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FCE4D6'} }; 
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    headers4.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 2);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FCE4D6'} }; 
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    let maxLength2=0
    headers5.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 17);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'9BC2E6'} }; 
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      maxLength2 = header.length > maxLength2 ? header.length : maxLength2;
    });
    sheet.getColumn(17).width = maxLength2;
    sheet.getColumn(17).alignment = { vertical: 'middle', horizontal: 'center' };

    sheet.getColumn(18).width = maxLength2;
    sheet.getColumn(18).alignment = { vertical: 'middle', horizontal: 'center' };

    sheet.getColumn(19).width = maxLength2;
    sheet.getColumn(19).alignment = { vertical: 'middle', horizontal: 'center' };
    
    sheet.getColumn(20).width = maxLength2;
    sheet.getColumn(20).alignment = { vertical: 'middle', horizontal: 'center' };
    
    sheet.getColumn(21).width = maxLength2;
    sheet.getColumn(21).alignment = { vertical: 'middle', horizontal: 'center' };
    
    sheet.getColumn(22).width = maxLength2;
    sheet.getColumn(22).alignment = { vertical: 'middle', horizontal: 'center' };
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 8).value) {
      nextRow++;
    }
  }
  // Record the row where start time is inserted
  startTimeRow = nextRow;
 
  const serialNumber = nextRow - 1;
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
 
  // Get the current time as a JavaScript Date object
  const currentTime1 = new Date();
 
  const currentTime1IST = currentTime1.toLocaleTimeString('en-IN', { hour12: false, timeZone: 'Asia/Kolkata' });
  const currentDate = currentTime1.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    timeZone: 'Asia/Kolkata'
  });
  const targetColumns = ['Start Time'];
  targetColumns.forEach((columnName, columnIndex) => {
    // Set the time in the 'Time' column
    sheet.getCell(nextRow, columnIndex + 8).value = currentTime1IST;
    // Set the number format for the cell to display time only
    sheet.getCell(nextRow, columnIndex + 8).numFmt = 'hh:mm:ss';
    sheet.getColumn(8).alignment = { vertical: 'middle', horizontal: 'center' };
  }); 
 
  // Set the date in the 'Date' column
  sheet.getCell(nextRow, 6).value = currentDate;
  sheet.getCell(nextRow, 6).numFmt = 'dd/mm/yyyy';
  sheet.getColumn(6).alignment = { vertical: 'middle', horizontal: 'center' };

  // let maxLength=0;
  // let maxLength1=0;
  // // Calculate maximum length
  // maxLength = currentTime1IST.length > maxLength ? currentTime1IST.length : maxLength;
  // maxLength1 = currentDate.length > maxLength1 ? currentDate.length : maxLength1;
  // // Set column width
  // sheet.getColumn(8).width = maxLength;
  // sheet.getColumn(6).width = maxLength1;
 
  let maxLength=0;

  let newValue='Invoice Adjustment';
  sheet.getCell(nextRow, 2).value = newValue;

  maxLength = newValue.length > maxLength ? newValue.length : maxLength;
  sheet.getColumn(2).width = maxLength;
  sheet.getColumn(2).alignment = { vertical: 'middle', horizontal: 'center' };
 
  sheet.getCell(nextRow, 15).value = '48 hrs';
  sheet.getColumn(15).alignment = { vertical: 'middle', horizontal: 'center' };

  sheet.getCell(nextRow, 16).value = 'MET';
  sheet.getColumn(16).alignment = { vertical: 'middle', horizontal: 'center' };
 
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${currentTime1IST}" pasted to Excel at Row ${nextRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
}
 
 
async function getCurrentTimestamp2() {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
  const workbook = new exceljs.Workbook();
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Error reading Excel file:', error.message);
    console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  // Always insert data in the next row
let nextRow=1;
  const headersExist = sheet.getCell(1, 9).value !== null;
 
  if (!headersExist) {
    const headers = ['End Time'];
    const headers1 = ['Processed Date'];
    let maxLength=0;
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 9);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      maxLength = header.length > maxLength ? header.length : maxLength;
    });
      sheet.getColumn(9).width = maxLength;
 
 
 
    let maxLength1=0;
    headers1.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 7);
      cell.value = header;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFF00'} };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      maxLength1 = header.length > maxLength1 ? header.length : maxLength1;
    });
    sheet.getColumn(7).width = maxLength1;
 
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 9).value) {
      nextRow++;
    }
  }
  const serialNumber = nextRow - 1;
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
 
  // Get the current time as a JavaScript Date object
  const currentTime2 = new Date();
  const currentTime2IST = currentTime2.toLocaleTimeString('en-IN', { hour12: false, timeZone: 'Asia/Kolkata' });
  const currentDate = currentTime2.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    timeZone: 'Asia/Kolkata'
  });
 
  const targetColumns = ['End Time'];
  if (startTimeRow !== null) {
    targetColumns.forEach((columnName, columnIndex) => {
      // Set the time in the 'Time' column
      sheet.getCell(startTimeRow, columnIndex + 9).value = currentTime2IST;
      // Set the number format for the cell to display time only
      sheet.getCell(startTimeRow, columnIndex + 9).numFmt = 'hh:mm:ss';
      sheet.getColumn(9).alignment = { vertical: 'middle', horizontal: 'center' };
 
      // let maxLength=0;
      // // Calculate maximum length
      // maxLength = currentTime2IST.length > maxLength ? currentTime2IST.length : maxLength;
      // // Set column width
      // sheet.getColumn(9).width = maxLength;
     
    });
  }else {
    console.log("Start time is not recorded yet. End Time cannot be inserted.");
    printWithNotification('Error',"Start time is not recorded yet. End Time cannot be inserted.");
    // Handle the case where start time is not recorded yet
    return;
  }
 
    // Set the date in the 'Date' column
    sheet.getCell(nextRow, 7).value = currentDate;
    sheet.getCell(nextRow, 7).numFmt = 'dd/mm/yyyy';
    sheet.getColumn(7).alignment = { vertical: 'middle', horizontal: 'center' };
 
    const assignDateCell = sheet.getCell(nextRow, 6);  // adjust the row and column indices as necessary
    const processedDateCell = sheet.getCell(nextRow, 7);  // adjust the row and column indices as necessary
 
    const [day, month, year] = assignDateCell.value.split('/');
    const assignDate = new Date(`${month}/${day}/${year}`);
 
    const [day2, month2, year2] = processedDateCell.value.split('/');
    const processedDate = new Date(`${month2}/${day2}/${year2}`);
 
    // Now you can calculate the difference
    const differenceInDays = ((processedDate - assignDate) / (1000 * 60 * 60 * 24)) + 1;
 
    sheet.getCell(nextRow, 14).value = differenceInDays;
    sheet.getColumn(14).alignment = { vertical: 'middle', horizontal: 'center' };
 
  console.log('Start Time:', sheet.getCell(nextRow, 8).value);
  console.log('End Time:', sheet.getCell(nextRow, 9).value);
 
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${currentTime2IST}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
   // console.log(`Time abc "${currentTime2.toLocaleTimeString('en-US', { hour12: false })}" recorded to Excel at Row ${startTimeRow}, Column:${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
  return{currentTime2, row: nextRow};
}
 
// Helper function to format milliseconds to time (hh:mm:ss)
function formatMillisecondsToTime(milliseconds){
  const totalSeconds = Math.floor(milliseconds / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
 
  return `${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
}
// Helper function to pad single-digit numbers with a leading zero
function pad(number) {
  return number < 10 ? `0${number}` : number;
}
 
async function calculateTAT(startRow, endRow) {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI1_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)' , 'test', excelFileName);
  const workbook = new exceljs.Workbook();
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Error reading Excel file:', error.message);
    return;
  }
 
 
 const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
 // Find the next available row (assuming there is a 'Time' column in the first row)
 let nextRow = 1;
 
 const headersExist = sheet.getCell(1, 10).value !== null;
 
 if (!headersExist) {
   const headers = ['Processed Time'];
   let maxLength=0;
   headers.forEach((header, columnIndex) => {
     const cell = sheet.getCell(nextRow, columnIndex + 10);
     cell.value = header;
     cell.font = { bold: true };
     cell.fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FCE4D6'} };
     cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
    maxLength = header.length > maxLength ? header.length : maxLength;
    });
    sheet.getColumn(10).width = maxLength;
 
   nextRow++;
 } else {
   while (sheet.getCell(nextRow, 10).value) {
     nextRow++;
   }
 }
 
  // Ensure 'Start Time' and 'End Time' columns exist
  const startColumn = 8; // Column index for 'Start Time'
  const endColumn = 9;   // Column index for 'End Time'
  if (!sheet.getCell(1, startColumn).value || !sheet.getCell(1, endColumn).value) {
    console.log('Invalid Excel format. "Start Time" or "End Time" columns do not exist.');
    return;
  }
 
  const startTimeCell = sheet.getCell(startRow, startColumn);
  const endTimeCell = sheet.getCell(endRow, endColumn);
 
  let startTime = startTimeCell.text;  // Use .text instead of .value
  let endTime = endTimeCell.text;      // Use .text instead of .value
 
  // Check if both 'Start Time' and 'End Time' have valid values
  if (!startTime || !endTime) {
    console.log('Invalid start or end time format. Cannot calculate TAT.');
    return;
  }
 
  // Convert the time strings to JavaScript Date objects
  const startDate = new Date(`01/01/2000 ${startTime}`);
  const endDate = new Date(`01/01/2000 ${endTime}`);
 
  // Calculate TAT
  const tatMilliseconds = endDate.getTime() - startDate.getTime();
  const tatFormatted = formatMillisecondsToTime(tatMilliseconds);
 
  // Update the 'TAT' column
  const tatColumn = 10;  // Column index for 'TAT'
  sheet.getCell(endRow, tatColumn).value = tatFormatted;
  sheet.getColumn(10).alignment = { vertical: 'middle', horizontal: 'center' };
  // let maxLength=0;
  // // Calculate maximum length
  // maxLength = tatFormatted.length > maxLength ? tatFormatted.length : maxLength;
  // // Set column width
  // sheet.getColumn(10).width = maxLength;
 
 
  startTime="00:00:00";
  endTime="00:00:00";
 
  // Save the changes to the Excel file
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    console.log(`TAT "${tatFormatted}" calculated and updated in Excel at Row ${endRow}, Column: TAT`);
  } catch (writeError) {
    console.error('Error writing to Excel file:', writeError.message);
  }
}