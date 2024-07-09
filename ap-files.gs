/*
   project name: ap_files
   description : Provides functions for maintenance on files in Google Drive.
   author      : Sidnei Leitão
   created     : 06/07/2024
   modified    : 09/07/2024 ad recursivity in the function getFilesList()
                 09/07/2024 ad parameter incSubFolders in the function getFilesList()

  Available Functions:

  getFilesList()
  - Parameters:
    - spreadSheetId: ID of the spreadsheet where file data will be recorded.
    - sheetName    : Name of the sheet where file data will be recorded.
    - folderId     : ID of the folder to be processed.
    - incSubFolders: True to also process subfolders.

  renameFiles()
  - Parameters:
    - spreadSheetId: ID of the spreadsheet containing the data of files to be renamed.
    - sheetName    : Name of the sheet containing the data of files to be renamed.
    - colFileId    : Index of the column containing the IDs of files to be renamed.
    - colFileName  : Index of the column containing the new names of the files.

*/

//-----------------------------------------------------------------------
function renameFiles(spreadSheetId,sheetName,colFileId = 0,colFileName=1)
{
  const ssTemp = SpreadsheetApp.openById(spreadSheetId);
  const sheetList = ssTemp.getSheetByName(sheetName);
  const dataFiles = sheetList.getRange(2,1,sheetList.getMaxRows(),2).getValues();
  const dataFilter = dataFiles.filter((row)=>row[colFileName] != "");

  dataFilter.forEach((row)=>{
    renameFile(row[colFileId],row[colFileName]);
  });  
}

//------------------------------------
const renameFile = (fileId,fileName)=>
{
  const file = DriveApp.getFileById(fileId);
  file.setName(fileName); 
}

//-------------------------------------------------------------------------
function getFilesList(spreadSheetId,sheetName,folderId,incSubFolders=false)
{

  const sheetList = SpreadsheetApp.openById(spreadSheetId).getSheetByName(sheetName);

  let folder = DriveApp.getFolderById(folderId);

  listFilesInFolder(sheetList,folder,incSubFolders);

}

//--------------------------------------------------------
function listFilesInFolder(sheetList,folder,incSubFolders)
{
  const files = folder.getFiles();    
  saveFilesInSheet(sheetList,files,folder);
  
  if(incSubFolders)
  {
    const subFolders = folder.getFolders();
    while(subFolders.hasNext())
    {
      const subFolder = subFolders.next();
      listFilesInFolder(sheetList,subFolder);
    }
  }
}

//----------------------------------------------------
function saveFilesInSheet(sheetFiles,filesList,folder)
{
  let dataFiles = [];
  while(filesList.hasNext()) {
    let file = filesList.next();
    dataFiles.push([file.getName(),file.getUrl(),file.getMimeType(),file.getId(),folder.getName()]);  
  }
  sheetFiles.getRange(2,1,dataFiles.length,dataFiles[0].length).setValues(dataFiles);
}
