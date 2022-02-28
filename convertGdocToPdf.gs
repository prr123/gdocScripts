// google apps script
// author: prr
// created 28/2/2022
// copyright license: MIT
// 

function convertGdocToPdf() {
  var content,fileName,newFile;//Declare variable names
  const currentDoc = DocumentApp.getActiveDocument();
  const docFile = DriveApp.getFileById(currentDoc.getId())
  const docName = docFile.getName()
//  console.log('test: ' + docName) 

  var docblob = DocumentApp.getActiveDocument().getAs('application/pdf');
    /* Add the PDF extension */
  docblob.setName(docName + ".pdf");

  var folders = docFile.getParents();
 
  newFile = DriveApp.createFile(docblob);//Create a new text file in the root folder
  if (folders.hasNext()) {
    var parent = folders.next();
    newFile.moveTo(parent)
  }
}
