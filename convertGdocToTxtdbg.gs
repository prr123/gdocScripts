function convertGdocToTxtdbg() {
  var content,fileName,newFile;//Declare variable names
  const currentDoc = DocumentApp.getActiveDocument();
  const docFile = DriveApp.getFileById(currentDoc.getId());
  const docName = docFile.getName().trim();
  console.log('test: ' + docName) 

  var ndocName = docName.replace(/\s/g,'_');
 
  var folders = docFile.getParents();
   if (folders.hasNext()) {
    var parent = folders.next();    
  }
  var content = 'File: ' + ndocName;
  if (parent !== undefined) {
    content = content + ' folder: ' + parent.getName();
  }
  content += '\r';

  var lang = currentDoc.getLanguage();
  content += 'language: ' + lang + '\r';

  var height = currentDoc.getBody().getPageHeight();
  var width = currentDoc.getBody().getPageWidth();
  content += ' Page Height: ' + height.toFixed(2) + '(pt) ' + (height*ptTomm).toFixed(2) + '(mm)';
  content += ' Page Width: ' + width.toFixed(2) +'(pt) ' + (width*ptTomm).toFixed(2) + '(mm)' +'\r';

  var marTop = currentDoc.getBody().getMarginTop();
  var marBot = currentDoc.getBody().getMarginBottom(); 
  var marR = currentDoc.getBody().getMarginRight();
  var marL = currentDoc.getBody().getMarginLeft();
  content += '  margins (TBRL mm): ' + (marTop*ptTomm).toFixed(2) + '|' + (marBot*ptTomm).toFixed(2) + '|'+ (marR*ptTomm).toFixed(2) + '|' + (marL*ptTomm).toFixed(2)+'\r\r';  


  content += DebugGdoc(ndocName,parent);
  var txtFileName = ndocName + 'Dbg.txt';//Create a new file name with date on end
  console.log("file: ", txtFileName)

  var txtFiles = DriveApp.getFilesByName(txtFileName);
  while (txtFiles.hasNext()) {
    const txtFile = txtFiles.next();
    txtFile.setTrashed(true); 
    console.log("removed existing txt file!")
  }
  newFile = DriveApp.createFile(txtFileName,content);//Create a new text file in the root folder
  if (parent !== undefined){
    newFile.moveTo(parent)
  }
};

//*****************************************************************************************
//*****************************************************************************************
// debug is a text version of the file and its attributes
//
// 3/1/2022 added heading 5 and heading 6 heading attributes 
//
//

function DebugGdoc(ndocNam, folder) {
  var cDoc = DocumentApp.getActiveDocument();
  var aSec = cDoc.getActiveSection();
  var nCh = aSec.getNumChildren();
  var body = cDoc.getBody();
  var elstr = 'Headings: \r'
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.TITLE);
  elstr += 'TTTLE \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.SUBTITLE);
  elstr += 'SUBTTTLE \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL);
  elstr += 'NORMAL \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1);
  elstr += 'Heading 1 \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2);
  elstr += 'Heading 2 \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3);
  elstr += 'Heading 3 \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING4);
  elstr += 'Heading 4 \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING5);
  elstr += 'Heading 5 \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING6);
  elstr += 'Heading 6 \r';
  for (att in atts){
    elstr = elstr + '  ' + att + ': '+ atts[att] +'\r';
  }
  elstr += '\r';

/*
  elstr = elstr + 'active section children: ' + nCh + '\r';
  for (var i= 0; i < nCh; i++) {
    var el = aSec.getChild(i);
    var numElCh = el.getNumChildren();
    var numImgs = 0;
    if (el.getType() == DocumentApp.ElementType.PARAGRAPH){
      var posImgs = el.getPositionedImages();
      numImgs = posImgs.length; 
    }
    elstr += 'el: ' + i + ' type: ' + el.getType() + ' children: '+ numElCh + ' posImages: '+ numImgs + '\r'  
//    elstr = elstr + 'children: ' + numElCh + '\r';
    for (var ch=0; ch<numElCh; ch++){
      var elCh = el.getChild(ch);
      var gchtyp = elCh.getType().name();
       //if (gchtyp !== DocumentApp.ElementType.TEXT) {
       //var gchNum = elCh.getNumChildren();
       //var gchposImgs = elch.getPositionedImages();
      elstr += '  child['+ i + ':' + ch + ']: ' + elCh.getType() + '\r';
    }
  }
*/
// body
  elstr += '*************** body *************\r'
  var imgCount = 0;
  var elheading;
  var bodyNumCh = body.getNumChildren();
  elstr += 'Number of Elements: ' + bodyNumCh +'\r\r';
  for (var idx = 0; idx< bodyNumCh; idx++){
    var child = body.getChild(idx);
    elstr += 'body child: ' + (idx +1) + ' type: ' + child.getType(); 
    switch ( child.getType() ) {
      case DocumentApp.ElementType.PARAGRAPH:
        var container = child.asParagraph();
        elheading = container.getHeading();
        var posImgList = container.getPositionedImages;
        elstr += ' Par Heading: ' + getHeadingStr(elheading)
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        var container = child.asListItem();
        elheading = container.getHeading();
        var posImgList = container.getPositionedImages;
         var nest = container.getNestingLevel();
         elstr += ' List Heading: ' + getHeadingStr(elheading);
         elstr += ' List Nest Level: ' + nest;
        break;
      case DocumentApp.ElementType.TABLE:
        var tabel = child.asTable();
        elstr += ' No Heading ';
        break;
      default:

        // Skip elements that can't contain PositionedImages.
        continue;
    }

    elstr += ' gchildren: ' + container.getNumChildren();

    // check for positioned images. Note that Table elements cannot have positioned images
    
    if (posImgList !== undefined) {
      var imgNam = ndocNam + '_Img_'
      elstr += ' Positioned Images: ' + posImgList.length +'\r';
      for (var im=0; im<posImgList.length; im++){
        var imel = posImgList[im]; 
        var layout = imel.getLayout();
        var layoutstr = getImgPos(layout);
        var blob = imel.getBlob();
        var imgId = imel.getId();
        var imgTyp = blob.getContentType();
        var imgExt = getImgExt(imgTyp);
        blob.setName(imgNam + imgCount + imgExt)
        // don't need to save image file for debug
        /*       
        newFile = DriveApp.createFile(blob);
        if (imgFolder !== undefined){
          newFile.moveTo(imgFolder)
        }
        */
        imgCount++;
        elstr += '  positioned_image: ' + im + ' type: ' + imgTyp + ' id: ' + imel.getId() + ' height: ' + imel.getHeight() + ' | ' +(imel.getHeight()*pxTomm).toFixed(1) + 'mm  width: ' + imel.getWidth() + ' | '+ (imel.getWidth()*pxTomm).toFixed(1) + 'mm\r';
        elstr += '    layout: "' + layoutstr + '" top offset: ' + (imel.getTopOffset()*ptTomm).toFixed(1) + 'mm left offset: ' + (imel.getLeftOffset()*ptTomm).toFixed(1) + 'mm\r';
      }
    } else {
      elstr += '\r';
    }
    

    let numElChild = container.getNumChildren();
    if (numElChild <1) {
      elstr += ' no grand children\r';
    }
    for (var ch=0; ch<numElChild; ch++){
      var elCh = container.getChild(ch);
      elstr += '  gchild['+ idx + ':' + ch + ']: ' + elCh.getType();
      switch (elCh.getType()) {
        case DocumentApp.ElementType.TEXT:
          var txtstr = elCh.getText();
          if (txtstr.length > 35) {
            elstr += ' [' + txtstr.length + ']: "' + txtstr.substring(0,35) + '..."';
          } else {
            elstr += ' [' + txtstr.length + ']: "' + txtstr + '"';
          }

          var txtindices = elCh.getTextAttributeIndices();
          elstr += ' text parts: ' + txtindices.length + ' ';

          break;
        case DocumentApp.ElementType.INLINE_IMAGE:
          // needs to be fixed!
          var urlstr = elCh.getLinkUrl();
          if (urlstr === null) {
            var imgBlob = elCh.getBlob();
            var imgTyp = imgBlob.getContentType();
            var imgExt = getImgExt(imgTyp);
            var imgFilNam = imgNam + imgCount + imgExt;
            imgBlob.setName(imgFilNam)
            newFile = DriveApp.createFile(imgBlob);
            if (folder !== undefined){
              newFile.moveTo(folder)
            } 
            elstr = elstr + ' img: ' + imgFilNam + ' width: ' + elCh.getWidth() + ' height ' + elCh.getHeight();          
          } else {
            elstr = elstr + ' url: ' + urlstr + ' name: ' + elCh.getId + ' width: ' + elCh.getWidth() + ' height ' + elCh.getHeight();
          }
          imgCount++; 
          break;
        case DocumentApp.ElementType.INLINE_DRAWING:
          elstr += ' inline drawing ';
          break;
        case DocumentApp.ElementType.TABLE_ROW:
          elstr += ' table row [' +ch + ']:';
          elstr += ' col [' + elCh.getNumCells() +']';
          break;
        default:
          elstr += ' unclassified sub-element ';
          break;
      }
      
      elstr += '\r';
    }     
  }
  
  elstr += '\r' +'Attributes for each element ' + '\r'; 
  for (i= 0; i < nCh; i++) {
    var el = aSec.getChild(i);
    elstr += '\r****************************************************\r'
    elstr = elstr + 'el: ' + i + ' type: ' + el.getType();
    if (el.getHeading !== undefined){
      elstr = elstr + ' Heading: ' + getHeadingStr(el.getHeading());
    }

    var numGCh = el.getNumChildren();
    if (numGCh == 0) {
      elstr += ' no grand children';
    } else {
      elstr += '  grand children: ' + numGCh;
    }

    var imgPos = el.getPositionedImages;
    if (imgPos !== undefined) {
      var numImgPos = imgPos.length;
      elstr += ' Pos Images: ' + numImgPos + '\r';
    } else {
      elstr += ' no Pos Images\r';
    }
    // this needs to be fixed
    for (j = 0; j<numGCh; j++){
        var elch = el.getChild(j)
        elstr += '  gchild: ' + (j+1) + ' child type: ' + elch.getType();
        if (elch.getType() == DocumentApp.ElementType.TEXT) {
          var txtstr = elch.getText();
          elstr += ' [' + txtstr.length + '] "' + txtstr + '"';
        }
        elstr += '\r';
    }
    
  
    switch (el.getType()) {
    case DocumentApp.ElementType.PARAGRAPH:
      elstr += procPar(el);
      if (numImgPos > 0){
        elstr += procPosImg(imgPos);
      }
      break;
    case DocumentApp.ElementType.LIST_ITEM:
      elstr += procList(el);
      if (numImgPos > 0){
        elstr += procPosImg(imgPos);
      }
      break;
    case DocumentApp.ElementType.TABLE:
      elstr += procTable(el);
    default:
      break;
    }
  }
  return elstr;
}

function procPosImg(imgPosList) {
  var dbgstr = '';
  var pImgNum = imgPosList.length;
  for (i=0; i<pImgNum; i++ ){
    var pImg = imgPosList[i];
    dbgstr += ' Positioned Image: ' + i + ' Img Id: ' + pImg.getId + '\r';
    if (pImg === undefined) {
      continue;
    } 
    var layout= pImg.getLayout();
    dbgstr += '   Layout: ' + getImgPos(layout) + '\r';

    var blob = pImg.getBlob();
    var imgTyp = blob.getContentType();
    var imgExt = getImgExt(imgTyp);
    dbgstr += ' Image Type:  ' + imgExt + '\r';
  }
}

function procList(el) {

  var glyph = el.getGlyphType();
  var txtstr = el.getText();
  var numChP = el.getNumChildren();
  var nest = el.getNestingLevel();

  var indst = el.getIndentStart();
  var indfl = el.getIndentFirstLine();
  var indend = el.getIndentEnd();
  var id = el.getListId();

  var glyphstr = getGlyphstr(glyph);
  var tlistStr = '  List Element: ' + id + ' List level: ' + nest + ' Glyph: ' + getGlyphstr(glyph) + ' grand children: '+ numChP + '\r';
  var listStr = tlistStr + '  List Text: ' + txtstr  + '\r';

  var attstr = '  List Attributes: \r';
  var atts = el.getAttributes();
  for (att in atts) {
      attstr = attstr + '    ' + att + ': ' + atts[att] + '\r';
  }
  listStr += attstr +'\r';

/*
  for (k=0; k<numChP; k++){
    elch = el.getChild(k)
    listStr += 'grand child: ' + k + ' type: '+ elch.getType() + '\r';
    if (elch.getType() == DocumentApp.ElementType.TEXT) {
      var indices = elch.getTextAttributeIndices();
      break;
    }
  }
 */   
  return listStr;
}

function procTable(tabel) {
  var tableStr = '';
  var tableCols;
  var tableRow;
  var tab = tabel;
  var tabNumRows = tab.getNumRows();
  tableStr += 'Table Number of Rows: ' + tabNumRows + '\r';
  for (irow = 0; irow< tabNumRows; irow++){
    tableRow = tab.getRow(irow);
    tableCols = tableRow.getNumCells();
    tableStr += '  ' + 'Row: ' + irow + ':' + ' Cols: ' +tableCols + ' min height: ' + tableRow.getMinimumHeight() + '\r';
  }
  tableStr += 'Table Columns\r';
  for (icol = 0; icol<tableCols; icol++) {
    tableStr += '  Col: ' + icol + ': width: ' + tab.getColumnWidth(icol)+'\r';

  }
  tableStr += '\r*** Table Content ***\r';
  for (irow = 0; irow< tabNumRows; irow++){
    tableRow = tab.getRow(irow);
    var tableColNum = tableRow.getNumCells();
    for (icol = 0; icol < tableColNum; icol++){
      var tableCell = tableRow.getCell(icol);
      var txt = tableCell.getText();
      tableStr += 'row: ' + irow + ' col: ' + icol + ' text: "' + txt + '"\r';
      var cellAtts = tableCell.getAttributes();
      for (att in cellAtts) {
         tableStr += ' att: ' + att + ' : ' + cellAtts[att] + '\r';
      }
      //var el = tableCell.getChild;
      var numChP = tableCell.getNumChildren();
      for (k=0; k<numChP; k++){
        var elch = tableCell.getChild(k);
        var elchTyp = elch.getType();
        var numGrGCh = elch.getNumChildren();
        var gchStr = 'grand child type: ' + elchTyp + ' : ' + numGrGCh + '\r';
        tableStr += gchStr;
        switch(elchTyp) {
          case DocumentApp.ElementType.PARAGRAPH:         
            var tstr = procPar(elch);
            tableStr += tstr;
            break;
          case DocumentApp.ElementType.LIST_ITEM:
            break;
          case DocumentApp.ElementType.TABLE:
            break;
          default:
            break;
        }
      }
    }     
  }

  tableStr += '*** Table Attributes ***\r';
  var tabatts = tab.getAttributes();
  for (var att in tabatts) {
    tableStr += att + ': ' + tabatts[att] + '\r';
  } 
  return tableStr;
}

function procPar(el) {
  var parstr = '';
  var attstr = '  Par Attributes: \r';
  var atts = el.getAttributes();
  for (att in atts) {
      attstr = attstr + '  ' + att + ': ' + atts[att] + '\r';
  }
  parstr += attstr +'\r';

  var txtstr = el.getText();

  parstr = parstr + '\rText (' + txtstr.length + '): ' + txtstr +'\r'
  parstr = parstr + '\rText Attributes\r'
  hozal = atts[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]
  parstr = parstr + '  horizontal alignment: ' + hozal + '\r'
  parstr = parstr + '  bold: ' + atts[DocumentApp.Attribute.BOLD] + '\r'
  parstr = parstr + '  italic: ' + atts[DocumentApp.Attribute.ITALIC] + '\r'
  parstr = parstr + '  underline: ' + atts[DocumentApp.Attribute.UNDERLINE] + '\r'

  var numChP = el.getNumChildren();
  for (k=0; k<numChP; k++){
    elch = el.getChild(k)
    if (elch.getType() == DocumentApp.ElementType.TEXT) {
      var indices = elch.getTextAttributeIndices();
      break;
    }
  } 

// indicies are not defined for every paragraph, only if there is a child TEXT
  if (indices != null) {
    parstr = parstr + '  number of text indices: '+ indices.length + '\r'
    for (idx = 0; idx < indices.length; idx++){
      var partAtts = elch.getAttributes(indices[idx]);
      var startPos = indices[idx];
      var endPos = idx+1 < indices.length ? indices[idx+1]: txtstr.length;

      var partText = txtstr.substring(startPos, endPos);
      parstr = parstr + '    index: '+ idx + ' pos: ' + startPos + ':' + endPos + ' text: "' + partText + '"\r'

      for (patt in partAtts) {
        parstr = parstr + '      par att: ' + patt + ': ' + partAtts[patt] + '\r';
      }
    }
  } 

  parstr += '*** Paragraph Children: \r'
  
  if (numChP > 0) {
    attstr = '\r Par Child Att: \r'
    for (j = 0; j<numChP; j++){
      var elch = el.getChild(j)
      parstr = parstr + 'child: ' + j +' child type: ' + elch.getType() + '\r'
      if (elch.getType() !== DocumentApp.ElementType.TEXT) {
        var patts = elch.getAttributes();
        for (att in atts) {
          attstr = attstr + '  ' + att + ': ' + atts[att] + '\r';
        }
        parstr += attstr
      }
    }
  }
 
  if (el.getHeading) {
    parstr = parstr + '    Heading: ' + el.getHeading() + '\r';
  }

  
  if (el.getTextAlignment()) {
      parstr = parstr + '    Alignment: ' + el.getTextAlignment() + '\r';  
  }

//  if (el.getHeader())
  return parstr
}


function getGlyphstr (glyph) {
  var glyphstr = '';
        
  switch (glyph) {
    case DocumentApp.GlyphType.BULLET:
      glyphstr = 'bullet';
      break;
    case DocumentApp.GlyphType.HOLLOW_BULLET:
      glyphstr = 'hollow bullet';
      break;
    case DocumentApp.GlyphType.SQUARE_BULLET:
      glyphstr = 'square bullet';
      break;        
    case DocumentApp.GlyphType.LATIN_UPPER:
      glyphstr = 'latin UP';
      break;      
    case DocumentApp.GlyphType.LATIN_LOWER:
      glyphstr = 'latin lower';
      break;    
    case DocumentApp.GlyphType.NUMBER:
      glyphstr = 'number';
      break;
    case DocumentApp.GlyphType.ROMAN_UPPER:
      glyphstr = 'roman upper';
      break;  
    case DocumentApp.GlyphType.ROMAN_LOWER:
      glyphstr = 'roman lower';
      break;
    default:

      break;
  }
  return glyphstr;
}

function getHeadingStr(headingEnum) {
  hdStr = '';
  switch (headingEnum) {
    case DocumentApp.ParagraphHeading.HEADING1:
      hdStr += 'Heading1';
      break;
    case DocumentApp.ParagraphHeading.HEADING2:
      hdStr += 'Heading2';  
      break;
    case DocumentApp.ParagraphHeading.HEADING3:
      hdStr += 'Heading3';
      break;
    case DocumentApp.ParagraphHeading.HEADING4:
      hdStr += 'Heading4';
      break;
    case DocumentApp.ParagraphHeading.HEADING5:
      hdStr += 'Heading5';
      break;
    case DocumentApp.ParagraphHeading.HEADING6:
      hdStr += 'Heading6';
      break;
    case DocumentApp.ParagraphHeading.TITLE:
      hdStr += 'Title';
      break;
    case DocumentApp.ParagraphHeading.SUBTITLE:
      hdStr += 'Subtitle';
      break;
    case DocumentApp.ParagraphHeading.NORMAL:
      hdStr += 'normal';
      break;          
    default:
      hdStr += 'undef heading';
      break;
  }
  return hdStr
}
function getImgExt(imgcont) {
  var imgExtStr ='';
  switch (imgcont){
    case 'image/png':
      imgExtStr ='.png';
      break;
    case 'image/jpg':
      imgExtStr = '.jpg';
      break;
    case 'image/jpeg':
      imgExtStr = '.jpeg';
      break;
    case 'image/bmp':
      imgExtStr = '.bmp';
      break;
    case 'image/gif':
      imgExtStr = '.gif';
      break;
  }

  return imgExtStr
}

function getImgPos(layout) {
  var laystr = '  img layout: ';
  switch (layout) {
    case DocumentApp.PositionedLayout.ABOVE_TEXT:
      laystr += 'above text';
      break;
    case DocumentApp.PositionedLayout.BREAK_BOTH:
      laystr += 'break both';
      break;
    case DocumentApp.PositionedLayout.BREAK_LEFT:
      laystr += 'break left';
      break;
    case DocumentApp.PositionedLayout.BREAK_RIGHT:
      laystr += 'break right';
      break;
    case DocumentApp.PositionedLayout.WRAP_TEXT:
      laystr += 'wrap text';
      break;
    default:
  }
  return laystr;
}

function getLayoutString( PositionedLayout ) {
  var layout;
  switch ( PositionedLayout ) {
    case DocumentApp.PositionedLayout.ABOVE_TEXT:
      layout = "ABOVE_TEXT";
      break;
    case DocumentApp.PositionedLayout.BREAK_BOTH:
      layout = "BREAK_BOTH";
      break;
    case DocumentApp.PositionedLayout.BREAK_LEFT:
      layout = "BREAK_LEFT";
      break;
    case DocumentApp.PositionedLayout.BREAK_RIGHT:
      layout = "BREAK_RIGHT";
      break;
    case DocumentApp.PositionedLayout.WRAP_TEXT:
      layout = "WRAP_TEXT";
      break;
    default:
      layout = "";
      break;
  }
  return layout;
}
