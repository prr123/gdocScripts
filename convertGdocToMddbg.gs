// **************************************************************************
// markdown
// ***************************************************************************

class MdObj {
  constructor () {
    this.divIdStr = '';
    this.ptTomm = 0.35277777777778;
    this.mmTopt = 2.8346456692913207;
    this.cDoc = DocumentApp.getActiveDocument();
    this.dFile = DriveApp.getFileById(this.cDoc.getId());
    this.dName = this.dFile.getName().trim();
    this.ndName = (this.dName).replace(/\s/g,'_'); // replaces space with underscore
    this.MdFilNam = this.ndName + 'dbg.md';
    this.DocId = this.cDoc.getId();
    this.MdImgFoldNam = this.ndName + '_MdImg';
    this.MdImgFoldId;
  }
}
class MdStrObj {
  constructor () {
    this.mdStr = '';
    this.tocStr = '';0
  }
}

class PHdObj {
  constructor () {
    this.prefix = '';
    this.suffix ='';
    this.tocPrefix = '';
    this.block = false;
  }

}

// main function conversion to Markdown
//

function ConvertGdocToMd() {
  //
  // initialize object and create output file and folders
  //
  md = new(MdObj);
  var imgFolder;
  console.log('doc name: ' + md.dName);

  var pfolders = md.dFile.getParents();

  console.log("md file: " + md.MdFilNam + '/' + md.MdImgFoldNam);
 
  var imgFolders = DriveApp.getFoldersByName(md.MdImgFoldNam);

  if (imgFolders.hasNext() == false){
    imgFolder = pfolders.next().createFolder(md.MdImgFoldNam);  
  } else {
    // delete img files
    imgFolder = imgFolders.next();
    var imgfiles = imgFolder.getFiles();
    while (imgfiles.hasNext()) {
      const imgfile = imgfiles.next();
      imgfile.setTrashed(true);
    }
  }
  md.MdImgFoldId = imgFolder.getId();

  // delete existing md file
  var mdFiles = DriveApp.getFilesByName(md.MdFilNam);
  while (mdFiles.hasNext()) {
    const mdFile = mdFiles.next();
    mdFile.setTrashed(true); 
    console.log("removed existing md file!")
  }

var mdContent = cvtToMd(); 

 // save content
  newMdFile = DriveApp.createFile(md.MdFilNam,mdContent);
   if (pfolders.hasNext()) {
    var parent = pfolders.next();
    newMdFile.moveTo(parent)
  }
  console.log('success!'); 
}



function cvtToMd() {

  var tdbg = true;
  var cDoc = md.cDoc
  var outstr = '[//]: * (Document Title: ' + cDoc.getName() + ')\n';
  outstr += '[//]: * (Document Id: ' + md.DocId + ')\n'; 
//  outstr += '[//]: * (Revision Id:' + cDoc.RevisionId + ')\r';
  var tocstr = '<p style="font-size:20pt; text-align:center">Table of Contents</p>\r';

  var body = cDoc.getBody();

  var numBodyEl = body.getNumChildren();
  for (var el=0; el<numBodyEl; el++) {
    var bodyEl = body.getChild(el);
    var bodyEltyp = bodyEl.getType();
    switch (bodyEltyp) {
      case DocumentApp.ElementType.PARAGRAPH:
        var parObj = cvtParToMd(bodyEl);
        tocstr += parObj.tocStr;
        outstr += parObj.mdStr;
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        var parObj = cvtListToMd(bodyEl);
        tocstr += parObj.tocStr;
        outstr += parObj.mdStr;
        break;
      case DocumentApp.ElementType.TABLE:

        break;
      case DocumentApp.ElementType.TABLE_OF_CONTENTS:

        break;
      case DocumentApp.ElementType.PAGE_BREAK:

        break;
      case DocumentApp.ElementType.UNSUPPORTED:
        outstr += '[//]: * (el type unsupported) \r';
        break;
      default:

        break;
    }
  }

  if (tdbg) {
      return outstr
  }

  return tocstr + outstr;
}

function cvtParToMd(el) {
  var parObj = new(MdStrObj)
  var pHdObj = getParHeadAtt(el);
  var parTxt = cvtParText(el);
// need to deal with par attr
  parObj.mdStr = pHdObj.prefix + parTxt + pHdObj.suffix;
  parObj.tocStr = pHdObj.tocPrefix + parTxt + '  \r';
  return parObj
}

function cvtListToMd(el) {
  var parObj = new(MdStrObj)
  var atts = el.getAttributes();
  var nestLev = atts[DocumentApp.Attribute.NESTING_LEVEL];
//  var listId = atts[DocumentApp.Attribute.LIST_ID];
  var glyph = el.getGlyphType();
  var glObj = CvtGlyphToObj(glyph);
  var pHdObj = getParHeadAtt(el);  
  var parTxt = cvtParText(el);
// need to deal with par attr
  var liststr = '';
  for (i=0; i<nestLev; i++) {
    liststr += '   ';
  }
  if (!glObj.glOrd) {
    liststr += '+ ';
  }
  if (pHdObj.prefix == '\r') {
    parObj.mdStr = liststr + parTxt + pHdObj.suffix;
  } else {
    parObj.mdStr = liststr + pHdObj.prefix + parTxt + pHdObj.suffix;
  }
  parObj.tocStr = pHdObj.tocPrefix + parTxt + '  \r';
  return parObj
}

function getParHeadAtt(el) {
  var par_head = el.getHeading();
  var par_att = el.getAttributes();
  var parHead = new(PHdObj);
  var hd_atts
  var aSec = md.cDoc.getActiveSection();
  var tit_head = false;

  switch (par_head) {
    case DocumentApp.ParagraphHeading.TITLE:
      parHead.prefix = '<p style="font-size:20pt; text-align:center;'
      parHead.suffix = '</p>\r\r';
      parHead.block = true;
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.TITLE);
      if (par_att.BOLD == true){
        parHead.prefix += 'font-weight:bold;'
      } else {
        if (hd_atts.BOLD) {
          parHead.prefix += 'font-weight:bold;'
        }
      }
      if (par_att.ITALIC){
        parHead.prefix += 'font-style:italic;'
      } else {
        if (hd_atts.ITALIC){
          parHead.prefix += 'font-style:italic;'
        }
      }
      parHead.prefix +='">';
      tit_head = true;
      break;
    case DocumentApp.ParagraphHeading.SUBTITLE:
      parHead.prefix = '<p style="font-size:18pt; text-align:center'
      parHead.suffix = '</p>\r\r';
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.SUBTITLE);
      parHead.block = true;
            if (par_att.BOLD == true){
        parHead.prefix += 'font-weight:bold;'
      } else {
        if (hd_atts.BOLD) {
          parHead.prefix += 'font-weight:bold;'
        }
      }
      if (par_att.ITALIC){
        parHead.prefix += 'font-style:italic;'
      } else {
        if (hd_atts.ITALIC){
          parHead.prefix += 'font-style:italic;'
        }
      }
      parHead.prefix +='">';
      tit_head = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING1:
      parHead.prefix ='# ';
      parHead.suffix = '  \r\r';
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1);
      parHead.block = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING2:
      parHead.prefix = '## ';
      parHead.suffix = '  \r\r';
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2);
      parHead.block = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING3:
      parHead.prefix = '### ';
      parHead.suffix = '  \r\r';
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3);
      parHead.block = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING4:
      parHead.prefix = '#### ';
      parHead.suffix = '  \r\r';
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING4);
      parHead.block = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING5:
      parHead.prefix = '##### ';
      parHead.suffix = '  \r\r';
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING5);
      parHead.block = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING6:
      parHead.prefix = '###### ';
      parHead.suffix = '  \r\r';
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING6);
      parHead.block = true;
      break;
    case DocumentApp.ParagraphHeading.NORMAL:
      parHead.block = false;
      parHead.prefix='\r';
      parHead.suffix = '  \r\r';
      hd_atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL);
      break;

    default:

      break;     
  }

  var attprefix = '';
  var attsuffix = '';
  if (tit_head){  
    return parHead
  }

  if (par_att.ITALIC == true){
    attprefix += '_';
    attsuffix += '_';
  } else {
    if (hd_atts.ITALIC == true) {
      attprefix += '_';
      attsuffix += '_';   
    }
  }
  if (par_att.BOLD == true) {
    attprefix = '**' + attprefix;
    attsuffix += '**';
  } else {
    if (hd_atts.BOLD == true) {
      attprefix = '**' + attprefix;
      attsuffix += '**';
    }
  }

  parHead.prefix += attprefix;
  parHead.suffix = attsuffix + parHead.suffix;
  return parHead;
}

function cvtParText(el) {

 
  var txtstr = el.getText();
  var txtidx = null;
  var numChP = el.getNumChildren();
  for (k=0; k<numChP; k++){
    elch = el.getChild(k)
    if (elch.getType() == DocumentApp.ElementType.TEXT) {
      txtidx = elch.getTextAttributeIndices();
      break;
    }    
    // possible other types
  } 
  

  var ntxtstr = '';
  italicStr = '';
  boldStr = '';
  if ((txtidx == undefined)||(txtidx == null)) {
    if (txtAtts.ITALIC) {
      italicStr = '_';
    }
    if (txtAtts.BOLD) {
      boldStr = '**';
    }
    ntxtstr = boldStr + italicStr + txtstr + italicStr + boldStr;
    return ntxtstr
  }
  
  var idxnum = txtidx.length;
  var ist = 0;
  var iend = 0;
  var txtattPrefix = '';
  var txtattSuffix = '';
  for (var i=0; i<idxnum; i++){
    if ((i+1) < idxnum) {
      iend = txtidx[i+1];
    } else {      
      iend = txtstr.length;
    }
    var txtAtts = elch.getAttributes(ist);    
    var subtxtstr = txtstr.substring(ist,iend);
    
    if (txtAtts.ITALIC) {
      subtxtstr = '_' + subtxtstr + '_';
    }
    if (txtAtts.BOLD) {
     subtxtstr = '**' + subtxtstr + '**';
    }
    ntxtstr +=subtxtstr
    ist = iend;
  }
  return ntxtstr
}

function createTocId (tstr){
  tlstr = tstr.toLowerCase();
  ttstr = tlstr.trim();
  //eliminate multiple whitespaces
  tocstr= ttstr.replace(/  +/g, '-');
  return tocstr
}
