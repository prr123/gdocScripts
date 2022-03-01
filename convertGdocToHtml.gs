//*******************************************************
//*******************************************************
// bugs
// 20 Dec 2021 heading 5 has no new line
// file testHeading
// 3 Mar 2022 inline image height has = instead of :
//
class listObj {
  constructor () {
    this.clas = null;
    this.listId = null;
    this.gltyp = null;
    this.count = 0;
    this.glInd = 0;
    this.txtInd = 0;
  }
}

class HtmlObj {
  constructor () {
    this.divIdStr = '';
    this.ptTomm = 0.35277777777778;
    this.mmTopt = 2.8346456692913207;
    this.cDoc = DocumentApp.getActiveDocument();
    this.dFile = DriveApp.getFileById(this.cDoc.getId());
    this.dName = this.dFile.getName().trim();
    this.ndName = (this.dName).replace(/\s/g,'_'); // replaces space with underscore
    this.HtmlFilNam = this.ndName + 'dbg.html';
    this.ImgFoldNam = this.ndName + '_Img';
    this.HtmlHead = '<!DOCTYPE html>\r<!-- file\: ' + this.HtmlFilNam + ' -->\r<head>\r';
    this.HtmlHeadBody = '</head>\r<body>\r'
    this.HtmlEnd = '</body></html>\r';
    this.imgCount = 0;
    this.tabCount =0;
    this.parCount = 0;
    this.h1Count = 0;
    this.h2Count = 0;
    this.h3Count = 0;
    this.h4Count = 0;
    this.h5Count = 0;
    this.h6Count = 0;
    this.spanCount = 0;
    this.subDivCount = 0;
    this.withTOC = true;
    this.width;
    this.htmlDivHead = '<div class="' + this.ndName + '">\r';
    this.htmlTocHead = '<div class="' + this.ndName + '" id="' + this.ndName + '_TOC">\r';
    this.htmlDivTail = '</div>\r';
    this.BodyCss = '';
    this.BodyHtml= '';
    this.TocHtml =  '';
    this.TocCss = '';
    this.Title = '';
    this.unList = [];
    this.orList = [];
    this.cNlev = 0;
    this.cListUn;
    this.indList=0;
    this.ImgFoldId;
  }
  Init() {    
    for (var i =0; i< 4; i++) {
        (this.unList).push(new(listObj));
        (this.orList).push(new(listObj));
    }
  };
} 

class posImgObj {
  constructor () {
    this.width = 0;
    this.height = 0;
    this.topOffset = 0;
    this.leftOffset =0;
    this.layout = null;
    this.imgFilnam = '';
  }
}


function convertGdocToHtmldbg() {
  
  wo = new(HtmlObj);
  wo.Init();
  var imgFolder;
  console.log('test: ' + wo.dName);

  var pfolders = wo.dFile.getParents();

  console.log("html file: " + wo.HtmlFilNam + '/' + wo.ImgFoldNam);
 
  var imgFolders = DriveApp.getFoldersByName(wo.ImgFoldNam);

  if (imgFolders.hasNext() == false){
    imgFolder = pfolders.next().createFolder(wo.ImgFoldNam);  
  } else {
    // delete img files
    imgFolder = imgFolders.next();
    var imgfiles = imgFolder.getFiles();
    while (imgfiles.hasNext()) {
      const imgfile = imgfiles.next();
      imgfile.setTrashed(true);
    }
  }
  wo.ImgFoldId = imgFolder.getId();

  // delete existing html file
  var htmlFiles = DriveApp.getFilesByName(wo.HtmlFilNam);
  while (htmlFiles.hasNext()) {
    const htmlFile = htmlFiles.next();
    htmlFile.setTrashed(true); 
    console.log("removed existing html file!")
  }
  
// body

  var cssHeadDoc = ConvertDocHeadAttToCSS();

  ConvertGDocToHtml();

  var htmlDiv = wo.htmlDivHead + wo.BodyHtml + wo.htmlDivTail;
  var htmlToc = wo.htmlTocHead + wo.TocHtml + wo.htmlDivTail;
  
  var cssDoc = '<style>\r' + cssHeadDoc + wo.BodyCss + wo.TocCss + '</style>\r';

  var htmlContent = wo.HtmlHead + cssDoc + wo.HtmlHeadBody + htmlToc + htmlDiv + wo.HtmlEnd;

// save content
  newFile = DriveApp.createFile(wo.HtmlFilNam,htmlContent);
   if (pfolders.hasNext()) {
    var parent = pfolders.next();
    newFile.moveTo(parent)
  }
  console.log('success!');
};

function ConvertDocHeadAttToCSS() {
  var cDoc = DocumentApp.getActiveDocument();
  var aSec = cDoc.getActiveSection();

  var cssStr = '';

  var classStr = '.' + wo.ndName;
  cssStr += classStr +' {\r';
  cssStr += '  margin-right: ' + (aSec.getMarginRight()*wo.ptTomm).toFixed(0) + 'mm;\r'; 
  cssStr += '  margin-left: ' + (aSec.getMarginLeft()*wo.ptTomm).toFixed(0) + 'mm;\r';
  cssStr += '  margin-top: 0;\r';
  cssStr += '  margin-bottom: 0;\r';
  wo.width = ((aSec.getPageWidth()-aSec.getMarginLeft()-aSec.getMarginRight())*wo.ptTomm).toFixed(0)
  cssStr += '  width: '+ wo.width + 'mm;\r';
  cssStr += '  border: solid red;\r';
  cssStr += '  border-width: 1px;\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';

  cssStr += classStr + '> * {\r';
  cssStr += ' margin: 0;\r}\r';

  cssStr += classStr + ':first-child {\r';
  cssStr += '  margin-top: ' + (aSec.getMarginTop()*wo.ptTomm).toFixed(0) + 'mm;\r}\r';

  cssStr += classStr + ':last-child {\r';
  cssStr += '  margin-bottom: ' + (aSec.getMarginBottom()*wo.ptTomm).toFixed(0) + 'mm;\r}\r';

  cssStr += '/* title */\r';
  cssStr += '.' +wo.ndName + '_title  {\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.TITLE);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';

  cssStr += '/* subtitle */\r';
  cssStr += '.' + wo.ndName + '_subtitle {\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.SUBTITLE);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';

  cssStr += classStr + ' h1 {\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';
  
  cssStr += classStr + ' li h1 {\r';
  cssStr += '  display: inline;' + '\r}\r';
  
  cssStr += classStr + ' h2 {\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';

  cssStr += classStr + ' li h2 {\r';
  cssStr += 'display: inline;' + '\r}\r';

  cssStr += classStr + ' h3 {\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';

  cssStr += classStr + ' li h3 {\r';
  cssStr += 'display: inline;' + '\r}\r';

  cssStr += classStr + ' h4 {\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING4);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';

  cssStr += classStr + ' li h4 {\r';
  cssStr += '  display: inline;\r}\r';

  cssStr += classStr + ' h5 {\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING5);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';

  cssStr += classStr + ' li h5 {\r';
  cssStr += '  display: inline;\r}\r';

  cssStr += classStr + ' h6 {\r';
  var atts = aSec.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING6);
  cssStr += ConvertTxtAttToCss(atts) + '}\r';

  cssStr += classStr + ' li h6 {\r';
  cssStr += '  display: inline;\r}\r';

  cssStr += classStr + ' ul,ol {\r';
  cssStr += '  margin:0;\r  padding-left:0;\r}\r';

  cssStr += classStr + ' li p {\r';
  cssStr += '  display: inline;\r  margin: 0;\r}\r';

  cssStr += classStr + ' p {\r';
  cssStr += '  margin:0;\r}\r';

  return cssStr
}

function ConvertGDocToHtml() {
  var cDoc = wo.cDoc;
  var body = cDoc.getBody();
  var bodyNumChildren = body.getNumChildren();
  var tocHtml = '';
  
  var tocCssid = '#'+wo.ndName + '_TOC ';
//  var tocCssTitleId = tocCssid + '_title';
  wo.TocCss = tocCssid + ' {\r  margin-bottom: 20px;\r  padding-top:10px;\r  padding-bottom:10px;\r}\r'; 
  wo.TocCss += tocCssid + 'h1 {\r  margin-top: 10px;\r  margin-bottom: 0px;\r}\r';
  wo.TocCss += tocCssid + 'h2 {\r  padding-left: 20px;\r  margin: 0px;\r}\r';
  wo.TocCss += tocCssid + 'h3 {\r  padding-left: 40px;\r  margin: 0px;\r}\r';
  wo.TocCss += tocCssid + 'h4 {\r  padding-left: 60px;\r  margin: 0px;\r}\r';
  wo.TocCss += tocCssid + 'h5 {\r  padding-left: 80px;\r  margin: 0px;\r}\r';
  wo.TocCss += tocCssid + 'h6 {\r  padding-left: 100px;\r  margin: 0px;\r}\r';
//  wo.TocCss += tocCssTitleId + '{\r';
//  wo.TocCss += '}\r';

  // Walk through all the child elements of the body.
  for (var el = 0; el < bodyNumChildren; el++) {
    var imgList = [];
    var cssHead = '';
    var prefix = '';
    var suffix = '';
    var subDiv = false;
    //var noNewLin = false;

    var child = body.getChild(el);
    
    switch (child.getType()) {
      case DocumentApp.ElementType.PARAGRAPH:
        // close list if there is a new line
        CloseList();
        CreateParHtml(child);
        break;

        // end case paragraph

      case DocumentApp.ElementType.LIST_ITEM:
        CreateListHtml(child); 
        break;
          
      case DocumentApp.ElementType.TABLE:
        container = child.asTable;
        var tabObj = ConvertTable(child);
        wo.BodyHtml += tabObj.htmlStr;
        wo.BodyCss += tabObj.cssStr;
        break;
      default:
      continue;
    } 
    
  } // main element loop (paragraph, listitem, table)
  CloseList();
  if ((wo.Title).length > 1) {
    tocHtml = '<p class="' + wo.dName + '_title">' + wo.Title + '</p>\r';
  }
  tocHtml += '<p class="' + wo.dName + '_title">Table of Contents</p>\r';
  wo.TocHtml = tocHtml + wo.TocHtml;
  return;
}

// paragraph
function CreateParHtml(child) {

  var imgNamBase = wo.ndName + '_Img_'; 
  
  var imgFolder = DriveApp.getFolderById(wo.ImgFoldId);
  var test = imgFolder.getName();
  // close open lists
  CloseList();

  var atts = child.getAttributes();

  var chHasTxt = false;  
  var chTxtStr = child.getText();
  if (chTxtStr.length > 0) {
    chHasTxt = true;
  }

  var par_head = child.getHeading();
  if (par_head == DocumentApp.ParagraphHeading.TITLE) {
    wo.Title = chTxtStr;
  }

        // create a function 
  var container = child.asParagraph;
  var numGChild = child.getNumChildren();        

  var posImgList = child.getPositionedImages();
  var imgPNum = posImgList.length;
  var subDiv = false;
  if (imgPNum > 0) {
    subDiv = true;
    wo.subDivCount++;
  }
  

  var headObj =  ConvertHeading(par_head, true);   

  var cssHead = headObj.cssStr;
  var prefix = headObj.htmlPrefix;
  var suffix = headObj.htmlSuffix;
     
  if (chHasTxt) {
          //bodyObj.BodyCss += CvtParAttToCSS(atts, cssHead);
          //pstylstr + ' {\r' + parAttstr + '}\r'
    var cssAttStr = ConvertTxtAttToCss(atts);
    if ((cssAttStr.length > 0)&&(cssHead.length>0)) {
      wo.BodyCss += cssHead + cssAttStr + '}\r';
    }
          
        // if there are positioned images we need to create a new div which contain paragraphs or headings
        // and the postioned images
        // otherwise simple paragraphs of heading are enough

    var elObj = CvtPar(child, headObj.idStr);
  }
          
  var subDivParHtmlStr = '';
  var subDivParStylStr = ''; 
  var subDivMainCss = '';
  var subDivSubCss = '';
  var subDivSubCssafter = '';

    // no positioned Images
  if (!subDiv) {            
    if (chHasTxt) {
      wo.BodyHtml += prefix + elObj.htmlStr + suffix +'\r';
      if (headObj.htmlTocPrefix.length > 1) {
        wo.TocHtml += headObj.htmlTocPrefix + elObj.htmlStr + headObj.htmlTocSuffix + '\r';
      }
      wo.BodyCss += elObj.cssStr;  
    } else {
      wo.BodyHtml += '<br>\r';
    }

  } else {
    // case of positioned images

    var subDivMainClassStr = wo.ndName + '_sub_' + wo.subDivCount.toString();
     //var subDivSubClassStr = subDivMainClassStr + '_img';
     //div change id to class
    subDivMainCss = '.' + subDivMainClassStr +' {\r';
    //subDivSubCss = '.' + subDivSubClassStr + '{\r';
    subDivSubCssafter = '.' + subDivMainClassStr + '::after {/r';
    subDivSubCssafter += 'content: ""\rclear: both;\r';
    subDivSubCssafter += '}/r';
    
    
    subDivMainCss += '  border: solid black;\r';
    subDivMainCss += '  border-width: 1px;\r';
    subDivMainCss += '}\r';

   //  subDivSubCss += '  border: solid green;\r';
   //  subDivSubCss += '  border-width: 1px;\r';
   //  subDivSubCss += '}\r';
    
    if (chHasTxt) {            
      subDivParHtmlStr += '  ' + prefix + elObj.htmlStr + suffix +'\r';
      subDivParStylStr += '.' + subDivMainClassStr +' p {\r';
      subDivParStylStr += elObj.cssStr;
       // subDivParStylStr +='position: relative;\r';
      subDivParStylStr += '}\r';
    } else {
      subDivParHtmlStr += '  <br>\r';
    }
    var subDivImgHtmlStr = '';
    var subDivImgStylStr = subDivMainCss;
    for (var im=0; im<imgPNum; im++){
      imgel = posImgList[im];
      if (imgel === undefined) { 
        break; 
      }
      var imgLayout = imgel.getLayout();
      var imgblob = imgel.getBlob();
      var imgId = imgel.getId();
      var mimgId = 'Img_' + modId(imgId);
      var imgTyp = imgblob.getContentType();
      var imgExt = getImgExt(imgTyp);
      var imgFilNam = imgNamBase + wo.imgCount + imgExt;
      imgblob.setName(imgFilNam)       
      newFile = DriveApp.createFile(imgblob);

      newFile.moveTo(imgFolder);

      wo.imgCount++;
      //var imgDispId = subDivMainClassStr + '_img_' + (im +1);
      subDivImgHtmlStr += '<!-- positioned Image id: ' + imgId + ' Top: ' + imgel.getTopOffset() + ' Left: ' + imgel.getLeftOffset() + ' Layout: ' + getImgPos(imgel.getLayout()) + '-->\r';
      subDivImgHtmlStr += '  <img src="' + wo.ImgFoldNam  + '/' + imgFilNam + '" id = "' + mimgId + '">\r';

      subDivImgStylStr += '#' + mimgId + ' {\r';
//      subDivImgStylStr += 'position: absolute;\r';
//      subDivImgStylStr += 'top: ' + imgel.getTopOffset() + 'pt; left: ' + imgel.getLeftOffset() + 'pt;\r'
      switch (imgLayout) {
        case DocumentApp.PositionedLayout.WRAP_TEXT:
            // right or left
          subDivMainCss += ' display: block;\r';
          if (imgel.getLeftOffset()*wo.ptTomm > (wo.width/2)) {
            subDivImgStylStr  += ' float: right;\r';
          } else {
             subDivImgStylStr  += ' float: left;\r';
          }
          break;
        case DocumentApp.PositionedLayout.BREAK_RIGHT:
          subDivMainCss += '  display: flex;\r';
          subDivMainCss += '  justify-content: flex-start;\r';
          break;
        case DocumentApp.PositionedLayout.BREAK_LEFT:
          subDivMainCss += '  display: flex;\r';
          subDivMainCss += '  justify-content: flex-end;\r';
          break;
        default:
          if (wo.dbg) {
            console.log('error cannot process img layout: ' + imgLayout)
          }
          break;
      }

      subDivImgStylStr += '  width: ' + imgel.getWidth() + 'px;\r';
      subDivImgStylStr += '  height: ' + imgel.getHeight() + 'px;\r';
      subDivImgStylStr += '  padding: 10px;\r';
      subDivImgStylStr += '}\r';
    } // end pos img loop 

    wo.BodyHtml += '  <div class ="' + subDivMainClassStr+ '">\r';
    wo.BodyHtml  += subDivImgHtmlStr + subDivParHtmlStr +'  </div>\r';

    wo.BodyCss += subDivParStylStr + subDivImgStylStr;
  }
    // other sub elements of paragraph
    // loop through grand children
  for (var k = 0; k< numGChild; k++) {
    var gchild = child.getChild(k);
    var gchildtyp = gchild.getType();
      // I am assuming paragraphs do not contain other paragraphs, lists or tables for the time being 
    switch(gchildtyp) {
      case DocumentApp.ElementType.INLINE_IMAGE:        
        var urlstr = gchild.getLinkUrl();
        var imgId = modId(gchild.getId());
        wo.BodyCss += '#' + imgId + ' {/r';
        if (urlstr === null) {
          var imgBlob = gchild.getBlob();
          var imgTyp = imgBlob.getContentType();
          var imgExt = getImgExt(imgTyp);
          var imgFilNam = imgNamBase + wo.imgCount + imgExt;
          imgBlob.setName(imgFilNam)
          newFile = DriveApp.createFile(imgBlob);

          newFile.moveTo(imgFolder)
          wo.BodyHtml += '<!-- inline image: ' + imgFilNam + ' Id: '+ imgId +' -->\r';
          wo.BodyHtml += '<img src="' + wo.ImgFoldNam + '/' + imgFilNam + '" id="' + imgId +'">\r<br>\r'; 
        } else {  
          wo.BodyHtml += '<img src="' + urlstr + '" id="' + imgId +'"> >\r<br>\r';
        }
        wo.BodyCss += '  width:' + gchild.getWidth() + '\r  height:' + gchild.getHeight() + '\r';
        wo.BodyCss += '}/r;'
        wo.imgCount++;
        break;       
      // todo
      case DocumentApp.ElementType.INLINE_DRAWING:
              container = child.asInlineDrawing;
        break;
      case DocumentApp.ElementType.TEXT:
        break;
      case DocumentApp.ElementType.PAGE_BREAK:
        wo.BodyHtml += '<!-- page break -->\r';
        break;
      default:
        wo.BodyHtml += '<!-- gchild: ' + k + ' type: ' + gchildtype + ' -->\r';
        break;
    }
  } //end loop of grand children
  return
}

// list item
function CreateListHtml(child) {

  var atts = child.getAttributes();
  var nestLev = atts[DocumentApp.Attribute.NESTING_LEVEL];
  var listId = atts[DocumentApp.Attribute.LIST_ID];
  var glyph = child.getGlyphType();
  var glObj = CvtGlyphToObj(glyph);
  var imgNam = wo.ndName + '_Img_';
  var imgFolder = DriveApp.getFolderById(wo.ImgFoldId);

  function NewUnList() {
    // new ul
    var ind0 = atts[DocumentApp.Attribute.INDENT_START];
    var ind1 = atts[DocumentApp.Attribute.INDENT_FIRST_LINE];
    var listId = atts[DocumentApp.Attribute.LIST_ID];
    var nestLev = atts[DocumentApp.Attribute.NESTING_LEVEL];
    wo.unList[nestLev].listId = listId;
    wo.unList[nestLev].glInd = ind1;
    wo.unList[nestLev].txtInd = ind0;
    wo.cNlev = nestLev;
    var nListId = 'u' + nestLev + modId(listId);
    wo.BodyHtml +='<ul class="'+ nListId +'">\r';  
    wo.BodyHtml +='  <li>';        
    wo.BodyCss += '.' + wo.ndName + ' ul.' + nListId + ' {\r  list-style-type:' + glObj.glTyp + ';\r'
    wo.BodyCss += '  list-style-position: inside;\r'
    wo.BodyCss += '  padding-left: 0;\r}\r';
    wo.BodyCss += '.' + wo.ndName + ' ul.' + nListId + ' li {\r';
    wo.BodyCss += '  padding-top: 2pt;\r';
    wo.BodyCss += '  padding-bottom: 2pt;\r';
    wo.BodyCss += '  padding-left:' + (ind1) +'pt;\r}\r';
    wo.BodyCss += '.' + wo.ndName + ' ul.' + nListId + ' li > * {\r';
    wo.BodyCss += '  padding-left:' + (ind0 - ind1 -14) +'pt;\r}\r';
    return
  }

  function NewOrList() {
  
    var ind0 = atts[DocumentApp.Attribute.INDENT_START];
    var ind1 = atts[DocumentApp.Attribute.INDENT_FIRST_LINE];
    var listId = atts[DocumentApp.Attribute.LIST_ID];
      //bug
     //var nestLev = atts[DocumentApp.Attribute.NESTING_LEVEL];
     // need to change if two different lists start after each other without nesting difference
    if (wo.indList == 0) {
        nestLev = 0;
    } else {      
      nestLev = wo.cNlev + 1;
    }
    wo.orList[nestLev].listId = listId;
    wo.orList[nestLev].glInd = ind1;
    wo.orList[nestLev].txtInd = ind0;
    wo.cNlev = nestLev;
    wo.indList = ind0;
    var nListId = 'o' + nestLev + modId(listId);
    wo.BodyHtml +='<ol class="'+ nListId +'">\r';  
    wo.BodyHtml +='  <li>';     
    wo.BodyCss += '.' + wo.ndName + ' ol.' + nListId + ' {\r  list-style-type:' + glObj.glTyp + ';\r'
    wo.BodyCss += '  list-style-position: inside;\r'
    wo.BodyCss += '  padding-left: 0;\r}\r';
    wo.BodyCss += '.' + wo.ndName + ' ol.' + nListId + ' li {\r';
    wo.BodyCss += '  padding-top: 2pt;\r';
    wo.BodyCss += '  padding-bottom: 2pt;\r';
    wo.BodyCss += '  padding-left:' + (ind1) +'pt;\r}\r';
    wo.BodyCss += '.' + wo.ndName + ' ol.' + nListId + ' li > * {\r';
    wo.BodyCss += '  padding-left:' + (ind0 - ind1 -14) +'pt;\r}\r';
    return
  }

  if (listId === undefined) {}
  var ordListTyp = glObj.glOrd;
  if (wo.cListUn === undefined) {
    wo.cListUn = ordListTyp;
  } else {
    if (wo.cListUn != ordListTyp) {
      CloseList();
      wo.cListUn = ordListTyp;
    }
  }
  
  if (!ordListTyp) {
  // unorderered list
  // is this the first list item
    
    switch(nestLev) {
      case wo.cNlev:
        // new list
        if (wo.unList[nestLev].listId == null) {
          NewUnList()  
        } else {
          if (wo.unList[nestLev].listId == listId) {
            wo.BodyHtml +='  <li>';
          } else {
            // new list error?
          }
        }

        break;
      case wo.cNlev+1:
        if (wo.unList[nestLev].listId == null) {
          NewUnList()
        } else {
          if (wo.unList[nestLev].listId == listId) {
            wo.BodyHtml +='  <li>';
          } else {
            // new list error?
          }
        }
        // sub list
        break;
      case wo.cNlev -1:
        // back one level
        // prev list
        wo.BodyHtml += '<!-- end ul level: ' + nestLev + ' -->\r';
        wo.BodyHtml +='</ul>\r';
        wo.cNlev--;
        if (wo.unList[nestLev].listId == listId) {
            wo.BodyHtml +='  <li>';
        }      
        break;
      default:
        // should not happen!! 
    } 
  } else {
      // ordered list
    var ind = atts[DocumentApp.Attribute.INDENT_START];
    var curNL = wo.cNlev;
    if (wo.orList[curNL].listId == null) {
          NewOrList();
    } else {
      if (wo.orList[curNL].listId == listId){
        wo.BodyHtml +='  <li>';
      } else {
        // nl increased or decreased?
        // check for decrease; there must be an existing list
        if (ind < wo.indList) {
            // if true decrease
          for (var nl=curNL; nl> -1; nl--) {
            if (wo.orList[nl].listId == listId) { 
              wo.BodyHtml +='  <li>';
              break;
            } 
            wo.BodyHtml += ' </ol>\r';
            wo.orList[nl].count = 0;
            wo.orList[nl].Id = null;
            wo.orList[nl].listId = null;
            wo.cNlev--
          }          
        } else {
          // nest level increased
//          CloseList();
          NewOrList();
        }
      }
    }
  }
 
  var chHasTxt = false;  
  var chTxtStr = child.getText();
  if (chTxtStr.length > 0) {
    chHasTxt = true;
  }

  var par_head = child.getHeading();

//  if (par_head == DocumentApp.ParagraphHeading.TITLE) {
//    wo.Title = chTxtStr;
//  }

        // create a function 
  var container = child.asListItem;
  var numGChild = child.getNumChildren();        

  var posImgList = child.getPositionedImages();
  var imgPNum = posImgList.length;
  var subDiv = false;
  if (imgPNum > 0) {
    subDiv = true;
    wo.subDivCount++;
  }
        
  var headObj =  ConvertHeading(par_head, false);   

  var cssHead = headObj.cssStr;
  var prefix = headObj.htmlPrefix;
  var suffix = headObj.htmlSuffix;
     
  if (chHasTxt) {
          //bodyObj.BodyCss += CvtParAttToCSS(atts, cssHead);
          //pstylstr + ' {\r' + parAttstr + '}\r'
    var cssAttStr = ConvertTxtAttToCss(atts);
    if ((cssAttStr.length > 0)&&(cssHead.length > 0)) {
      wo.BodyCss += cssHead + cssAttStr + '}\r';
    }
          
        // if there are positioned images we need to create a new div which contain paragraphs or headings
        // and the postioned images
        // otherwise simple paragraphs of heading are enough

    var elObj = CvtPar(child, headObj.idStr);
  }
          
  var subDivParHtmlStr = '';
  var subDivParStylStr = '';

    // no positioned Images
  if (!subDiv) {            
    if (chHasTxt) {
      wo.BodyHtml += prefix + elObj.htmlStr + suffix +'\r';
      if (headObj.htmlTocPrefix.length > 1) {
        wo.TocHtml += headObj.htmlTocPrefix + elObj.htmlStr + headObj.htmlTocSuffix + '\r';
      }
      wo.BodyCss += elObj.cssStr;  
    } else {
      wo.BodyHtml += '<br>\r';
    }
  } else {
    // case of positioned images
    var subDivIdStr = wo.ndName + '_' + wo.subDivCount.toString();
    var subDivStylstr = '#' + subDivIdStr +' {\r';
    subDivStylstr += 'position: relative;\r';  
    subDivStylstr += '}\r';
    wo.BodyCss += subDivStylstr;
    if (chHasTxt) {            
      subDivParHtmlStr += '  ' + prefix + elObj.htmlStr + suffix +'\r';
      subDivParStylStr += '#' + subDivIdStr +' p {\r';
      subDivParStylStr += elObj.cssStr;
      subDivParStylStr +='position: relative;\r}\r';
    } else {
      subDivParHtmlStr += '  <br>\r';
    }
    var subDivImgHtmlStr = '';
    var subDivImgStylStr = '';
    for (var im=0; im<posImgList.length; im++){
      var imgel = posImgList[im];
      if (imgel === undefined) { 
        break; 
      }
      var imgblob = imgel.getBlob();
      var imgId = imgel.getId();
      var imgTyp = imgblob.getContentType();
      var imgExt = getImgExt(imgTyp);
      var imgFilNam = imgNam + wo.imgCount + imgExt;
      imgblob.setName(imgFilNam)       
      newFile = DriveApp.createFile(imgblob);

      newFile.moveTo(imgFolder);

      wo.imgCount++;
      var imgDispId = subDivIdStr + '_img_' + (im +1);
      subDivImgHtmlStr += '<!-- positioned Image id: ' + imgId + ' Top: ' + imgel.topOffset() + ' Left: ' + imgel.getLeftOffset() + ' Layout: ' + imgel.getImgPos(imel.getLayout()) + '-->\r';
      subDivImgHtmlStr += '  <img src="' + wo.ImgFoldNam  + '/' + imgFilNam + '" id = "' + imgDispId + '">\r';

      subDivImgStylStr += '#' + imgDispId + ' {\r';
      subDivImgStylStr += 'position: absolute;\r';
      subDivImgStylStr += 'top: ' + imgel.getTopOffset() + 'pt; left: ' +imgel.leftOffset() + 'pt;\r'
      subDivImgStylStr += 'width: ' + imgel.getWidth() + 'px;\r';
      subDivImgStylStr += 'height: ' + imgel.getHeight() + 'px;\r';
      subDivImgStylStr += '}\r';
    } // end pos img loop 

    wo.BodyHtml += '  <div id ="' + subDivIdStr+ '">\r' + subDivParHtmlStr + subDivImgHtmlStr + '  </div>\r';
    wo.BodyCss += subDivParStylStr + subDivImgStylStr;
  }
  // other sub elements of paragraph
  var imgClassStr = subDivIdStr + '_inlImg';
  var drawClassStr = subDivIdStr + '_inlDraw';
    // loop through grand children
  for (var k = 0; k< numGChild; k++) {
    var gchild = child.getChild(k);
    var gchildtyp = gchild.getType();
      // I am assuming paragraphs do not contain other paragraphs, lists or tables for the time being 
    switch(gchildtyp) {
      case DocumentApp.ElementType.INLINE_IMAGE:        
        var urlstr = gchild.getLinkUrl();
        var imgId = gchild.getId(); 
        if (urlstr === null) {
          var imgBlob = gchild.getBlob();
          var imgTyp = imgBlob.getContentType();
          var imgExt = getImgExt(imgTyp);
          var imgFilNam = imgNam + wo.imgCount + imgExt;
          imgBlob.setName(imgFilNam)
          newFile = DriveApp.createFile(imgBlob);
          newFile.moveTo(imgFolder)
          wo.BodyHtml += '<!-- inline image: ' + imgFilNam + ' Id: '+ ingId +' -->\r';
          wo.BodyHtml += '<img src="' + wo.ImgFoldNam + '/' + imgFilNam + '" width=' + gchild.getWidth() + ' height=' + gchild.getHeight() + '>\r<br>\r';
        } else {  
          wo.BodyHtml += '<img src="' + urlstr + '" width=' + gchild.getWidth() + ' height=' + gchild.getHeight() + '>\r<br>\r';
        }
        wo.imgCount++;
        break;       
      // todo
      case DocumentApp.ElementType.INLINE_DRAWING:
              container = child.asInlineDrawing;
        break;
      case DocumentApp.ElementType.TEXT:
        break;
      case DocumentApp.ElementType.PAGE_BREAK:
        wo.BodyHtml += '<!-- page break -->\r';
        break;
      default:
        wo.BodyHtml += '<!-- gchild: ' + k + ' type: ' + gchildtype + ' -->\r';
        break;
    }
  } //end loop of grand children
  wo.BodyHtml += '  </li>\r';
  return
}

function CloseList() {
  if (wo.cListUn === undefined) {return;}
  var ord = wo.cListUn;
  var clev = wo.cNlev;
  if (!ord) {
    for (var i=clev; i> -1; i--) {    
      wo.BodyHtml += '</ul>\r';
      wo.unList[i].count = 0;
      wo.unList[i].Id = null;
      wo.unList[i].listId = null;
    }
    wo.cListUn = undefined;
    return
  }
  for (i=clev; i> -1; i--) {
    wo.BodyHtml += '</ol>\r';
    wo.orList[i].count = 0;
    wo.orList[i].Id = null;
    wo.orList[i].listId = null;
  }
  wo.cListUn = undefined;
  wo.cNlev =0;
  return;
}

function modListId(id) {
  var midStr;
  midStr = id.replace('\.','\\.');
  return midStr
}

function modId(id) {
  var modStr;
  var idx = id.indexOf('.');
  modStr = id.substring(idx+1);
  return modStr
}


function CvtGlyphToObj (glyph) {

  var glyphObj = {
    glTyp: '',
    glOrd: false
  }
      
  switch (glyph) {
    case DocumentApp.GlyphType.BULLET:
      glyphObj.glTyp = 'disc';
      glyphObj.glOrd = false;
      break;
    case DocumentApp.GlyphType.HOLLOW_BULLET:
      glyphObj.glTyp = 'circle';
      glyphObj.glOrd = false;
      break;
    case DocumentApp.GlyphType.SQUARE_BULLET:
      glyphObj.glTyp = 'square';
      glyphObj.glOrd = false;
      break;        
    case DocumentApp.GlyphType.LATIN_UPPER:
      glyphObj.glTyp = 'upper-latin';
      glyphObj.glOrd = true;
      break;      
    case DocumentApp.GlyphType.LATIN_LOWER:
      glyphObj.glTyp = 'lower-latin';
      glyphObj.glOrd = true;
      break;    
    case DocumentApp.GlyphType.NUMBER:
      glyphObj.glTyp = 'decimal';
      glyphObj.glOrd = true;
      break;
    case DocumentApp.GlyphType.ROMAN_UPPER:
      glyphObj.glTyp = 'upper-roman';
      glyphObj.glOrd = true;
      break;  
    case DocumentApp.GlyphType.ROMAN_LOWER:
      glyphObj.glTyp = 'lower-roman';
      glyphObj.glOrd = true;
      break;
    default:
      glyphObj.glTyp = '-';
      glyphObj.glOrd = false;
      break;
  }
  return glyphObj;
}

function getHozAl(hozal) {

  var nvastr = '';

   switch (hozal) {
    case DocumentApp.HorizontalAlignment.CENTER:
      nvastr = 'center';
      break;
    
    case DocumentApp.HorizontalAlignment.RIGHT:
      nvastr = 'right';
      break;

    case DocumentApp.HorizontalAlignment.JUSTIFY:
      nvastr = 'justify';
      break;

    default: 
      
  }

  return nvastr
}

function ConvertTxtAttToCss(atts) {

  var indStr = '  ';

  var hozal = atts[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT];  
  switch (hozal) {
    case DocumentApp.HorizontalAlignment.CENTER:
      nvastr = 'center';
      break;
    
    case DocumentApp.HorizontalAlignment.RIGHT:
      nvastr = 'right';
      break;

    case DocumentApp.HorizontalAlignment.JUSTIFY:
      nvastr = 'justify';
      break;
    case DocumentApp.HorizontalAlignment.LEFT:
      nvastr = 'left';
      break;
    default: 
      nvastr = '';
  }

  var CssStr = '';
  if (nvastr.length > 0) {
    CssStr += indStr + 'text-align: ' + nvastr  + ';\r'
  }
  // still have to deal with vertical align

  if (atts[DocumentApp.Attribute.ITALIC]) {
    CssStr += indStr + 'font-style: italic;\r'
  }

  if (atts[DocumentApp.Attribute.BOLD]) {
    CssStr += indStr + 'font-weight: bold;\r'
  }

  if (atts[DocumentApp.Attribute.UNDERLINE]) {
    CssStr += indStr + 'text-decoration: underline;\r'
  }

  if (atts[DocumentApp.Attribute.FONT_FAMILY] !== null) {
    CssStr += indStr + 'font-family: ' + atts[DocumentApp.Attribute.FONT_FAMILY] + ';\r'
  }

  if (atts[DocumentApp.Attribute.FONT_SIZE] !== null) {
    CssStr += indStr + 'font-size: ' + atts[DocumentApp.Attribute.FONT_SIZE] + 'pt;\r'
  }

  if (atts[DocumentApp.Attribute.FOREGROUND_COLOR] !== null) {
    CssStr += indStr + 'color: ' + atts[DocumentApp.Attribute.FOREGROUND_COLOR] + ';\r'
  }

  if (atts[DocumentApp.Attribute.BACKGROUND_COLOR] !== null) {
    CssStr += indStr + 'background-color:' + atts[DocumentApp.Attribute.BACKGROUND_COLOR] + ';\r'
  }

  if (atts[DocumentApp.Attribute.GLYPH_TYPE] != null) {
    var glyph = atts[DocumentApp.Attribute.GLYPH_TYPE];
    var glObj = CvtGlyphToObj(glyph)
    CssStr += indStr + 'list-style-type:' + glObj.glTyp + ';\r';
  }

  if (atts[DocumentApp.Attribute.HEIGHT] != null) {
    CssStr += indStr + 'height: ' + atts[DocumentApp.Attribute.HEIGHT] + 'pt;\r';
  }

  if (atts[DocumentApp.Attribute.LINE_SPACING] != null) {
    CssStr += indStr + 'line-height: ' + atts[DocumentApp.Attribute.LINE_SPACING] + ';\r';
  }

  if (atts[DocumentApp.Attribute.INDENT_FIRST_LINE] != null) {
    CssStr += indStr + 'text-indent: ' + atts[DocumentApp.Attribute.INDENT_FIRST_LINE] + 'pt;\r';
  }

  if (atts[DocumentApp.Attribute.INDENT_START] != null) {
    CssStr += indStr + 'padding-left: ' + atts[DocumentApp.Attribute.INDENT_START] + 'pt;\r';
  }

  if (atts[DocumentApp.Attribute.INDENT_END] != null) {
    CssStr += indStr + 'padding-right: ' + atts[DocumentApp.Attribute.INDENT_END] + 'pt;\r';
  }

  if (atts[DocumentApp.Attribute.SPACING_BEFORE] != null) {
    CssStr += indStr + 'padding-top: ' +  atts[DocumentApp.Attribute.SPACING_BEFORE] +'pt;\r';
  }

  if (atts[DocumentApp.Attribute.SPACING_AFTER] != null) {
    CssStr += indStr + 'padding-bottom: ' + atts[DocumentApp.Attribute.SPACING_AFTER] + 'pt;\r';
  }

  if (atts[DocumentApp.Attribute.BORDER_WIDTH] != null) {
    CssStr += indStr + 'border-width: ' + atts[DocumentApp.Attribute.BORDER_WIDTH] +'pt;\r';
    CssStr += indStr + 'border-style: solid;\r';
  }

  if (atts[DocumentApp.Attribute.BORDER_COLOR] != null) {
    CssStr += indStr + 'border-color: ' + atts[DocumentApp.Attribute.BORDER_COLOR] +';\r';
  }  

  return CssStr
}

function ConvertHeading(par_head, withId) {
//  var dividstr = wo.ndName;
  let cssHtmlObj = {
    htmlPrefix: '',
    htmlTocPrefix: '',
    htmlSuffix: '',
    htmlTocSuffix: '',
    cssToc: '',
    cssStr: '',
    idStr:'',
  }
  var prefix = '';
  var tocPrefix = '';
  var tocSuffix = '';
  var suffix = '';
  var idstr;
  var tocStyl = '';
  var stylstr = '';
  switch (par_head) {          
    case DocumentApp.ParagraphHeading.HEADING6: 
      wo.h6Count++;
      idstr = wo.ndName + 'h6_' + wo.h6Count;
      prefix += '<h6 id="' + idstr +'">';suffix += '</h6>';
      tocPrefix += '<h6><a href = "#' + idstr + '">';
      tocSuffix += '</a></h6>';
      tocStyl += '#' + idstr + '_Toc h6 {\r  padding-left: 50px;\r  margin: 0px;\r}\r';
      stylstr += '#' + idstr;
      break;
    case DocumentApp.ParagraphHeading.HEADING5: 
      wo.h5Count++;
      idstr = wo.ndName + 'h5_' + wo.h5Count;
      prefix += '<h5 id="' + idstr +'">';suffix += '</h5>';
      tocPrefix += '<h5><a href = "#' + idstr + '">'; 
      tocSuffix += '</a></h5>';
      tocStyl += '#' + idstr + '_Toc h5 {\r  padding-left: 40px;\r  margin: 0px;\r}\r';
      stylstr += '#' + idstr;
      break;
    case DocumentApp.ParagraphHeading.HEADING4:
      wo.h4Count++;
      idstr = wo.ndName + 'h4_' + wo.h4Count;
      prefix += '<h4 id="' + idstr +'">';suffix += '</h4>';
      tocPrefix += '<h4><a href = "#' + idstr + '">'; 
      tocSuffix += '</a></h4>';
      tocStyl += '#' + idstr + '_Toc h4 {\r  padding-left: 30px;\r  margin: 0px;\r}\r'; 
      stylstr += '#' + idstr;
      break;
    case DocumentApp.ParagraphHeading.HEADING3:
      wo.h3Count++;
      idstr = wo.ndName + 'h3_' + wo.h3Count;
      prefix += '<h3 id="' + idstr +'">';suffix += '</h3>';
      tocPrefix += '<h3><a href = "#' + idstr + '">';
      tocSuffix += '</a></h3>';
      tocStyl += '#' + idstr + '_Toc h3 {\r  padding-left: 20px;\r  margin: 0px;\r}\r';
      stylstr += '#' + idstr;
      break;
    case DocumentApp.ParagraphHeading.HEADING2:
      wo.h2Count++;
      idstr = wo.ndName + 'h2_' + wo.h2Count;
      prefix += '<h2 id="' + idstr +'">';suffix += '</h2>';
      tocPrefix += '<h2><a href = "#' + idstr + '">';
      tocSuffix += '</a></h2>';
      tocStyl += '#' + idstr + '_Toc h2 {\r  padding-left: 10px;\r  margin: 0px;\r}\r';
      stylstr += '#' + idstr;
      break;
    case DocumentApp.ParagraphHeading.HEADING1:
      wo.h1Count++;
      idstr = wo.ndName + 'h1_' + wo.h1Count;
      prefix += '<h1 id="' + idstr +'">';suffix += '</h1>';
      tocPrefix += '<h1><a href = "#' + idstr + '">';
      tocSuffix += '</a></h1>';
      tocStyl += '#' + idstr + '_Toc h1 {\r  margin-top: 10px;\r  margin-bottom: 0px;\r}\r';
      stylstr += '#' + idstr;
      break;
    case DocumentApp.ParagraphHeading.TITLE:
      idstr = wo.ndName + '_title';
      prefix += '<p class="' + idstr +'">';suffix += '</p>';
      stylstr += '#' + idstr;
      break;
    case DocumentApp.ParagraphHeading.SUBTITLE:
      idstr = wo.ndName + '_subtitle';
      prefix += '<p class="' + idstr +'">';suffix += '</p>'; 
      stylstr += '#' + idstr;
      break;
    case DocumentApp.ParagraphHeading.NORMAL:
      if (withId){
        wo.parCount++;
        idstr = wo.ndName + 'p_' + wo.parCount;
        prefix += '<p id="' + idstr +'">';
        stylstr += '#' + idstr;
      } else {
        prefix += '<p>';
      }
      suffix += '</p>';  
      break;
    default:
    // add log
      break;
  }
  if (stylstr.length>0) { 
    cssHtmlObj.cssStr = stylstr + ' {\r';
    cssHtmlObj.idStr = idstr;
  }
  cssHtmlObj.htmlPrefix = prefix;
  cssHtmlObj.htmlTocPrefix = tocPrefix;
  cssHtmlObj.htmlTocSuffix = tocSuffix;
  cssHtmlObj.htmlSuffix = suffix;
  cssHtmlObj.cssToc = tocStyl;
  
  return cssHtmlObj;                
}


function CvtPar(item, idstr) {

  let itemObj = {
    cssStr: '',
    htmlStr: ''
  };

  var text = item.getText();
  var numChP = item.getNumChildren();
 
 // can there be more than one text type in a pararaph?
  var indices = null;
  for (k=0; k<numChP; k++){
    elch = item.getChild(k)
    if (elch.getType() == DocumentApp.ElementType.TEXT) {
      indices = elch.getTextAttributeIndices();
      break;
    }    
    // possible other types
  } 
  // rewrite
  if ((indices === null)||(indices === undefined)){
    return itemObj;
  }
  if (indices.length == 1) {
    itemObj.htmlStr = text;
    //itemObj.cssStr = '';
    return itemObj;
  }

  var spanStylStr = '';
  var txtHtmlStr = '';
  var spanCount = 0;
  for (var i=0; i < indices.length; i ++) {
    var partAtts = elch.getAttributes(indices[i]);
    var startPos = indices[i];
    var endPos = i+1 < indices.length ? indices[i+1]: text.length;
    var partText = text.substring(startPos, endPos);

    var txtStylStr = '';

    if (partAtts.LINK_URL) {
      txtHtmlStr = txtHtmlStr + '<a href = "' + partAtts.LINK_URL + '">'
    }
    if (partAtts.STRIKETHROUGH) {
      txtStylStr += txtStylStr + 'text-decoration: line-through;\r';
    //  spanEl = true;
    }
    var spanEl = false;

    txtStylStr += ConvertTxtAttToCss(partAtts);

    // fix later

    if (spanEl) {
      spanCount++;
      txtHtmlStr = txtHtmlStr + '<span>';
      spanStylStr = spanStylStr + idstr + ' span:nth-of-type(' + spanCount + '){\r' + txtStylStr + '}\r';
    }
    var partTxtIns = false;
    if (partText.indexOf('[')==0 && partText[partText.length-1] == ']') {
      txtHtmlStr = txtHtmlStr + '<sup>' + partText + '</sup>';
      partTxtIns = true;
    }
    
    if ((partText.trim().indexOf('http://') == 0) || (partText.trim().indexOf('https://') == 0)) {
      txtHtmlStr = txtHtmlStr + '<a href="' + partText + '" rel="nofollow">' + partText + '</a>';
      partTxtIns = true;
    }
    
    if (!partTxtIns) {
      txtHtmlStr += partText;
    }
    if (spanEl) {
      txtHtmlStr += '</span>';
    }

    if (partAtts.LINK_URL) {
      txtHtmlStr = txtHtmlStr + '</a>';
    } 
  } //for loop

  if (spanStylStr.length > 0) {
    itemObj.cssStr += spanStylStr; 
  }
  itemObj.htmlStr += txtHtmlStr;
  
  return itemObj;
}


function ConvertTable(tab) {
  let tabObj = {
    cssStr: '',
    htmlStr: ''
  };

  var rowCount = tab.getNumRows();
  var irow = 0;
  var topRow = tab.getRow(0);
  var colCount = topRow.getNumCells();

  var tabWidth = 0;
  for (var icol=0; icol < colCount; icol++) {
    tabWidth += tab.getColumnWidth(icol);
  }
  // table set up
  wo.tabCount++;
  var tabidstr = wo.ndName + '_tab' + wo.tabCount;
  tabObj.cssStr = 'table#' + tabidstr + ' {\r';
  tabObj.htmlStr = '<table id="' + tabidstr + '">\r';

  // styling
  tabObj.cssStr += '  border: 1px solid black;\r  border-collapse: collapse;\r';
  tabObj.cssStr += '  width: ' + tabWidth + 'pt;\r';
  tabObj.cssStr += '  margin:auto;\r';
  tabObj.cssStr += '}\r';

  tabObj.cssStr += '#' + tabidstr + ' th,td {\r';
  tabObj.cssStr += '  border: 1px solid black;\r';
  tabObj.cssStr += '  vertical-align: top;\r';
  tabObj.cssStr += '}\r';
  

  for (var irow = 0; irow < rowCount; irow++) {  
    tableRow = tab.getRow(irow);
    var rowHeight = tableRow.getMinimumHeight();
    if (rowHeight > 0) {
      tabObj.cssStr += '#' + tabidstr + ' tr:nth-child(' + (irow + 1) + ') {\r';
      tabObj.cssStr += '  height: ' + rowHeight + 'pt;\r';
      tabObj.cssStr += '}\r'; 
    }
  }
  for (var icol = 0; icol < colCount; icol++) {  
    var colWidth = tab.getColumnWidth(icol);
    if (colWidth > 0) {
      tabObj.cssStr += '#' + tabidstr + ' th:nth-child(' + (icol + 1) + ') {\r';
      tabObj.cssStr += '  width: ' + colWidth + 'pt;\r';
      tabObj.cssStr += '}\r'; 
    }
  }


  // table header
  tabObj.htmlStr += '<thead><tr>\r';
  for (var icol=0; icol < colCount; icol++) {
    var tcell = topRow.getCell(icol);
    var thstr = tcell.getText();
    tabObj.htmlStr += '<th>' + thstr + '</th>\r';
  }
  tabObj.htmlStr += '</thead></tr>\r';

  // table body
  tabObj.htmlStr += '<tbody>\r';
  for (var irow = 1; irow < rowCount; irow++) {
    tabObj.htmlStr += '<tr>\r';
    var tabRow = tab.getRow(irow);
    for (var icol=0; icol< colCount; icol++) {
    var tcell = tabRow.getCell(icol);
    var cellstr = tcell.getText();
      tabObj.htmlStr += '  <td>' + cellstr + '</td>\r';
    }
    tabObj.htmlStr += '</tr>\r';
  }
  tabObj.htmlStr += '</tbody>\r';
  // table end
  tabObj.htmlStr += '</table>\r';
  
  return tabObj;
}
