function onInstall(e) {
  onOpen(e);
  // Perform additional setup as needed.
}

function onOpen(e) {  
  Logger.log('AuthMode: ' + e.authMode);
  var lang = Session.getActiveUserLocale();
//  var menu = DocumentApp.getUi().createMenu("瞬速メッセンジャー")
  var ui = DocumentApp.getUi();  
  var menu = ui.createAddonMenu();
  if(e && e.authMode == 'NONE'){
    var startLabel = lang === 'ja' ? '使用開始' : 'start';
    menu.addItem(startLabel, 'askEnabled');
  } else {
    if( lang === 'ja')
    {
    menu.addItem('HTMLコード生成', 'ConvertGoogleDocToCleanHtml');
    }
    else{
    menu.addItem('Generate HTML', 'ConvertGoogleDocToCleanHtml');
    }
  };
  menu.addToUi();

};

function askEnabled(){
  var lang = Session.getActiveUserLocale();
  var title = 'Your Script\'s Title';
  var msg = lang === 'ja' ? '瞬速メッセンジャーが有効になりました。ブラウザを更新してください。' : 'Rapid Messenger has been enabled.';
  var ui = DocumentApp.getUi();
  ui.alert(title, msg, ui.ButtonSet.OK);
};



function getLastIndexOfRichEditor(body){
  var searchTypeHR = DocumentApp.ElementType.HORIZONTAL_RULE;
  var numChildren = body.getNumChildren();
  var firstHR = body.findElement(searchTypeHR);
  var lastIndex = firstHR == null ? numChildren : body.getChildIndex(firstHR.getElement().getParent());
  Logger.log("lastIndex:" +lastIndex );
  return lastIndex;
}

function ConvertGoogleDocToCleanHtml() {
  var body = DocumentApp.getActiveDocument().getBody();
  var lastIndex = getLastIndexOfRichEditor(body);

  var output = [];
  var images = [];
  var listCounters = {};

  // Walk through all the child elements of the body.
  for (var i = 0; i < lastIndex; i++) {
    var child = body.getChild(i);
    output.push(processItem(child, listCounters, images));
  }
  Logger.log("output:" + output.length);

  var html = output.join('\r');
  //emailHtml(html, images);
  // createDocumentForHtml(html, images);
  
  // ファイルの末尾にソースを生成
  var sourceBody = getHtmlPartBody(body);
//  sourceBody.appendParagraph('\r');
  sourceBody.appendParagraph(html);
//  for(var j=0; j < images.length; j++)
//    sourceBody.appendImage(images[j].blob);
  
  
}

// 生成したソースを、現在のファイルの末尾に追加するためのbodyを返します。
function getHtmlPartBody(body){
    // 水平線以降の文を取得します。
    var searchTypeParagraph = DocumentApp.ElementType.PARAGRAPH;
    var searchTypeHR = DocumentApp.ElementType.HORIZONTAL_RULE;

    var firstHR = body.findElement(searchTypeHR);
    if (firstHR == null) {
      body.appendHorizontalRule();
      firstHR = body.findElement(searchTypeHR);
    }
    var theHr = firstHR.getElement();
    Logger.log(theHr.getParent());

    // 水平線の後が本文。古い本文を消し、新しい本文を追加する  
//    DocumentApp.getActiveDocument().getBody();

    var images = body.getImages();
    images.map(function(theImage) {
      var isAfterHR = body.getChildIndex(theHr.getParent()) < body.getChildIndex(theImage.getParent());
      if(isAfterHR ){
        theImage.removeFromParent();              
      } 
    });

    // 最後の1文は削除できないので、空白に置き換える
    var paragraphs = body.getParagraphs();  
    paragraphs.map(function(theParagraph) {
       if(body.getChildIndex(theHr.getParent()) < body.getChildIndex(theParagraph)){
         if(!theParagraph.isAtDocumentEnd())
           theParagraph.removeFromParent();         
         else 
           theParagraph.setText(' ');
      } 
    });

  return body;
}


// 生成したソースを、新しいファイルに出力します。
function createDocumentForHtml(html, images) {
  var name = DocumentApp.getActiveDocument().getName()+".html";
  var newDoc = DocumentApp.create(name);
  newDoc.getBody().setText(html);
  for(var j=0; j < images.length; j++)
    newDoc.getBody().appendImage(images[j].blob);
  newDoc.saveAndClose();
}

function dumpAttributes(atts) {
  // Log the paragraph attributes.
  for (var att in atts) {
    Logger.log(att + ":" + atts[att]);
  }
}

function processItem(item, listCounters, images) {
  var output = [];
  var prefix = "", suffix = "";

  if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    switch (item.getHeading()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
      case DocumentApp.ParagraphHeading.HEADING6: 
        prefix = "<h6>", suffix = "</h6>"; break;
      case DocumentApp.ParagraphHeading.HEADING5: 
        prefix = "<h5>", suffix = "</h5>"; break;
      case DocumentApp.ParagraphHeading.HEADING4:
        prefix = "<h4>", suffix = "</h4>"; break;
      case DocumentApp.ParagraphHeading.HEADING3:
        prefix = "<h3>", suffix = "</h3>"; break;
      case DocumentApp.ParagraphHeading.HEADING2:
        prefix = "<h2>", suffix = "</h2>"; break;
      case DocumentApp.ParagraphHeading.HEADING1:
        prefix = "<h1>", suffix = "</h1>"; break;
      default: 
        prefix = "<p>", suffix = "</p>";
    }

    if (item.getNumChildren() == 0)
      return "<br/>";
  }
  else if (item.getType() == DocumentApp.ElementType.INLINE_IMAGE)
  {
    processImage(item, images, output);
  }
  else if (item.getType()===DocumentApp.ElementType.LIST_ITEM) {
    var listItem = item;
    var gt = listItem.getGlyphType();
    var key = listItem.getListId() + '.' + listItem.getNestingLevel();
    var counter = listCounters[key] || 0;

    // First list item
    if ( counter == 0 ) {
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        prefix = '<ul><li>', suffix = "</li>";

          suffix += "</ul>";
        }
      else {
        // Ordered list (<ol>):
        prefix = "<ol><li>", suffix = "</li>";
      }
    }
    else {
      prefix = "<li>";
      suffix = "</li>";
    }

    if (item.isAtDocumentEnd() || (item.getNextSibling() && (item.getNextSibling().getType() != DocumentApp.ElementType.LIST_ITEM))) {
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        suffix += "</ul>";
      }
      else {
        // Ordered list (<ol>):
        suffix += "</ol>";
      }

    }

    counter++;
    listCounters[key] = counter;
  }

  output.push(prefix);

  if (item.getType() == DocumentApp.ElementType.TEXT) {
    processText(item, output);
  }
  else {


    if (item.getNumChildren) {
      var numChildren = item.getNumChildren();

      // Walk through all the child elements of the doc.
      for (var i = 0; i < numChildren; i++) {
        var child = item.getChild(i);
        output.push(processItem(child, listCounters, images));
      }
    }

  }

  output.push(suffix);
  return output.join('');
}


function processText(item, output) {
  var text = item.getText();
  var indices = item.getTextAttributeIndices();

  if (indices.length <= 1) {
    // Assuming that a whole para fully italic is a quote
    if (-1 < text.indexOf('http://') || -1 < text.indexOf('https://')) {
      text = '<a href="' + text + '" rel="nofollow">' + text + '</a>';
    }
    if(item.isBold()) {
      text = '<strong>' + text + '</strong>';
    }
    if(item.isItalic()) {
      text = '<blockquote>' + text + '</blockquote>';
    }
    output.push(text);
  }
  else {

    for (var i=0; i < indices.length; i ++) {
      var partAtts = item.getAttributes(indices[i]);
      var startPos = indices[i];
      var endPos = i+1 < indices.length ? indices[i+1]: text.length;
      Logger.log(i + "個目 startPos:" + startPos + " endPos:" + endPos);
      var partText = text.substring(startPos, endPos);

      Logger.log(partText);

      var color = item.getForegroundColor(startPos);
      
      if(color != null){
        Logger.log("color:" + color);
        output.push('<span style="color:' + color + ';">');
      }
      if (partAtts.ITALIC) {
        output.push('<i>');
      }
      if (partAtts.BOLD) {
        output.push('<strong>');
      }
      if (partAtts.UNDERLINE) {
        output.push('<u>');
      }

      // If someone has written [xxx] and made this whole text some special font, like superscript
      // then treat it as a reference and make it superscript.
      // Unfortunately in Google Docs, there's no way to detect superscript
      if (partText.indexOf('[')==0 && partText[partText.length-1] == ']') {
        output.push('<sup>' + partText + '</sup>');
      }
      else if (partText.trim().indexOf('http://') == 0 || partText.trim().indexOf('https://') == 0) {
        output.push('<a href="' + partText + '" rel="nofollow">' + partText + '</a>');
      }
      else {
        output.push(partText);
      }
      if(color != null){
        output.push('</span>');
      }
      if (partAtts.ITALIC) {
        output.push('</i>');
      }
      if (partAtts.BOLD) {
        output.push('</strong>');
      }
      if (partAtts.UNDERLINE) {
        output.push('</u>');
      }

    }
  }
}

/// 画像を、imgタグに変換します。
/// item: 画像
/// images: imagesコンポーネントへの参照
/// output: 出力HTMLへの参照
function processImage(item, images, output)
{
  images = images || [];
  var blob = item.getBlob();
  var contentType = blob.getContentType();
  var extension = "";
  if (/\/png$/.test(contentType)) {
    extension = ".png";
  } else if (/\/gif$/.test(contentType)) {
    extension = ".gif";
  } else if (/\/jpe?g$/.test(contentType)) {
    extension = ".jpg";
  } else {
    throw "Unsupported image type: "+contentType;
  }
  
  // GoogleDriveへ、画像を保存
  var imagePrefix = "Image_";
  var imageCounter = images.length;
  var name = DocumentApp.getActiveDocument().getId() + imageCounter + extension;
  var imgFolder = createImageFolder();
  blob.setName(name);
    
  var driveFile =imgFolder.createFile(blob);
//    var driveFile = DriveApp.createFile(formBlob);
  //  var image = {myfile:formBlob};
  
  if(driveFile)
  var fileId = driveFile.getId();

  
  // 保存した画像のURLで、imgタグを生成
  output.push('<img  src="http://drive.google.com/uc?export=view&id=' +fileId+'"  width=290px />');
//  images.push( {
//    "blob": blob,
//    "type": contentType,
//    "name": name});
}


