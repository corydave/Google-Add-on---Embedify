/*

Copyright (c) 2016 Dave Ghidiu

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

*/

var VERSION_NUMBER = "1.0";
var embedText = "";

/**
 * A spot for testing code quickly
 *
 */
function test() {

// This spot reserved for quickdraw code testing.

}


/** 
 * During the 'install' process, this method will merely call onOpen().
 */
function onInstall() {
  onOpen();
}




/**
 * Returns the version number of the software
 * @returns {Number} VERSION_NUMBER - The version of the software stored as a global variable
 */
function version() {

  return VERSION_NUMBER;

}




/**
 * Displays the menu bar item
 */
function onOpen() {
  
  DocumentApp.getUi().createAddonMenu()
    .addItem("Embed-ify", "embedify")
    .addItem("About", "about")
    .addToUi();
    
}




/**
 * Displays the "embedifyHTML.html" file (which will set the permission and parameters for the embed code)
 */
function embedify() {
 

  // Display the "embedifyHTML.html" file in a modal dialog
  var html = HtmlService.createTemplateFromFile('embedifyHTML').evaluate().setHeight(400).setWidth(600);
  DocumentApp.getUi().showModalDialog(html, ' ');
      
}




/**
 * Displays the "aboutEmbedify.html" file in a modal dialog
 */
function about() {
 
  // Display the "embedifyHTML.html" file in a modal dialog
  var html = HtmlService.createTemplateFromFile('aboutEmbedify').evaluate().setHeight(350).setWidth(500);
  DocumentApp.getUi().showModalDialog(html, ' ');
    
}




/**
 * Sets the access of this file to ANYONE_WITH_LINK and the permission is changed based on the argument
 * @param {String} permission is the permission level ("comment" or "view") that the document will be shared as
 */
function setPermission(permission) {
  
  // Get the file ID of the current document
  var ID = DocumentApp.getActiveDocument().getId();
  
  // Get the document by the ID
  var doc = DriveApp.getFileById(ID);

  // Set access to ANYONE_WITH_LINK and sets permission equal to the value of the parameter
  if (permission.equals("comment")) {
    doc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.COMMENT);
  } else if (permission.equals("edit")) {
    doc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  } else {
    doc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);  
  }

}




/**
 * Returns the URL of the document
 * @returns {String} URL - The URL of the document
 */
function getURL() {
  return DocumentApp.getActiveDocument().getUrl();
}




/**
 * Sets the access of this file to ANYONE_WITH_LINK and the permission is changed based on the argument
 * @param {String} align: The parameter ("center", "left", "right") for aligning the document - will need <div style=""></div> with LEFT and RIGHT
 * @param {String width: The width of the <iframe>
 * @param {String height: The height of the <iframe>
 */
function embedString(align, width, height) {

  // Get the link of the current file
  var link = DriveApp.getFileById(DocumentApp.getActiveDocument().getId()).getUrl();
  
  // Prepare the HTML Embed code
  var prepend = "<div style='align:" + align + "'>";
  var embed = "<iframe src='" + link + "' width='" + width + "px' height='" + height + "px'>...loading...</iframe>";
  var append = "</div>";
  
  // Piece together the HTML Embed code
  var embedCode = prepend + embed + append;
    
  // Return the code
  return embedCode;

}
