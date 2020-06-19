/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    document.getElementById("insert-html").onclick = insertHTML;
    document.getElementById("read-html").onclick = readHTML;
  }
});
export async function readHTML() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Html, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
  });
}

// Function that writes to a div with id='message' on the page.
export async function  write(message){
  document.getElementById('message').innerText += message; 
}

export async function insertHTML() {
  return Word.run(async context => {
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p><svg data-diagramref="myfirstDia" height="100" width="100">'+
    '<circle cx="50" cy="50" r="40" stroke="black" stroke-width="3" fill="red" />' +
    'Sorry, your browser does not support inline SVG.  ' +
  '</svg> ', "End");
    await context.sync();
  });

}

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
