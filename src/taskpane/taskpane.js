/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import { base64Image } from "../../base64Image";
import axios from "axios";

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }//determine si notre version de word supporte word.js

    //event handlers et autre pour initialisation logique
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("insert-title").onclick = insertTitle;
    document.getElementById("insert-header1").onclick = insertHead1;

    document.getElementById("apply-style").onclick = applyStyle;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    document.getElementById("replace-text").onclick = replaceText;
    document.getElementById("insert-image").onclick = insertImage;
    document.getElementById("insert-html").onclick = insertHTML;
    document.getElementById("insert-table").onclick = insertTable;
    document.getElementById("create-content-control").onclick = createContentControl;
    document.getElementById("replace-content-in-control").onclick = replaceContentInControl;

  }
});


function insertParagraph() {
  Word.run(context => {
      axios.get('https://reqres.in/api/users')
        .then(resp => {
          // handle success
          showOutput(resp);
        })
        .catch(error => {
          // handle error
          console.error(error);
        }
      )
      function showOutput(resp) {
        var docBody = context.document.body;
        docBody.insertParagraph(
          `${resp.data.data[0].email}`,
          "End"
        );
        var styleText = context.document.body.paragraphs.getLast();
        styleText.styleBuiltIn = Word.Style.normal;
      }

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}


function insertTitle() {
    Word.run(context => {
      axios.get('https://reqres.in/api/users')
        .then(resp => {
          // handle success
          setTitle(resp);
        })
        .catch(error => {
          // handle error
          console.error(error);
        }
      )
      
      function setTitle(resp) {
        var docBody = context.document.body;
        docBody.insertParagraph(
          `${resp.data.data[0].email}`,
          "Start"
        );
        var firstParagraph = context.document.body.paragraphs.getFirst();
        firstParagraph.styleBuiltIn = Word.Style.title;
      }

      return context.sync();
      
    })
  .catch(function (error) 
  {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) 
    {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertHead1() {
  Word.run(context => {
    axios.get('https://reqres.in/api/users')
      .then(resp => {
        // handle success
        setHead1(resp);
      })
      .catch(error => {
        // handle error
        console.error(error);
      }
    )
    
    function setHead1(resp) {
      var docBody = context.document.body;
      docBody.insertParagraph(
        `${resp.data.data[0].email}`,
        "End"
      );
      var styleText = context.document.body.paragraphs.getLast();
      styleText.styleBuiltIn = Word.Style.heading1;
    }

    return context.sync();
  
  })
  .catch(function (error) 
  {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) 
    {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function applyStyle() {
  Word.run(function (context) {

      var firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.styleBuiltIn = Word.Style.intenseReference;

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function applyCustomStyle() {
  Word.run(function (context) {

      var lastParagraph = context.document.body.paragraphs.getLast();
      lastParagraph.style = "MyCustomStyle";

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function changeFont() {
  Word.run(function (context) {

      var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
      secondParagraph.font.set({
        name: "Courier New",
        bold: true,
        size: 18
      });

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertTextIntoRange() {
  Word.run(function (context) {

      var doc = context.document;
      var originalRange = doc.getSelection();
      originalRange.insertText(" (C2R)", "End");

      originalRange.load("text");
      return context.sync()
        .then(function() {
            doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
        })
        .then(context.sync);
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertTextBeforeRange() {
  Word.run(function (context) {

      var doc = context.document;
      var originalRange = doc.getSelection();
      originalRange.insertText("Office 2019, ", "Before");
      originalRange.load("text");
      return context.sync()
        .then(function() {
          doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
        })
        .then(context.sync);
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function replaceText() {
  Word.run(function (context) {

      var doc = context.document;
      var originalRange = doc.getSelection();
      originalRange.insertText("many", "Replace");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertImage() {
  Word.run(function (context) {

      context.document.body.insertInlinePictureFromBase64(base64Image, "End");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertHTML() {
  Word.run(function (context) {

      var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
      blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");    

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertTable() {
  Word.run(function (context) {

      var secondParagraph = context.document.body.paragraphs.getFirst().getNext();

      var tableData = [
        ["Name", "ID", "Birth City"],
        ["Bob", "434", "Chicago"],
        ["Sue", "719", "Havana"],
      ];
secondParagraph.insertTable(3, 3, "After", tableData);


      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function createContentControl() {
  Word.run(function (context) {

      var serviceNameRange = context.document.getSelection();
      var serviceNameContentControl = serviceNameRange.insertContentControl();
      serviceNameContentControl.title = "Service Name";
      serviceNameContentControl.tag = "serviceName";
      serviceNameContentControl.appearance = "Tags";
      serviceNameContentControl.color = "blue";

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function replaceContentInControl() {
  Word.run(function (context) {

      var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
      serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
  
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}
