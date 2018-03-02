/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See License.txt in the project root.
*/

// The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
	$(document).ready(function () {
        $('#set-data').click(writeText);
        
        // Use this to check whether the API is supported in the Word client.
        if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
            // Do something that is only available via the new APIs
            $('#bind').click(bind);
            $('#displayAllBindings').click(displayAllBindings);
            $('#clearBinding').click(clearBinding);
            Office.context.document.addHandlerAsync("documentSelectionChanged", selectionChanged, function (result) {});
        } else {
            // Just letting you know that this code will not work with your version of Word.
            $('#supportedVersion').html('Sorry, this add-in requires Word 2016 or greater.');
        }        
    });    

	 //UI Components init
     $(".ms-Pivot").Pivot();
     $(".ms-SearchBox").SearchBox();
     $(".ms-Dropdown").Dropdown();
     $(".ms-ListItem").ListItem();
};

// Reads data from current document selection and displays a notification
function writeText() {
    Office.context.document.setSelectedDataAsync("Citation goes here",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed"){
            	$('#display-data').text("Failure" + error.message);
            }
            else
            {
            	$('#display-data').text("Done");
            }
        });
}

function displayAllBindings() {
    Office.context.document.bindings.getAllAsync(function (asyncResult) {
      var bindingString = '';
      for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
      }
      write('Existing bindings: ' + bindingString);
    });
  }

  // Event handler function.
  function selectionChanged(eventArgs) {
    getText();
  }

  function getText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {
        valueFormat: "unformatted",
        filterType: "all"
      },
      function (asyncResult) {
        var error = asyncResult.error;
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          write(error.name + ": " + error.message);
        } else {
          // Get selected data.
          var dataValue = asyncResult.value;
          write(dataValue);
        }
      });
  }


  // Function that writes to a div with id='message' on the page.
  function write(message) {
    document.getElementById('display-data').innerText = message;
  }

  function bind() {
    Word.run(function (context) {
      Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, {
          id: 'MyBinding'
        },
        function (asyncResult) {
          write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
      );
    });
  }

  function clearBinding() {
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the content controls collection.
        var contentControls = context.document.contentControls;

        // Queue a command to load the content controls collection.
        contentControls.load('text');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {

          if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
          } else {

            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
              console.log('Content control cleared of contents.');
            });
          }

        });
      })
      .catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
          console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
      });
  }

