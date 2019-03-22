(function () {

    function readCustomProperty(propertyname) {
        appnote.showNotification("Notification", "Checking document properties " + propertyname);
        Word.run(function (context) {
            appnote.showNotification("Notification", "Proceeding to execute context");
            var customDocProps = context.document.properties.customProperties;
            context.load(customDocProps);
            return context.sync()
                .then(function () {
                    appnote.showNotification("Notification", "Proceeding to evaluate execute context results");
                    var docidprop = customDocProps.getItemOrNullObject(propertyname);
                    context.load(docidprop);
                    return context.sync()
                        .then(function () {
                            if (typeof (docidprop.value) === 'undefined') {
                                appnote.showNotification("Notification", "This document does not have saved properties. Please click on Save doc property button to add the document properties.");
                            }
                            else {
                                appnote.showNotification("Notification", "This document has saved properties. Please click on Save doc property button to change the document properties. ");
                                appnote.showNotification("Notification", "This document has saved properties " + propertyname + " = " + docidprop.value);
                                PopulateWordPropsToGlobalVariables(docidprop.value);
                            }
                        });
                });
        });
    }

    function saveDocViaWord() {

        appnote.showNotification("Notification","Proceeding to save document changes to document repository");
        $("body").css("cursor", "progress");

        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a commmand to load the document save state (on the saved property).
            context.load(thisDocument, 'saved');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                if (thisDocument.saved === false) {
                    // Queue a command to save this document.
                    thisDocument.save();

                    // Synchronize the document state by executing the queued commands, 
                    // and return a promise to indicate task completion.
                    return context.sync().then(function () {
                        setDummyPropsInGlobalVar();
                        insertCustomProperty(codecrasherpropname, codecrasherpropertyvalue);
                    });
                } else {
                    appnote.showNotification("Notification", "The document has not changed since the last save.");
                    setDummyPropsInGlobalVar();
                    insertCustomProperty(codecrasherpropname, codecrasherpropertyvalue);
                }
            });
        })
            .catch(function (error) {
                $("body").css("cursor", "default");
                appnote.showNotification("Notification","Error occured saving the document in word: " + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    appnote.showNotification("Notification","Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

    function NavigateToEditMode(procmess) {
        window.onbeforeunload = null;
        $("body").css("cursor", "progress");
        var curr = "CodeCrasherWebForm.aspx";
        appnote.showNotification("Notification", "Navigating to edit mode..." + curr);
        appnote.showNotification("Notification", "Please wait while we load the saved changes...");
        $(location).attr("href", curr);
    }

    function SaveDocumentProperties() {
        appnote.showNotification("Notification", "Proceeding to save the doc custom properties");
        var ID = "{ID-GUID-Part}";
        var Attachment = "{f49846d6-c039-4307-9596-36a586b742de}";
        saveDocIDToProps(ID, Attachment);
    }

    function saveDocIDToProps(ID, Attachment) {
            gbcodecrasherdocid = ID;
            gbcodecrasherattachmentid = Attachment;
            var propvalstring = gbcodecrasherdocid + gbcodecrasherpropssepconst + gbcodecrasherattachmentid;
            appnote.showNotification("Notification", "Saving the codecrasher custom doc properties" + propvalstring);
            if (gbcodecrasherviaoffice) {
                ExecuteInContext(function () {
                    return insertHostSpecificCustomProperty(gbcodecrasherpropname, propvalstring, NavigateToEditMode);
                });
            }
            else {
                NavigateToEditMode();
            }
    }

    function DisplayIfCodeCrasherDoc(propval) {
        if (EvalueDocPropsForCodeCrasher(propval)) {
            appnote.showNotification("Notification", "This document has custom properties. Please click on Save button to change the document properties. ");
        }
        else {
            appnote.showNotification("Notification", "This document does not have custom properties. Please click on Save button to save the custom document properties.");
        }
        $("body").css("cursor", "default");
    }

    function LoadAllScripts() {
        console.log("Loading all scripts");
        $.getScript('../pages/js/URI.js', function () {
            $('head').append('<link rel="stylesheet" type="text/css" href="../pages/js/Appbar.css">');
            $.getScript('../pages/js/Appbar.js', function () {
                appnote.initialize();
                appnote.showNotification('Notification', 'Welcome to Office WebForm. Please wait while we initialize your environment.');
                $.getScript('../pages/js/CodeCrasherOffice365.js', function () {
                    $("body").css("cursor", "progress");
                    if (gbcodecrasherviaoffice)
                        LoadOffice(DisplayIfCodeCrasherDoc);
                    else
                        readCodeCrasherDocProperties("");
                    $("body").css("cursor", "default");
                });
            });
        });

    }


    function onDocumentLoaded() {

        LoadAllScripts();

         $("#btnhtml").on("click", SaveDocumentProperties);

         $(document.body).css('margin-left', '10px');
         $(document.body).css('margin-top', '10px');
         $("#notification-message").show();

    }

    /* jQuery may not be available at this point, so use vanilla JS to hook into the events */
    if (window.addEventListener) {
        window.addEventListener("load", onDocumentLoaded);

    } else if (window.attachEvent) {
        window.attachEvent("onload", onDocumentLoaded);
    }
})();
