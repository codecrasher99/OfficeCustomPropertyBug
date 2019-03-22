var gbcodecrasherviaoffice = true;
var gbcodecrasherdocid;
var gbcodecrasherattachmentid;
var gbisCodecrasherDoc = false;
var gbcodecrasherpropname = "CodeCrasherID";
var gbcodecrasherpropssepconst = "{@@@}";
var gbcodecrasherofficehost;
var gbcodecrasherruncontext;
var gbcodecrashercontextobject;

function DisplayOnlineVsDesktop() {
    var docIDMess = "You are using Office 365 " + gbcodecrasherofficehost;
    appnote.showNotification("Console", "Performing check for office online vs desktop");
    if (window.top === window) {
        //the add-in is not running in Online
        appnote.showNotification("Console", docIDMess + " desktop ");
    }
    else {
        //the add-in is running in online
        appnote.showNotification("Console", docIDMess + " online ");
    }
}

function LoadOffice(callback) {
    appnote.showNotification('Console', 'Please wait while we load the Office environment');
    $.getScript('https://appsforoffice.microsoft.com/lib/1/hosted/Office.js', function () {
        //script is loaded and executed put your dependent JS here
        Office.initialize = function (reason) {
            $(document).ready(function () {
                appnote.showNotification('Console', 'Office initialized and document ready');
                gbcodecrasherofficehost = Office.context.host;
                DisplayOnlineVsDesktop();
                ExecuteInContext(function () {
                    return readCustomProperty(callback);
                });
            });
        };
    });
    appnote.showNotification("Console", "Done loading office js");
}

function readCodeCrasherDocProperties(propval) {
    appnote.showNotification("Console", "In readCodeCrasherDocProperties");
    if (gbcodecrasherviaoffice) {
        if (gbisCodecrasherDoc) {
            if (propval.indexOf(gbcodecrasherpropssepconst) >= 0) {
                var pos = propval.indexOf(gbcodecrasherpropssepconst);
                gbcodecrasherdocid = propval.substring(0, pos);
                gbcodecrasherattachmentid = propval.substring(pos + 5);
            }
        }
    }
    else {
        var curr = window.location.href;
        var uri = getParameterByName("ID", curr);

        // check if parameter ID is present
        if (!(uri === null)) {
            gbisCodecrasherDoc = true;
            gbcodecrasherdocid = uri.toString();
            //Put a dummy attachment ID for testing
            gbcodecrasherattachmentid = "{f49846d6-c039-4307-9596-36a586b742de}";
            console.log("CodeCrasher ID from uri " + gbcodecrasherdocid);
        }
    }
}

function EvalueDocPropsForCodeCrasher(propval) {
    appnote.showNotification("Console", "In EvalueDocPropsForCodeCrasher");
    var retval;
    if (typeof (propval) === 'undefined') {
        retval = false;
        appnote.showNotification("Notification", "This document does not exist in CodeCrasher ");
    }
    else {
        appnote.showNotification("Notification", "This document exists in CodeCrasher " + gbcodecrasherpropname + " = " + propval);
        gbisCodecrasherDoc = true;
        readCodeCrasherDocProperties(propval);
        retval = true;
    }
    return retval;
}

function GetHostSpecificPropertyCollection() {
    appnote.showNotification("Console", "In GetHostSpecificPropertyCollection for host " + gbcodecrasherofficehost);
    var customDocProps;
    switch (gbcodecrasherofficehost) {
        case Office.HostType.Word:
            customDocProps = gbcodecrasherruncontext.document.properties.customProperties.getItemOrNullObject(gbcodecrasherpropname);
            break;

        case Office.HostType.Excel:
            customDocProps = gbcodecrasherruncontext.workbook.properties.custom.getItemOrNullObject(gbcodecrasherpropname);
            break;
    }
    return customDocProps;
}

function AddHostSpecificPropertyCollection(propertyname, value) {
    appnote.showNotification("Console", "In AddHostSpecificPropertyCollection for host " + gbcodecrasherofficehost);
    switch (gbcodecrasherofficehost) {
        case Office.HostType.Word:
            gbcodecrasherruncontext.document.properties.customProperties.add(propertyname, value);
            gbcodecrasherruncontext.document.save();
            break;

        case Office.HostType.Excel:
            gbcodecrasherruncontext.workbook.properties.custom.add(propertyname, value);
            //TODO Saving of excel workbook doesnt work. When Microsoft fixes the same uncomment the below line
            //gbcodecrasherruncontext.workbook.save();
            break;
    }
}

function insertHostSpecificCustomProperty(propertyname, value, CallbackNavigate) {
    appnote.showNotification("Console", "In insertHostSpecificCustomProperty " + propertyname);
    AddHostSpecificPropertyCollection(propertyname, value);
    return gbcodecrasherruncontext.sync().
        then(function () {
            appnote.showNotification("Console", "Property " + propertyname + " with value " + value + " has been inserted successfully ");
            CallbackNavigate();
        })
        .catch(function (e) {
            appnote.showNotification("Notification", "Error occured inserting property " + e.message);

        });
    }


function readCustomProperty(callbackWithPropertyValueAsArgument) {
    appnote.showNotification("Console", "Reading document properties " + gbcodecrasherpropname);
    appnote.showNotification("Console", "Proceeding to execute context");
    var customDocProps = GetHostSpecificPropertyCollection();
    gbcodecrasherruncontext.load(customDocProps);
    return gbcodecrasherruncontext.sync()
        .then(function () {
            callbackWithPropertyValueAsArgument(customDocProps.value);
        })
        .catch (function (e) {
        appnote.showNotification("Notification", "Error occured reading custom property " + e.message);

    });

}

function getOfficeFileSavedLocation(successcallback) {
    appnote.showNotification("Console", "In getOfficeFileSavedLocation");
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        appnote.showNotification("Console", "getFilePropertiesAsync resultstatus - " + asyncResult.status);
        if ((asyncResult.status === Office.AsyncResultStatus.Succeeded)) {
            var savedURL = asyncResult.value.url;
            appnote.showNotification("Console", "getFilePropertiesAsync returned URL - " + savedURL);
            successcallback(savedURL);
        }
        else {
            //TODO in Excel getFilePropertiesAsync gives error when called on unsaved document. 
            //We are using this error condition to identify the files saved status.
            //When Microsft fixes the issue this code needs to be reworked.
            $("body").css("cursor", "default");
            var err = asyncResult.error;
            appnote.showNotification("Console", "Error occured retriving file save status " + err.message, false);
            appnote.showNotification("Notification", "The document has not been saved in Office365. Please save the document to Office365 to use CodeCrasher properties", false);
        }

    });

}

function ExecuteInContext(functionname) {
    appnote.showNotification("Console", "Executing in context " + gbcodecrasherofficehost);
    switch (gbcodecrasherofficehost) {
        case Office.HostType.Word:
            Word.run(function (context) {
                gbcodecrasherruncontext = context;
                gbcodecrashercontextobject = context.document;
                return functionname();
            });
            break;

        case Office.HostType.Excel:
            Excel.run(function (context) {
                gbcodecrasherruncontext = context;
                gbcodecrashercontextobject = context.workbook;
                return functionname();
            });
            break;

        case Office.HostType.PowerPoint:
            PowerPoint.run(function (context) {
                gbcodecrasherruncontext = context;
                return functionname();
            });
            break;

    }

}

