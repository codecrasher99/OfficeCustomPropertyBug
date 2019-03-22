/* Common app functionality */

var appnote = (function () {
    "use strict";

    var appnote = {};

     // Common initialization function (to be called from each page)
    appnote.initialize = function () {
     
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });

        var hider;

        function dodelayHide() {
            hider = setTimeout(function () {
                $('#notification-message').hide();
            }, 5000);
            //console.log(" dodelayHide - " + hider);
        }

        function stopdelayHide() {
            //console.log(" stopdelayHide - " + hider);
            clearTimeout(hider);
        }

        // After initialization, expose a common notification function
        appnote.showNotification = function (header, text, dotimerhide)
        {
            dotimerhide = (typeof dotimerhide === 'undefined') ? true : dotimerhide;
            dotimerhide = false;
            console.log(text);
            if (header == 'Notification')
            {
                $('#notification-message-header').text(header);
                $('#notification-message-body').text(text);
                //stopdelayHide();
                $('#notification-message').slideDown('fast');
                if (dotimerhide) {
                    dodelayHide();
                }
            }
        };

    };

    return appnote;
})();

