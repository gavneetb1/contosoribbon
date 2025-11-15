// 1. How to construct online meeting details.
// Not shown: How to get the meeting organizer's ID and other details from your service.
const newBody = '<div style="font-family: Arial, sans-serif; padding: 20px; background-color: #f9f9f9; border-left: 4px solid #0078d4; margin: 10px 0;">' +
    '<h2 style="color: #0078d4; margin-top: 0; font-size: 18px;">🎥 Join Contoso Meeting</h2>' +
    '<div style="background-color: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">' +
    '<p style="margin: 10px 0;"><strong style="color: #323130;">Meeting Link:</strong></p>' +
    '<a href="https://contoso.com/meeting?id=123456789" target="_blank" style="display: inline-block; background-color: #0078d4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; margin: 10px 0;">Join Meeting Now</a>' +
    '<hr style="border: none; border-top: 1px solid #edebe9; margin: 20px 0;">' +
    '<p style="margin: 10px 0;"><strong style="color: #323130;">📞 Phone Dial-in:</strong> <span style="color: #106ebe; font-weight: bold;">+1 (123) 456-7890</span></p>' +
    '<p style="margin: 10px 0;"><strong style="color: #323130;">🔑 Meeting ID:</strong> <span style="font-family: monospace; background-color: #f3f2f1; padding: 4px 8px; border-radius: 3px;">123 456 789</span></p>' +
    '<hr style="border: none; border-top: 1px solid #edebe9; margin: 20px 0;">' +
    '<p style="margin: 10px 0; color: #605e5c; font-size: 14px;">Want to test your video connection?</p>' +
    '<a href="https://contoso.com/testmeeting" target="_blank" style="display: inline-block; background-color: #107c10; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; margin: 10px 0;">Test Your Connection</a>' +
    '</div>' +
    '</div><br>';


let mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

// 2. How to define and register a function command named `insertContosoMeeting` (referenced in the manifest)
//    to update the meeting body with the online meeting details.
function insertContosoMeeting(event) {
    // Get HTML body from the client.
    mailboxItem.body.getAsync("html",
        { asyncContext: event },
        function (getBodyResult) {
            if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                updateBody(getBodyResult.asyncContext, getBodyResult.value);
            } else {
                console.error("Failed to get HTML body.");
                getBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}
// Register the function.
Office.actions.associate("insertContosoMeeting", insertContosoMeeting);

// Settings function (placeholder for now)
function openSettings(event) {
    // For now, just show a notification that settings was clicked
    const message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Settings option clicked - functionality coming soon!",
        icon: "Icon.80x80",
        persistent: true,
    };

    Office.context.mailbox.item.notificationMessages.replaceAsync(
        "SettingsNotification",
        message
    );

    // Complete the event
    event.completed();
}
// Register the settings function
Office.actions.associate("openSettings", openSettings);

// 3. How to implement a supporting function `updateBody`
//    that appends the online meeting details to the current body of the meeting.
function updateBody(event, existingBody) {
    // Append new body to the existing body.
    mailboxItem.body.setAsync(existingBody + newBody,
        { asyncContext: event, coercionType: "html" },
        function (setBodyResult) {
            if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                setBodyResult.asyncContext.completed({ allowEvent: true });
            } else {
                console.error("Failed to set HTML body.");
                setBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}

// /*
//  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
//  * See LICENSE in the project root for license information.
//  */

// /* global Office */

// Office.onReady(() => {
//   // If needed, Office.js is ready to be called.
// });

// /**
//  * Shows a notification when the add-in command is executed.
//  * @param event {Office.AddinCommands.Event}
//  */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true,
//   };

//   // Show a notification message.
//   Office.context.mailbox.item.notificationMessages.replaceAsync(
//     "ActionPerformanceNotification",
//     message
//   );

//   // Be sure to indicate when the add-in command function is complete.
//   event.completed();
// }

// // Register the function with Office.
// Office.actions.associate("action", action);
