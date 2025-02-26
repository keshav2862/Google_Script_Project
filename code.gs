// Form URL and Event Id as key for verification
var FORM_URL = "https://docs.google.com/forms/d/e/......./viewform";
var EVENT_ID_KEY = "entry.839337160";
function handleCalendarEvent(e) {
try {
// Error handling for missing Calendar Event
if (!e || !e.calendarEvent) {
throw new Error("Missing 'calendarEvent' or not triggered
properly.");
}
var event = e.calendarEvent;
var eventId = event.getId();
// I am going to choose participant A as the active user and
participant B as my secondary email for testing purposes. The participant B
will be chosen from a guest list just in case we need more emails.
var participantAEmail = Session.getActiveUser().getEmail();
var participantAName = "participantA";
var guestList = event.getGuestList();
if (guestList.length === 0) {
throw new Error("No guest found in the event.");
}
// To make it simple I just chose the first index of my guest list as
my participantB
var participantB = guestList[0];
var participantBEmail = participantB.getEmail();
var participantBName = "participantB";
// Now creating the folder with the given naming convention and I will
use the driveapp library to create the folder.
var folderName = participantAName + " " + participantBName + "
Handoff";
var folder = DriveApp.createFolder(folderName);
// Sharing the folder with participant A and saving the handoff
details.
folder.addEditor(participantAEmail);
var handoffData = {
folderId: folder.getId(),
participantAEmail: participantAEmail,
participantBEmail: participantBEmail
};
PropertiesService.getScriptProperties().setProperty("handoff_" +
eventId, JSON.stringify(handoffData));
// I am logging just to see if it executed
Logger.log("Folder created for the event: " + eventId);
// Creating a pre-populated form link to match the event id and
emailing to the participant B.
var prepopulatedLink = FORM_URL + "?" + EVENT_ID_KEY + "=" +
encodeURIComponent(eventId);
var emailSubject = "Please Complete the form given below:";
var emailMessage = "Dear " + participantBName + ",\n\n" +
"Please complete the form by clicking the link below:\n" +
prepopulatedLink + "\n\n" +
"Regards,\n\n" +
participantAName;
MailApp.sendEmail(participantBEmail, emailSubject, emailMessage);
// Logging and error checking
Logger.log("Pre-populated form link emailed to Participant B: " +
participantBEmail);
} catch (error) {
Logger.log("Error in handleCalendarEvent: " + error.message);
}
}
// I have created a Dummy function below for testing the above function.
For production purposes we have a simulate a proper calendar event using
the above function using a timer trigger.
function testHandleCalendarEvent() {
var dummyEvent = {
calendarEvent: {
getId: function () { return EVENT_ID_KEY; },
getGuestList: function () {
return [{
getEmail: function () { return "My secondary email"; }, // I used
my secondary email for testing (gmail.com)
getName: function () { return "ParticipantB"; }
}];
}
}
};
handleCalendarEvent(dummyEvent);
}
// Onsubmit function to get details from form. Just to let you that this
apps script project is binded to my test form.
function onFormSubmit(e) {
try {
//getting the item responses
var formResponse = e.response;
var itemResponses = formResponse.getItemResponses();
var responseText = "";
var eventId = null;
// Getting the text response for each question
itemResponses.forEach(function (itemResponse) {
var question = itemResponse.getItem().getTitle();
var answer = itemResponse.getResponse();
responseText += question + ": " + answer + "\n";
// Here we are looking for the event id to match with the event id
given in the testHandleCalendarEvent
if (question.toLowerCase().indexOf("event id") !== -1) {
eventId = answer;
Logger.log("Found event: " + eventId);
}
});
if (!eventId) {
throw new Error("Event ID not in the form please recheck");
}
// We get and retrieve the handoff details using event id and logging.
var handoffDataString =
PropertiesService.getScriptProperties().getProperty("handoff_" + eventId);
if (!handoffDataString) {
throw new Error("No data found for : " + eventId);
}
var handoffData = JSON.parse(handoffDataString);
Logger.log("Handoff data found: " + handoffDataString);
// Create a PDF and saving it in the folder using a helper function.
var pdfTitle = "Handoff Document between Participant A and Participant
B";
var pdfBlob = createPdfFromText(responseText, pdfTitle);
Logger.log("PDF created: " + pdfTitle);
var folder = DriveApp.getFolderById(handoffData.folderId);
folder.createFile(pdfBlob);
// Emailing both participants.
var subject = "Completion of the Handoff Procedure";
var message = "Hello participants ,\n\nThe handoff process has been
completed. The folder is saved in Participant A's Drive
Folder.\n\nRegards.";
MailApp.sendEmail(handoffData.participantAEmail, subject, message);
MailApp.sendEmail(handoffData.participantBEmail, subject, message);
Logger.log("Email sent.");
} catch (error) {
Logger.log("Error in onFormSubmit: " + error.message);
}
}
//Helper function to create pdf
function createPdfFromText(text, title) {
// Creting a google doc temporarily
var doc = DocumentApp.create(title);
var body = doc.getBody();
body.appendParagraph(text);
doc.saveAndClose();
// Converting the document to a PDF blob file.
var docFile = DriveApp.getFileById(doc.getId());
var pdfBlob = docFile.getAs("application/pdf").setName(title + ".pdf");
// Deleting the document later.
DriveApp.getFileById(doc.getId()).setTrashed(true);
return pdfBlob;
}
