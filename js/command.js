Office.initialize = function () {};
function SetLogcation(event) {
  let meetingLink = localStorage.getItem("meetngLink");
  if (meetingLink) {
    setLocation(meetingLink);
  } else {
    statusUpdate("errorMessage", "Please configure meeting link first");
  }
  event.complete();
}
function setLocation(locationValue) {
  item.location.setAsync(
    locationValue,
    { asyncContext: { var1: 1, var2: 2 } },
    function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        statusUpdate("errorMessage", "Unable to insert location");
      } else {
        statusUpdate("informationalMessage", "Location added successfully");
      }
    }
  );
}
// Helper function to add a status message to the info bar.
function statusUpdate(actionType, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: actionType,
    //icon: icon,
    message: text,
    persistent: false,
  });
}
