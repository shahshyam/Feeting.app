Office.initialize = function () {};
function SetLogcation(event) {
  let meetingLink = localStorage.getItem("meetngLink");
  if (meetingLink) {
    setLocation(meetingLink);
    addTextOndescription();
  } else {
    statusUpdate("errorMessage", "Please configure meeting link first");
  }
  event.completed();
}
function setLocation(locationValue) {
  let item = Office.context.mailbox.item;
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
function addTextOndescription() {
  let item = Office.context.mailbox.item;
  let body =
    "<p><br/><br/>___________________________________________________________________ <br>" +
    "Cool! You are invited to an audio-only, hands-free, walking meeting via <a href='http://feeting.app/'>feeting.app</a>." +
    "<br/>At the given moment, just put in a pair of airbuds or headphones and, from your phone. " +
    "Press the link and follow the flow, you'll automatically be taken to your feeting. " +
    "<br/><br/>No worries if you're not able to walk, you can join from your desktop as well. " +
    "Feeting is the #1 walking meeting platform 🚶👣 🤙 </p>";
  item.body.getTypeAsync(function (result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      if (result.value == Office.MailboxEnums.BodyType.Html) {
        item.body.setSelectedDataAsync(
          body,
          {
            coercionType: Office.CoercionType.Html,
            asyncContext: { var3: 1, var4: 2 },
          },
          function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
              statusUpdate(
                "informationalMessage",
                "Location added successfully"
              );
            } else {
              statusUpdate("errorMessage", "Unable to insert location");
            }
          }
        );
      } else {
        item.body.setSelectedDataAsync(
          body,
          {
            coercionType: Office.CoercionType.Text,
            asyncContext: { var3: 1, var4: 2 },
          },
          function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
              statusUpdate(
                "informationalMessage",
                "Location added successfully"
              );
            } else {
              statusUpdate("errorMessage", "Unable to insert location");
            }
          }
        );
      }
    }
  });
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
