(function () {
  ("use strict");
  Office.onReady(function () {
    $(document).ready(function () {
      $("#submitbtn").click(SaveLink);
      //let meetingLink = localStorage.getItem("meetngLink");
      let settings = Office.context.roamingSettings;
      let meetingLink = settings.get("meetngLink");
      if (meetingLink) {
        $("#meetinglink").val(meetingLink);
      }
    });
  });
  function SaveLink() {
    let link = $("#meetinglink").val();
    //localStorage.setItem("meetngLink", link);
    let settings = Office.context.roamingSettings;
    settings.set("meetngLink", link);
    settings.saveAsync(saveMyAddInSettingsCallback);
  }
  function saveMyAddInSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      // Handle the failure.
    }
  }
})();
