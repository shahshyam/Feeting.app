(function () {
  ("use strict");
  Office.onReady(function () {
    $(document).ready(function () {
      $("#submitbtn").click(SaveLink);
      let meetingLink = localStorage.getItem("meetngLink");
      if (meetingLink) {
        $("#meetinglink").val(meetingLink);
      }
    });
  });
  function SaveLink() {
    let link = $("#meetinglink").val();
    localStorage.setItem("meetngLink", link);
  }
})();
