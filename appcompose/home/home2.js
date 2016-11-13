(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      jQuery('#set-subject').click(setSubject);
      jQuery('#get-subject').click(getSubject);
      jQuery('#add-to-recipients').click(addToRecipients);
    });
  };

  function setSubject(){
    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync('Hello world!');
  }

  function getSubject(){
    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.getAsync(function(result){
      app.showNotification('The current subject is', result.value);
    });
  }

  function addToRecipients(){
    var item = Office.context.mailbox.item;
    var addressToAdd = {
      displayName: Office.context.mailbox.userProfile.displayName,
      emailAddress: Office.context.mailbox.userProfile.emailAddress
    };

    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
    }
  }

  // The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
	$(document).ready(function () {
		$('#set-data').click(writeText);
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



})();
