// This function is called when Office.js is ready to start your Add-in
Office.initialize = function (reason) { 
	$(document).ready(function () {
		$('#add-to-recipients').click(addToRecipients);
	});
}; 

// Adds the current user to the recipient list
function addToRecipients() {
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