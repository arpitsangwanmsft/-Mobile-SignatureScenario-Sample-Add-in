// JavaScript source code
// The initialize function must be run each time a new page is loaded
var context;
var mailbox;
var item;

Office.initialize = function (reason) {
	context = Office.context;
	mailbox = context.mailbox;
	item = mailbox.item;
}


// Function Name as mentioned in the manifest under launch event extension point
async function getMessageBody(event) {
	// To get "From" of current compose session that addin has launched on
	Office.context.mailbox.item.from.getAsync(function (result) {
		if (result.status !== Office.AsyncResultStatus.Succeeded) {
			console.log(JSON.stringify(result.error));
		} else {
			console.log(`Got from: ${JSON.stringify(result.value)}`);
		}
	});

	// To get Compose-Type(new mail, reply, or forward) and coercionType(HTML) of message
	Office.context.mailbox.item.getComposeTypeAsync(function (asyncResult) {
		if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
			console.log(
				"getComposeTypeAsync succeeded with composeType: " +
				asyncResult.value.composeType +
				" and coercionType: " +
				asyncResult.value.coercionType
			);
		} else {
			console.error(asyncResult.error);
		}
	});

	/* Gets a token identifying the user and the Office Add-in.
	 * The token is returned as a string in the asyncResult.value property.
	 */
	Office.context.mailbox.getUserIdentityTokenAsync(function (result) {
		if (result.status !== Office.AsyncResultStatus.Succeeded) {
			console.error(`Token retrieval failed with message: ${result.error.message}`);
		} else {
			console.log(`Got User Identity Token ${result.value}`);
		}
	});

	// UserProfile Properties
	console.log("UserDisplayName: " + Office.context.mailbox.userProfile.displayName);
	console.log("UserEmailAddress: " + Office.context.mailbox.userProfile.emailAddress);
	console.log("User time zone: " + Office.context.mailbox.userProfile.timeZone);

	// Conversation Id: 
	console.log("Conversation id: " + Office.context.mailbox.item.conversationId);

	// Diagnostics information to differentiate between ios/android/owa/win32/mac : 
	console.log("Host: " + Office.context.diagnostics.host);
	console.log("Platform the addin is running on: " + Office.context.diagnostics.platform);


	// Get-Set Custom Properties
	Office.context.mailbox.item.loadCustomPropertiesAsync(
		function (asyncResult) {
			console.log("Get/Set custom properties check");
			console.log("Initial custom property value " + JSON.stringify(asyncResult));
			var customProperties = asyncResult.value;
			customProperties.set("testKey", "testValue");
			customProperties.set("EventLogged", "true");

			customProperties.saveAsync(
				function callback(saveAsyncResult) {
					if (saveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
						console.log("SetCustomProperties succeeded with two fields EventLogged and testKey.")
						Office.context.mailbox.item.loadCustomPropertiesAsync(
							function (asyncResult) {
								console.log("Populated customProperties" + JSON.stringify(asyncResult))
							}
						)
					} else {
						console.error(`SetCustomProperties failed with message ${result.error.message}`);
					}
				})
		}
	);


	// To get Body of current message as HTML
	Office.context.mailbox.item.body.getAsync(
		"html",
		{ asyncContext: "This is passed to the callback" },
		function callback(result2) {
			if (result2.status == Office.AsyncResultStatus.Succeeded) {
				console.log(`Body as html: ${result2.value}`);
			} else {
				console.log('Failed');
			}
		}
	);

	// To get Body of current message as text
	Office.context.mailbox.item.body.getAsync(
		"text",
		{ asyncContext: "This is passed to the callback" },
		function callback(result) {
			if (result.status == Office.AsyncResultStatus.Succeeded) {
				console.log(`Body as text: ${result.value}`);
			} else {
				console.log('Failed');
			}

		}
	);

	// Get 'To' recipients of the message
	Office.context.mailbox.item.to.getAsync(function (res) {
		if (res.status !== Office.AsyncResultStatus.Succeeded) {
			console.log(res);
		} else {
			console.log(`Got To: ${JSON.stringify(res.value)}`);
		}
	});

	// Get 'cc' recipients of the message
	Office.context.mailbox.item.cc.getAsync(function (res) {
		if (res.status !== Office.AsyncResultStatus.Succeeded) {
			console.log(JSON.stringiy(res.error));
		} else {
			console.log(`Got CC: ${JSON.stringify(res.value)}`);
		}
	});

	// Get subject of the message
	Office.context.mailbox.item.subject.getAsync((result) => {
		if (result.status !== Office.AsyncResultStatus.Succeeded) {
			console.error(`Action failed with message ${result.error.message}`);
			return;
		}
		console.log(`Subject: ${result.value}`);
	});

	/*
	 * To disable native client Signature for current account
	 * Note: disabling in context of mobile implies clearing-off of the native signature because the disable  
	 * toggle is not available in mobile unlike other platforms.
	 * Therefore,Following scenarios could arise:
	 * Single-Account is added:
	 *		1) signature is cleared after api is called.

	 * Multiple-Accounts are added and per-account signature is enabled:
	 *		1) Clear the signature for current addin account.

	 * Multiple-Accounts are added but per-account signature is disabled:
	 *		1) Turn per-account Signature on 
	 *		2) Clear the native signature for current addin account.
	 *		3) For other Accounts: if per-account signature has ever been edited then retain it otherwise replace it with global signature 
	 */
	Office.context.mailbox.item.disableClientSignatureAsync(function (asyncResult) {
		if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
			console.log("disableClientSignatureAsync succeeded");
		} else {
			console.error(JSON.stringify(asyncResult.error));
		}
	});

	// Set Signature With Coersion type as HTML (signature will be treated as Html)
	Office.context.mailbox.item.body.setSignatureAsync(
		"<p><b>Here's a Signature, added by the Add-in!</b></p>",
		{
			"coercionType": "html"
		},
		function (asyncResult) {
			console.log(JSON.stringify(asyncResult));
		}
	);

	// Set Signature With Coersion type as Text (html tags will be appended as text to the signature)
	Office.context.mailbox.item.body.setSignatureAsync(
		"<p><b>Here's a Signature, added by the Add-in!</b></p>",
		{
			"coercionType": "text"
		},
		function (asyncResult) {
			console.log(JSON.stringify(asyncResult));
		}
	);

	// Add a notification message to the item.
	var test_id = "test_notif_id";
	Office.context.mailbox.item.notificationMessages.addAsync(
		test_id,
		{
			type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
			message: "Test Added"
		},
		function (asyncResult) {
			console.log(JSON.stringify(asyncResult));
		}
	);

	//Replace a notification on the current item
	Office.context.mailbox.item.notificationMessages.replaceAsync(
		test_id,
		{
			type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
			message: "Test II Added"
		},
		function (asyncResult) {
			console.log(JSON.stringify(asyncResult));
		}
	);

	//Remove the notification on the current item
	Office.context.mailbox.item.notificationMessages.removeAsync(
		test_id,
		function (asyncResult)
		{
			console.log(JSON.stringify(asyncResult));
		}
	);

	/* Adds a signature with an inline image attachment from a base64 encoded string
	 * Adding of file attachments is not supported yet (i.e. Only 'isInline:true' is supported).
	 * For Inline Attachments, exact name(for example: "taskpane.jpg" here) of the file passed in addFileAttachmentFromBase64Async
	 * has to be used inside the img tag like: <img src='cid:taskpane.jpg'> 
	 */
	Office.context.mailbox.item.addFileAttachmentFromBase64Async
		(	
			"/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAIBAQIBAQICAgICAgICAwUDAwMDAwYEBAMFBwYHBwcGBwcICQsJCAgKCAcHCg0KCgsMDAwMBwkODw0MDgsMDAz/2wBDAQICAgMDAwYDAwYMCAcIDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAz/wAARCAAgACADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD84P2P/wDgml4F/aB/Z08O+LtZ1bxba6lq32nzorK6t0gXy7mWJdoeBm+7GCcseSenSvS/+HNvwx/6Dvjz/wADbT/5Gr9kv+DQ7/lDbov/AGNWsf8Ao1K/T6vqMPnWXQpRhPBqTSSb5t2lq/h67nztbKcdOpKcMU4pttLl2XbfofyX/wDDm34Y/wDQd8ef+Btp/wDI1eaftgf8E0vAv7P37OniLxdo2reLbrUtJ+zeTFe3Vu8DeZcxRNuCQK33ZCRhhyB16V/Uh+0R4Y+KHi34v+KLX4S+I9F8L+JV0Tw7JPd6nEJIXtRca6HjAMMwDFzGc7Oinkd/zm/4ONvBXxs8Hf8ABGfx0PjF4w8OeLHuPFmgf2SdKgWP7KBJP52/bbwZ3Zixw2Np6Z56q2YZbUwc5ww8IyaaS53zJvS6XJbS991sctHA4+ni4wnXnKN02+VcrS1s3zX1226ns3/Bod/yht0X/satY/8ARqV+n1fydf8ABLv/AIOZvGn/AAS4/ZL074T+Hfhf4X8T2NlqF3qT6hqWpzxyyyTybiAkagKqgKMZJJBOeQB9D/8AEbz8U/8Aoh/w/wD/AAbXf+FfGn1h+wH7fXhj4X+LPGHiK1+LXiTWfC/hpdP8NyQXWmIzzSXQl8QBIyFhmJUoZT9zqo5Hf81f+C3Hgv4J+Dv+CMfxPX4OeMPEXixbjxb4aOrf2rG8ZtSJLrydga2g+9mXJAb7g6d/nnxX/wAHf/ij4geJtQ1LxF+zr8L/ABEuo2lnaPZapdy3lpH9le7eKRI5IziT/TJgWyeMAY5z4P8A8FBf+Dgm4/bw/ZI1z4Sw/Af4a/DPT9c1Gy1OW/8ADBNvI8trIWQSIIwJF2vIBkggtkHqD60cwisA8Led272uuTdPVWv077nmSwMnjVibRslvZ82zW97fhsf/2Q==",
			"taskpane.jpg",
			{
				"isInline": true,
				"asyncContext": { foo: 6, bar: 28 }
			},
			function (asyncResult) {
				console.log("Inline Attachment Added: " + JSON.stringify(asyncResult));
				// To add a signature containing the image that was added as an inline attachment by the previous api call
				Office.context.mailbox.item.body.setSignatureAsync(
					"<p>Here's an image in Signature, added as inline in base64 format !</p><td style='border-right: 1px solid #000000; padding-right: 5px;'><a href='http://www.google.com/'><img width='100%' src='cid:taskpane.jpg'></a></td>",
					{
						"coercionType": "html"
					},
					function (asyncResult2) {
						console.log(JSON.stringify(asyncResult2));
						/*
						* To Add a notification with custom message
						* Note: Only one type of  default "icon" and "type" is supported. Using any other type/icon will default down to the supported type/icon. 
						*/
						var notif_id = "my_notif_id_1";
						Office.context.mailbox.item.notificationMessages.addAsync(
							notif_id,
							{
								type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
								message: "Signature Added"
							},
							function (asyncResult) {
								console.log(JSON.stringify(asyncResult));
							}
						);

						// Get all notifications added by the addin on the current item
						Office.context.mailbox.item.notificationMessages.getAllAsync(
							function (asyncResult) {
								console.log(JSON.stringify(asyncResult));
							}
						);
						event.completed({ allowEvent: true });
					});
			});
}