Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

Public Class EpicorOutlookCRM

	Dim cbCRM As Office.CommandBar = Nothing
	Dim btnCall As Office.CommandBarButton = Nothing
	Dim btnEmail As Office.CommandBarButton = Nothing
	Dim lastSync As System.DateTime = Date.MinValue

	Private Sub EpicorOutlookCRM_Startup() Handles Me.Startup
		AddToolbar()
		'StartWatchingFolders() 'Folders are automatically watched by their respective event handlers.
		StartInitialSync()
	End Sub

	Private Sub EpicorOutlookCRM_Shutdown() Handles Me.Shutdown
		RemoveToolbar()
		StopWatchingFolders()
	End Sub

	' Create the CRM toolbar and related objects for application startup.
	Private Sub AddToolbar()
		' Verify the command bar and buttons don't already exist
		If (cbCRM Is Nothing And btnCall Is Nothing And btnEmail Is Nothing) Then
			' Create a new command bar and add it to Outlook
			cbCRM = Application.ActiveExplorer().CommandBars.Add("Epicor/Outlook CRM", Office.MsoBarPosition.msoBarTop, False, True)
			' Create a new button for CRM calls, add it the command bar, and specify a click event handler.
			btnCall = CType(cbCRM.Controls.Add(1), Office.CommandBarButton)
			With btnCall
				.Caption = "CRM Call"
				.Picture = getImage(My.Resources.CallIcon)
				.Style = Office.MsoButtonStyle.msoButtonIconAndCaption
				.Tag = "buttonCall"
			End With
			AddHandler btnCall.Click, AddressOf HandleToolbarButtonClick
			' Create a new button for CRM emails, add it the command bar, and specify a click event handler.
			btnEmail = CType(cbCRM.Controls.Add(1), Office.CommandBarButton)
			With btnEmail
				.Caption = "CRM Email"
				.Picture = getImage(My.Resources.EmailIcon)
				.Style = Office.MsoButtonStyle.msoButtonIconAndCaption
				.Tag = "buttonEmail"
			End With
			AddHandler btnEmail.Click, AddressOf HandleToolbarButtonClick
			cbCRM.Visible = True
		End If
	End Sub

	' Remove the CRM toolbar and related objects for application shutdown.
	Private Sub RemoveToolbar()
		' Set buttons and command bar invisible.
		btnCall.Visible = False
		btnEmail.Visible = False
		cbCRM.Visible = False
		' Remove click event handler from buttons.
		RemoveHandler btnCall.Click, AddressOf HandleToolbarButtonClick
		RemoveHandler btnEmail.Click, AddressOf HandleToolbarButtonClick
		' Delete buttons and command bar.
		btnCall.Delete()
		btnEmail.Delete()
		cbCRM.Delete()
		' Clean up.
		btnCall = Nothing
		btnEmail = Nothing
		cbCRM = Nothing
	End Sub

	' Handle click events for buttons on the CRM toolbar.
	Private Sub HandleToolbarButtonClick(ByVal button As Office.CommandBarButton, ByRef Cancel As Boolean)
		MsgBox("Button clicked: " & button.Tag)
		'Databases.ValidateProjectID("99-99-999")
	End Sub

	' TODO: I think this can be cleaned up and/or simplified further
	' Convert Icon resources to IPictureDisp for display on command bar buttons.
	Private Function getImage(ByVal icon As Icon) As stdole.IPictureDisp
		Dim image As stdole.IPictureDisp = Nothing
		Try
			Dim imageList As New ImageList
			imageList.Images.Add(icon)
			image = ConvertImage.Convert(imageList.Images(0))
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try
		Return image
	End Function

	' Remove the event handlers for incoming and outgoing messages.
	Private Sub StopWatchingFolders()
		RemoveHandler Application.NewMailEx, AddressOf HandleIncomingMail
		RemoveHandler Application.ItemSend, AddressOf HandleOutgoingMail
	End Sub


	' A little background on what's going on here with the incoming mail event handling.
	' The Application.NewMail And Application.NewMailEx Event handlers are unreliable.
	' They do reliably fire when a new message is received, but they fail to provide notification-
	' -when multiple messages are received simultaneously. They also do not fire when received-
	' -mail is synchronized during application startup. This makes them essentially useless-
	' -as a primary means of detecting new mail, but still useful as a means to start another-
	' -event handler. Ex. we can rely on NewMailEx to fire when at least one message is received,-
	' -but we can't tell for sure how many messages were actually received. We can, however-
	' -use this to kick off a scan of critical folders to check for new messages.

	' Wait until initial synchronization has finished. Save the date/time.
	'	Scan all messages currently in Outlook received before the above date/time, and pick out all which are unread. Validate those.
	' Start NewMailEx event handler. Start SyncEnd event handler.
	'	When NewMailEx fires, grab the EntityID, and start a scan for any messages with a received date later than the processing date/time.
	'   -Get the EntityIDs for those as well. Save the date/time received for the newest message processed
	'	When SyncEnd fires, start a scan for any messages with a received date later than the last processing date/time
	'	-Get the EntityIDs for any of those as well.

	' Handle messages synchronized when Outlook first opens
	' Handle messages received when Outlook is running

	Public Sub SyncEvent()
		'Dim dispatcherTimer As New System.Threading.Timer
		'AddHandler dispatcherTimer.Tick, AddressOf dispatcherTimer_Tick
		'dispatcherTimer.Interval = New TimeSpan(0, 0, 1)
		'dispatcherTimer.Start()
	End Sub

	'Handle messages received between end of previous sync and end of the current sync
	Private Sub StartInitialSync()
		AddHandler Application.Session.SyncObjects.AppFolders.SyncEnd, AddressOf HandleSyncEvent
		'Initial sync doesn't throw this event
		Application.Session.SyncObjects.AppFolders.Start()
	End Sub

	'Handle the message received before the first manual sync completed
	Private Sub HandleSyncEvent()

		' If this is the first time we're running the sync event handler (initial manual sync), register the NewMailEx event handler.
		If lastSync = Date.MinValue Then
			'AddHandler Application.NewMailEx, AddressOf HandleIncomingMail
		End If

		Dim currentSync As Date = Now

		MsgBox("Sync Completed! Last: " + lastSync.ToString + " Current: " + currentSync.ToString)

		'If message.received > lastSync && message.received < currentSync


		' Scan all folders for new items
		' -For each store...
		' --For each folder...
		' ---For each message...
		' ----Is item an email message
		' ----Is received > firstSync
		' ----Is unread
		' ----Does contain a project ID?
		' ----Process the message

		lastSync = currentSync
	End Sub

	' TODO: Clean this up
	' Separate a list of message Entry
	Private Function SeparateEntryIDCollection(ByVal entryIDCollection As String) As Queue(Of String)
		Dim queue As New Queue(Of String)
		For Each entryID As String In entryIDCollection.Split(New Char() {","c})
			queue.Enqueue(entryID)
		Next
		Return queue
	End Function

	' Handle the event generated when a message is received.
	Private Sub HandleIncomingMail(ByVal EntryIDCollection As String)
		' Need to handle multiple items
		Dim item As System.Object = Application.Session.GetItemFromID(EntryIDCollection)
		MsgBox(Len(EntryIDCollection).ToString)
		If (TypeOf Application.Session.GetItemFromID(EntryIDCollection) Is Outlook.MailItem) Then
			Dim message As Outlook.MailItem = CType(item, Outlook.MailItem)
			MsgBox("Item Received: " + message.Subject + " ID: " + EntryIDCollection)
		End If
	End Sub

	' Handle the event generated when a message is sent.
	Private Sub HandleOutgoingMail(ByVal Item As System.Object, ByRef Cancel As Boolean) Handles Application.ItemSend
		Dim message As Outlook.MailItem = CType(Item, Outlook.MailItem)
		MsgBox("Item Sent: " + message.Subject)
	End Sub

End Class
