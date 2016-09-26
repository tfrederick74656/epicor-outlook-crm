Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Core

Public Class EpicorOutlookCRM

	Private Const gstrAllAccountsSyncObjectName As String = "All Accounts"
	Private Const gintTimerInterval As Integer = 1 * (60 * 1000) ' Convert minutes to milliseconds

	Dim cbCRM As CommandBar = Nothing
	Dim btnCall As CommandBarButton = Nothing
	Dim btnEmail As CommandBarButton = Nothing
	Dim lastSync As DateTime = Date.MinValue
	Dim syncTimer As Timer = Nothing

	Private Sub EpicorOutlookCRM_Startup() Handles Me.Startup
		AddToolbar()
		'StartWatchingFolders() 'Folders are automatically watched by their respective event handlers.
	End Sub

	Private Sub HandleStartup()
		MsgBox("Outlook Startup")
	End Sub

	Private Sub HandleSync()
		MsgBox("Sync Start/End")
	End Sub

	Private Sub HandleSyncProgress(ByVal State As Outlook.OlSyncState, ByVal Description As String, ByVal Value As Long, ByVal Max As Long)
		MsgBox("Sync Progress")
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
		syncTimer.Stop()
		syncTimer.Dispose()
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

	'Start Outlook
	'Register a handler for AppFolder SyncEnd, NewMailEx, and ItemSend
	'Start a manual sync. We don't know when the initial sync ends, since it doesn't throw an event but manually started sync events won't start/finish until after the initial sync-
	'-and they do throw a SyncEnd event, so we can immediately manually start a sync, wait for it to finish, and then queue off of that SyncEnd event.
	'When the sync finishes, start a method to handle the initial sync and/or messages received offline
	' Unregister the SyncEnd handler
	' Record the current time as currentSync
	' Scan through all stores for messages received that match the criteria described and process those that do
	' This could take a while is the user has a lot of new mesages, so queue any other events until this process is completed
	' Set lastSync = currentSync since we've now processed everything received prior to the currentSync time
	' Start the SyncEvent timer
	' Dequeue and process any events received during the previous timeframe.
	' Exit the initialSync process
	'Handle timerTick, NewMailEx, and ItemSend as usual
	' timerTick triggers a full folder scan, NewMailEx searches only the inbox, ItemSend works on a single event at time

	Private Sub SetupSync()
		Dim syncObject As Outlook.SyncObject = GetAllAccountsSyncObject()
		AddHandler syncObject.SyncEnd, AddressOf MainSyncLogic
		syncObject.Start()
	End Sub

	' Set up a timer used to trigger periodic scans of all Outlook folders
	Private Sub timerSetup()
		syncTimer = New Timer
		AddHandler syncTimer.Tick, AddressOf HandleTickEvent
		With syncTimer
			.Interval = gintTimerInterval
			.Start()
		End With
	End Sub

	Private Sub HandleTickEvent(ByVal Sender As Object, ByVal e As EventArgs)

	End Sub

	' Return a SyncObject representing the "All Accounts" send/receive group in Outlook.
	Private Function GetAllAccountsSyncObject() As Outlook.SyncObject
		' IMPORTANT: SyncObjects.Count and SyncObjects.Item are 1-indexed! By default, 1 = "All Accounts", 2 = "Application Folders"; there is no 0. These can be modified by the user!
		' First, check to make sure there is at least one SyncObject. If not, something is very wrong; throw an exception.
		If Application.Session.SyncObjects.Count > 0 Then
			' Next, enumerate through the list of SyncObjects looking for a name matching gstrAllAccountsSyncObjectName. This is the "All Accounts" group.
			For itemNum As Integer = 1 To Application.Session.SyncObjects.Count
				If Application.Session.SyncObjects.Item(itemNum).Name Is gstrAllAccountsSyncObjectName Then
					Return Application.Session.SyncObjects.Item(itemNum)
				End If
			Next
			' If no SyncObject has a matching name, throw an exception.
			' We could continue With an arbitrary SyncObject given it's just used for timing, but that assumes no one else will modify that usage in the future.
			Throw New Exception("No SyncObject with name '" + gstrAllAccountsSyncObjectName + "' exists. Please contact IT and verify that send/receive groups have not been modified.")
		Else
			Throw New Exception("No SyncObjects exists. Please contact IT and verify that send/receive groups have not been modified.")
		End If
	End Function

	Private Sub MainSyncLogic()

		' What calls this?
		' ItemSend
		' NewMailEx
		' FolderAdd
		' SyncEnd
		' Timer


		If lastSync = Date.MinValue Then
			' Initial Sync
			RemoveHandler Application.Session.SyncObjects.Item()
		End If
	End Sub

	Private Sub PlaceholderSub()
		MsgBox("DEVELOPMENT USE ONLY!")
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
	Private Sub HandleIncomingMail(ByVal EntryIDCollection As String) Handles Application.NewMailEx
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
