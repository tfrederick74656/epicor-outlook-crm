﻿Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

Public Class EpicorOutlookCRM

	Dim cbCRM As Office.CommandBar = Nothing
	Dim btnCall As Office.CommandBarButton = Nothing
	Dim btnEmail As Office.CommandBarButton = Nothing

	Private Sub EpicorOutlookCRM_Startup() Handles Me.Startup
		'Start watching Outlook folders for new messages.
		'Fetch and cache information from the Progress/SQL database.
		'Create the CRM toolbar.
	End Sub

	Private Sub EpicorOutlookCRM_Shutdown() Handles Me.Shutdown
		'Stop watching Outlook folders for new messages.
		RemoveToolbar()
	End Sub

	' Create the CRM toolbar and related objects for application startup.
	Private Sub AddToolbar()
		' Verify the command bar and buttons don't already exist
		If (cbCRM Is Nothing And btnCall Is Nothing And btnEmail Is Nothing) Then
			' Create a new command bar and add it to Outlook
			Dim cmdBars As Office.CommandBars = Me.Application.ActiveExplorer().CommandBars
			cbCRM = cmdBars.Add("Epicor/Outlook CRM", Office.MsoBarPosition.msoBarTop, False, True)
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

End Class