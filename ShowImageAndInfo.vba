Sub ShowImageAndInfo()
	Dim btnName As String
	btnName = Application.Caller
	
	Dim ws As Worksheet
	Set ws = ThisWorkbook.Sheets("Sheet1")
	
	Dim rowIndex As Long
	rowIndex = ws.Shapes(btnName).TopLeftCell.Row
	
	Dim imageURL As String
	imageURL = ws.Cells(rowIndex, 9).Value
	
	If imageURL <> "N/A" Then
		Dim localFilePath As String
		localFilePath = DownloadImage(imageURL)
		
		If localFilePath <> "" Then
			
			Dim VBComp As VBComponent
			Set VBComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
			With VBComp
				.Properties("Width") = 400
				.Properties("Height") = 600
				.Properties("Caption") = ws.Cells(rowIndex, 2).Value
			End With
			
			' ---------------- IMAGE CONTROL ----------------
			Dim ImgControl As MSForms.Image
			Set ImgControl = VBComp.Designer.Controls.Add("Forms.Image.1")
			
			ImgControl.Picture = LoadPicture(localFilePath)
			ImgControl.PictureSizeMode = fmPictureSizeModeZoom
			ImgControl.Left = 10
			ImgControl.Top = 10
			ImgControl.Width = 380
			ImgControl.Height = 300
			
			' Store image path for full-screen viewer
			ImgControl.Tag = localFilePath
			
			' Add click handler for full-screen viewer
			Dim CodeMod As CodeModule
			Set CodeMod = VBComp.CodeModule
			Dim LineNum As Long
			LineNum = CodeMod.CountOfLines + 1
			
			CodeMod.InsertLines LineNum, _
				"Private Sub " & ImgControl.Name & "_Click()" & VbCrLf & _
				"    ShowFullScreenImage Me." & ImgControl.Name & ".Tag" & VbCrLf & _
				"End Sub"
			
			' ---------------- DESCRIPTION ----------------
			Dim txtDesc As MSForms.TextBox
			Set txtDesc = VBComp.Designer.Controls.Add("Forms.TextBox.1")
			With txtDesc
				.Text = "Description: " & ws.Cells(rowIndex, 3).Value
				.Left = 10
				.Top = 320
				.Width = 380
				.Height = 40
				.Multiline = True
				.WordWrap = True
				.ScrollBars = fmScrollBarsVertical
				.Locked = True
				.BackColor = RGB(240, 240, 240)
			End With
			
			' ---------------- DIRECTOR ----------------
			Dim lblDirector As MSForms.Label
			Set lblDirector = VBComp.Designer.Controls.Add("Forms.Label.1")
			With lblDirector
				.Caption = "Director: " & ws.Cells(rowIndex, 5).Value
				.Left = 10
				.Top = 370
				.Width = 380
			End With
			
			' ---------------- WRITER ----------------
			Dim lblWriter As MSForms.Label
			Set lblWriter = VBComp.Designer.Controls.Add("Forms.Label.1")
			With lblWriter
				.Caption = "Writer: " & ws.Cells(rowIndex, 6).Value
				.Left = 10
				.Top = 420
				.Width = 380
			End With
			
			' ---------------- CAST ----------------
			Dim txtCast As MSForms.TextBox
			Set txtCast = VBComp.Designer.Controls.Add("Forms.TextBox.1")
			With txtCast
				.Text = "Cast: " & ws.Cells(rowIndex, 8).Value
				.Left = 10
				.Top = 460
				.Width = 380
				.Height = 70
				.Multiline = True
				.WordWrap = True
				.ScrollBars = fmScrollBarsVertical
				.Locked = True
				.BackColor = RGB(240, 240, 240)
			End With
			
			' ---------------- LINK BUTTON ----------------
			Dim btnLink As MSForms.CommandButton
			Set btnLink = VBComp.Designer.Controls.Add("Forms.CommandButton.1")
			With btnLink
				.Caption = "Open Link"
				.Left = 150
				.Top = 540
				.Width = 100
				.Height = 30
				.Tag = ws.Cells(rowIndex, 7).Value
				.BackColor = RGB(0, 102, 204)
				.ForeColor = RGB(255, 255, 255)
				.Font.Bold = True
				.Font.Size = 10
			End With
			
			' Add link button handler
			LineNum = CodeMod.CountOfLines + 1
			CodeMod.InsertLines LineNum, _
				"Private Sub " & btnLink.Name & "_Click()" & VbCrLf & _
				"    Dim link As String" & VbCrLf & _
				"    link = Me." & btnLink.Name & ".Tag" & VbCrLf & _
				"    If link <> """" Then" & VbCrLf & _
				"        ThisWorkbook.FollowHyperlink link" & VbCrLf & _
				"    Else" & VbCrLf & _
				"        MsgBox ""No link available."", vbExclamation" & VbCrLf & _
				"    End If" & VbCrLf & _
				"End Sub"
			
			' Show form
			With VBA.UserForms.Add(VBComp.Name)
				.Show
				ThisWorkbook.VBProject.VBComponents.Remove VBComp
			End With
			
		Else
			MsgBox "Failed to download the image.", vbExclamation
		End If
	Else
		MsgBox "No image available for this movie.", vbExclamation
	End If
End Sub

Public Sub ShowFullScreenImage(imagePath As String)
	If Dir(imagePath) = "" Then Exit Sub
	
	Dim VBComp As VBComponent
	Set VBComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
	
	With VBComp
		.Properties("Width") = Application.Width
		.Properties("Height") = Application.Height
		.Properties("Caption") = "Full Screen Image"
	End With
	
	Dim Img As MSForms.Image
	Set Img = VBComp.Designer.Controls.Add("Forms.Image.1")
	
	With Img
		.Picture = LoadPicture(imagePath)
		.PictureSizeMode = fmPictureSizeModeZoom
		.Left = 0
		.Top = 0
		.Width = Application.Width
		.Height = Application.Height
	End With
	
	With VBA.UserForms.Add(VBComp.Name)
		.Show
		ThisWorkbook.VBProject.VBComponents.Remove VBComp
	End With
End Sub