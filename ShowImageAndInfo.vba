Sub ShowImageAndInfo()
    ' Make sure Tools > References includes:
    ' - Microsoft Visual Basic for Applications Extensibility 5.3
    ' - Microsoft Forms 2.0 Object Library

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
            ' Create the UserForm dynamically
            Dim VBComp As VBComponent
            Set VBComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
            With VBComp
                .Properties("Width") = 400
                .Properties("Height") = 600
                .Properties("Caption") = ws.Cells(rowIndex, 2).Value
            End With

            ' Add Image control
            Dim ImgControl As MSForms.Image
            Set ImgControl = VBComp.Designer.Controls.Add("Forms.Image.1")
            With ImgControl
                .Left = 10
                .Top = 10
                .Width = 380
                .Height = 300
                .Picture = LoadPicture(localFilePath)
            End With

            ' Add Labels for Description, Director, Writer, Cast
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
                .Locked = True ' makes it read-only
                .BackColor = RGB(240, 240, 240) ' optional styling
            End With

            Dim lblDirector As MSForms.Label
            Set lblDirector = VBComp.Designer.Controls.Add("Forms.Label.1")
            With lblDirector
                .Caption = "Director: " & ws.Cells(rowIndex, 5).Value
                .Left = 10
                .Top = 370
                .Width = 380
            End With

            Dim lblWriter As MSForms.Label
            Set lblWriter = VBComp.Designer.Controls.Add("Forms.Label.1")
            With lblWriter
                .Caption = "Writer: " & ws.Cells(rowIndex, 6).Value
                .Left = 10
                .Top = 420
                .Width = 380
            End With

            Dim txtCast As MSForms.TextBox
            Set txtCast = VBComp.Designer.Controls.Add("Forms.TextBox.1")
            With txtCast
                .Text = "Cast: " & ws.Cells(rowIndex, 8).Value
                .Left = 10
                .Top = 430
                .Width = 380
                .Height = 100
                .MultiLine = True
                .WordWrap = True
                .ScrollBars = fmScrollBarsVertical
                .Locked = True ' makes it read-only
                .BackColor = RGB(240, 240, 240) ' optional styling
            End With

            Dim btnLink As MSForms.CommandButton
            Set btnLink = VBComp.Designer.Controls.Add("Forms.CommandButton.1")
            With btnLink
                .Caption = "Open Link"
                .Left = 150
                .Top = 540
                .Width = 100
                .Height = 30
                .Tag = ws.Cells(rowIndex, 7).Value ' store the link in Tag property

                ' Styling polish
                .BackColor = RGB(0, 102, 204) ' deep blue background
                .ForeColor = RGB(255, 255, 255) ' white text
                .Font.Bold = True
                .Font.Size = 10
            End With
            
            ' Add click handler dynamically
            Dim CodeMod As CodeModule
            Set CodeMod = VBComp.CodeModule
            Dim LineNum As Long
            LineNum = CodeMod.CountOfLines + 1
            
            CodeMod.InsertLines LineNum, _
                "Private Sub " & btnLink.Name & "_Click()" & vbCrLf & _
                "    Dim link As String" & vbCrLf & _
                "    link = Me." & btnLink.Name & ".Tag" & vbCrLf & _
                "    If link <> """" Then" & vbCrLf & _
                "        ThisWorkbook.FollowHyperlink link" & vbCrLf & _
                "    Else" & vbCrLf & _
                "        MsgBox ""No link available."", vbExclamation" & vbCrLf & _
                "    End If" & vbCrLf & _
                "End Sub"

            ' Show the form and clean up afterwards
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