Attribute VB_Name = "CentralizedErrorMod"
Option Explicit

'***Centralized Error handling/trapping***'
Public Sub CentralErrhandler(Optional ErrorSub As String)
Select Case Err.Number
    Case -2147217843
        MsgBox Err.Description & vbCrLf & "Check your UserName and/or Password"
    Case -2147467259
        MsgBox Err.Description & vbCrLf & _
                "The specified server is not online." & vbCrLf & _
                "Or the adress is wrong." & vbCrLf & _
                "Or the Port is wrong."
    Case -2147217900 'Wrong syntax custom T-SQL
        MsgBox Err.Description
    Case -2147217911
        MsgBox Err.Description
    Case -2147217904 ' Stored procedure expects parameters that was not provided
        MsgBox Err.Description
    Case 3021
        MsgBox "No Records!"
    Case 3709
        'Do nothing
    Case 3704
        'Do nothing
    Case Else
        If Err.Number <> 0 Then
            MsgBox Err.Description
'            If MsgBox("Do you want to report an error ?", vbYesNo) = vbYes Then
'                     WebEmailOpen ("mailto:knoton@hotmail.com?subject=Bug report Database EditorViewer&body=The Error Description is:" & Err.Description & " " & _
'                     "The source of the error is: " & Err.Source & " " & _
'                     "The Sub it happen is: " & ErrorSub)
'            End If
         End If
End Select
    frmDB.StatusBar1.Panels(4).Text = "ERROR"
End Sub
