Attribute VB_Name = "PublicRoutines"
Global rs As ADODB.Recordset
Public cnn As New Connection
Global cn As ADODB.Connection
Public cno As Long
Public loadOno As Long
Public LoadTime As Boolean


Public Function CenterWindow(frm As Form)
   ' frm.Left = (mdiMain.ScaleWidth - frm.Width) / 2
    'frm.Top = (mdiMain.ScaleHeight - frm.Height) / 2
End Function

Public Function OpenConnection() As Boolean
On Error GoTo HandleError
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.JET.OLEDB.4.0;Data source =" & App.Path & "\restaurant.mdb"
    cn.Open
    OpenConnection = True
    Exit Function
HandleError:
    OpenConnection = False
End Function

Public Function CloseConnection()
    cn.Close
    Set cn = Nothing
End Function

Public Function Query()
    Select Case UnloadMode
            Case vbFormCode
            Case vbFormControlMenu
            Case vbFormMDIForm
                MsgBox "First close the form"
                Cancel = True
            Case vbAppTaskManager
                MsgBox "First close the application then shut Down"
                Cancel = True
            Case vbAppWindows
                MsgBox "First close the Application"
                Cancel = True
    End Select
    
 End Function
            
            

