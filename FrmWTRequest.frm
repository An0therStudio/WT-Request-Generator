VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmWTRequest 
   Caption         =   "Scheduled Walkthrough Request"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   OleObjectBlob   =   "FrmWTRequest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmWTRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSend_Click()
Dim Ns As Outlook.NameSpace
Dim strHeader As String
Dim EmailTable As String
Dim strFooter As String
Dim emailBody As String
Dim strDate As String
Dim arrDate() As String
Dim strTime As String
Dim intStart As Integer
Dim blComplete As Boolean
Dim blValidDate As Boolean
Dim blValidTime As Boolean
Dim arrUserName() As String
Dim strUserName As String

On Error Resume Next

Set Ns = Application.GetNamespace("MAPI")
Ns.Logon
'Create user name from object
arrUserName = Split(Ns.CurrentUser.Name, ",")
If UBound(arrUserName) >= 1 Then
    strUserName = Trim(arrUserName(1)) + " " + Trim(arrUserName(0))
Else
    strUserName = Trim(arrUserName(0))
End If
blComplete = False
'Validate Input
If Trim(txtCid.Value & vbNullString) = vbNullString Or _
        Trim(txtName.Value & vbNullString) = vbNullString Or _
        Trim(txtTel.Value & vbNullString) = vbNullString Or _
        Trim(txtDate.Value & vbNullString) = vbNullString Or _
        Trim(txtTime.Value & vbNullString) = vbNullString Then
    blComplete = False
'Input is complete
Else
'Remove extra spaces
    txtCid.Value = Trim(txtCid.Value)
    txtName.Value = Trim(txtName.Value)
    txtTel.Value = Trim(txtTel.Value)
    txtDate.Value = Trim(txtDate.Value)
    txtTime.Value = Trim(txtTime.Value)
    
'Check for valid date. Must be numeric.
    strDate = txtDate.Value
    If IsDate(strDate) Then
        strDate = Format(txtDate.Value, "mm/dd/yyyy")
        arrDate() = Split(strDate, "/")
'Check that date parts are in range and that no weird values were assigned by date format.
        If CInt(arrDate(0)) >= 1 And CInt(arrDate(0)) <= 12 _
          And CInt(arrDate(1)) >= 1 And CInt(arrDate(1)) <= 31 _
          And CInt(arrDate(2)) >= 2016 And CInt(arrDate(2)) <= 2050 Then
            blValidDate = True
        End If
    Else
        blValidDate = False
    End If
    If blValidDate = False Then
        MsgBox ("Invalid Date: " + strDate)
    End If
'Check for valid time
    strTime = txtTime.Value
    'Look for colon
    intStart = InStr(strTime, ":")
    'If colon not found (Input was HHMM)
    If intStart = 0 Then
        'Check for garbage input
        If IsNumeric(strTime) Then
            'If input is valid format time to HH:MM
            If Len(strTime) = 4 Then
                'strTime = "0" + Left(strTime, 1) & ":" & Right(strTime, 2)
            'ElseIf Len(strTime) = 4 Then
                strTime = Left(strTime, 2) & ":" & Right(strTime, 2)
            End If
        End If
    'colon is found (Input was HH:MM or H:MM)
    Else
        'Check for garbage input
        If Len(Mid(strTime, intStart + 1)) = 2 And IsNumeric(Mid(strTime, intStart + 1)) Then
            If IsNumeric(Left(strTime, intStart - 1)) Then
                'If format is H:MM add extra zero. If format is already HH:MM no work needed
                If Len(Left(strTime, intStart - 1)) = 1 Then
                    strTime = "0" & strTime
                End If
            End If
        End If
    End If
    
    'Final check for valid input.
    If IsDate(strTime) Then
        blValidTime = True
    Else
        blValidTime = False
        MsgBox ("Invalid Time: " + strTime)
    End If
    
    If blValidDate And blValidTime Then
        blComplete = True
    Else
        blComplete = False
    End If
End If

'Input is valid
If blComplete Then

'Add header
    strHeader = "<p>Hello,"
    strHeader = strHeader & "<br><br>"
    strHeader = strHeader & "Please schedule a WT block "
    If cbSelf.Value Then strHeader = strHeader & "<span style='background-color:#40ff00'>for " & strUserName & "</span>"
    strHeader = strHeader & " with the below information:"
    strHeader = strHeader & "<br><br></p>"

' Create the Table
    EmailTable = "<table style='width:420px;border:2px solid black;border-collapse: collapse;'><tr><td style='text-align:center;border:1px solid black;padding:15px;background-color:#fcf010' colspan=2><span style='font-size:125%;font-weight:bold;font-family:cambria;'>Scheduled Walkthrough Request</span></td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-weight:bold;font-family:cambria;'>Conference ID</span></td><td style='border:1px solid black;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-family:cambria;'>"
    EmailTable = EmailTable & txtCid.Value
    EmailTable = EmailTable & "</span></td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#d7f1ff'><span style='font-weight:bold;font-family:cambria;'>Contact Name</span></td><td style='border:1px solid black;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#d7f1ff'><span style='font-family:cambria;'>"
    EmailTable = EmailTable & txtName.Value
    EmailTable = EmailTable & "</span></td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-weight:bold;font-family:cambria;'>Contact Number</span></td><td style='border:1px solid black;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-family:cambria;'>"
    EmailTable = EmailTable & txtTel.Value
    EmailTable = EmailTable & "</span></td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#d7f1ff'><span style='font-weight:bold;font-family:cambria;'>Scheduled Date</span></td><td style='border:1px solid black;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#d7f1ff'><span style='font-family:cambria;'>"
    EmailTable = EmailTable & strDate
    EmailTable = EmailTable & "</span></td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-weight:bold;font-family:cambria;'>Scheduled Time (ET)</span></td><td style='border:1px solid black;padding-left:15px;padding-top:10px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-family:cambria;'>"
    EmailTable = EmailTable & Format(strTime, "hh:mm")
    EmailTable = EmailTable & "</span></td></tr></table>"
    EmailTable = EmailTable & "</br><small><i>**Request Received Via "
    
'Determine request origin
    If cbPhone.Value Then
        EmailTable = EmailTable & "Phone**</i></small>"
    Else
        EmailTable = EmailTable & "Email**</i></small>"
    End If
    
'Add footer
    strFooter = "<br><br>Thank you very much,<br>" & strUserName
    
'Create Email
    HTMLBody = "<html><head>"
    HTMLBody = HTMLBody & "</head><body>"
    HTMLBody = HTMLBody & strHeader
    HTMLBody = HTMLBody & "<br>" & EmailTable
    HTMLBody = HTMLBody & strFooter
    HTMLBody = HTMLBody & "</body></html>"
    
 'Send Email
    Dim objMsg As MailItem
    Set objMsg = Application.CreateItem(olMailItem)
    With objMsg
      .To = "PODCanada@west.com"
      '.To = "walkthroughsupport@teleconferencingcenter.com"
      '.To = "lcree@west.com"
      .Subject = "Scheduled WT Request for " & txtCid
      .SentOnBehalfOfName = "walkthroughsupport@teleconferencingcenter.com"
      .BodyFormat = olFormatHTML
      .HTMLBody = HTMLBody
      .Save
      .Send
    End With
    Unload Me
    Set objMsg = Nothing
    
'Input is bad
Else
    a = MsgBox("Please complete form before sending!", vbCritical, "Error!")
End If

Ns.Logoff
Set Ns = Nothing
End Sub

'Authored by Luke Cree
'An0therStudio
'info@an0therstudio.com
