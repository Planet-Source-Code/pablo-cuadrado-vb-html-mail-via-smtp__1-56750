<div align="center">

## VB HTML Mail via SMTP


</div>

### Description

It's a function that allows you to send an e-mail with both HTML and Plain Text formats, using the Winsock control. It's an enhancement to Brian Anderson's code.
 
### More Info
 
SourceForm, DestAddress, Server, Optional BodyHTML, Optional BodyTXT, Optional SenderName, Optional SenderAddress, Optional DestName, Optional Subject.

You gotta have a Winsock control or create an instance of it.

SMTP Errors, status codes.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Pablo Cuadrado](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pablo-cuadrado.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pablo-cuadrado-vb-html-mail-via-smtp__1-56750/archive/master.zip)

### API Declarations

```
' A global variable.
Public Response As String
```


### Source Code

```
' Did you declare the global variable:
' Public Response As String ?
Dim Start As Single, Tmr As Single
Public Sub HTMLMail(Server As String, SourceForm As Form, DestAddress As String, BodyHTML As String, Optional BodyTXT As String = "", Optional SenderName As String = "", Optional SenderAddress As String = "", Optional DestName As String = "", Optional Subject As String = "")
' HTMLMail
' by Pablo Cuadrado - Argentina
' Estudio Quadra - Innovating the Internet
'
' Created on Friday, October 15th, 2004.
'
' Uses Winsock object to connect to a SMTP server.
'
' I have seen a lot of answers on how to do more
' than sending a plain text mail on a code posted
' by Brian Anderson. Well, this is is a Multipart
' mail, so you can send even more things.
'
' I've made a class which allows you to create
' multipart mails, contact me if you wish to have
' it. This is just a simple FUNCTION, that allows
' a "Bi-part" mail, with both a plain text and a
' HTML message embedded.
'
' By getting the right MIME types, you can embed
' anything (pics, files, etc.) on an e-mail.
'
' There is a SourceForm parameter:
' you can call the function in a form with a Winsock
' control, just by adding, for instance:
' ...Command1_Click ()
' HTMLMail "smtp.myserver.com", Me, ... and so on.
'
' The keyword "Me" is the form object itself.
' I did this in a project with more than one Winsock control.
' You may delete that parameter, and then in the line:
' With SourceForm.Winsock
' Just specify wich control will you use.
Dim Header(40) As String
Dim i As Integer
Dim StatusOutput As String
Dim Headers As String
Dim MIMEDate As String, MIMEHeaders As String
With SourceForm.SCWinsock
 If .State = sckClosed Then
 MIMEDate = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
 Header(1) = "mail from: " & Chr(32) & SenderAddress & vbCrlf
 Header(2) = "rcpt to: " & DestAddress & vbCrlf
 Header(3) = "Date: " & MIMEDate & vbCrlf
 Header(4) = "From: """ & SenderName & """ <" & SenderAddress & ">" + vbCrlf
 Header(5) = "To: " & DestName & vbCrlf
 Header(6) = "Subject: " & Subject & vbCrlf
 Header(7) = "MIME-Version: 1.0" & vbCrlf
 Header(8) = "Content-Type: multipart/alternative;" & vbCrlf
 ' Here is the trick: you make a string (boundary) that divides the parts.
 Header(9) = " boundary = " & Chr(34) & "----=Division" & Chr(34) & ";" & vbCrlf
 Header(10) = "X-Mailer: YourApp" & vbCrlf
 ' The order for the headers:
 ' From - Date - MimeHeaders - X-Headers - To - Subject
 MIMEHeaders = Header(7) & Header(8) & Header(9)
 Headers = Header(4) & Header(3) & MIMEHeaders & Header(10) & Header(5) & Header(6)
 ' Plain Text Part
 ' ===============
 '
 ' The division goes with the prefix "--"
 ' Many programs uses strings starting with "-" to make a visible line.
 ' M$ Outlook does.
 Header(11) = "------=Division"
 Header(12) = "Content-Type: text/plain;"
 Header(13) = " charset = " & Chr(34) & "iso-8859-1" & Chr(34) & ";"
 Header(14) = vbCrlf & vbCrlf
 Header(15) = BodyTXT & vbCrlf ' Cuerpo
 ' HTML Text Part
 ' ==============
 Header(16) = "------=Division"
 Header(17) = "Content-Type: text/html;"
 Header(18) = " charset = " & Chr(34) & "iso-8859-1" & Chr(34)
 Header(19) = "Content-Transfer-Encoding: quoted-printable" & vbCrlf
 ' Remove the header to ensure HTML compatibility.
 'Header(19) = vbCrlf
 Header(20) = BodyHTML & vbCrlf ' Cuerpo
 ' The last division hast both an "--" prefix, and a "--" suffix.
 Header(21) = "------=Division--" & vbCrlf
 .LocalPort = 0
 .Protocol = sckTCPProtocol
 .RemoteHost = Server
 .RemotePort = 25
 .Connect
 WaitFor ("220")
 StatusOutput = "Connecting..."
 ' Whenever there's an StatusOutput, you could
 ' point it to a text or label on your app to
 ' create a visible status.
 .SendData ("HELO " & Server & vbCrlf)
 WaitFor ("250")
 StatusOutput = "Connected..."
 ' First command (mail from)
 .SendData (Header(1))
 StatusOutput = "Sending..."
 WaitFor ("250")
 ' Second (rcpt to)
 .SendData (Header(2))
 WaitFor ("250")
 .SendData ("data" & vbCrlf)
 WaitFor ("354")
 ' The rest
 .SendData Headers & vbCrlf
 ' This line is often found on MIME messages.
 .SendData "This is a multi-part message in MIME format." & vbCrlf
 .SendData vbCrlf
 For i = 11 To 20
  .SendData (Header(i) & vbCrlf)
 Next i
 .SendData (Header(21) & vbCrlf)
 ' Terminate
 .SendData ("." & vbCrlf)
 WaitFor ("250")
 ' Quit
 .SendData ("quit" & vbCrlf)
 StatusOutput = "Unconnected..."
 WaitFor ("221")
 .Close
 StatusOutput = ""
 Else
 Select Case .State
  Case 1
  StatusOutput = "Socket Opened."
  Case 2
  StatusOutput = "Listening..."
  Case 3
  StatusOutput = "Connection pending"
  Case 4
  StatusOutput = "Resolving host"
  Case 5
  StatusOutput = "Host resolved"
  Case 6
  StatusOutput = "Connecting"
  Case 7
  StatusOutput = "Connected"
  Case 8
  StatusOutput = "The point is closing the connection."
  Case 9
  StatusOutput = "Error."
  Case Else
  StatusOutput = "Undefined."
 End Select
 ' Just a box in case anything happens.
 MsgBox (StatusOutput)
 End If
End With
End Sub
Sub WaitFor(ResponseCode As String)
 Start = Timer
 While Len(Response) = 0
 Tmr = Start - Timer
 DoEvents
 If Tmr > 50 Then
  MsgBox "SMTP service error, timed out while waiting for response", 64, "Error!"
  Exit Sub
 End If
 Wend
 While Left(Response, 3) <> ResponseCode
 DoEvents
 If Tmr > 50 Then
  MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, "Error"
  Exit Sub
 End If
 Wend
 Response = ""
End Sub
'
' The following code goes wherever the Winsock
' control is placed.
'
Private Sub SCWinsock_DataArrival(ByVal bytesTotal As Long)
 SCWinsock.GetData Response
End Sub
```

