VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Microcrap Shoutlook"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox Text3 
         Height          =   2175
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1680
         TabIndex        =   2
         Top             =   690
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Height          =   615
         Left            =   2160
         TabIndex        =   1
         Top             =   4800
         Width           =   1815
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   1080
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   327681
      End
      Begin VB.Label Label3 
         Caption         =   "Body :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Subject :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "To Email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim Start As Single, Tmr As Single

Sub SendEmail(MailServerName As String, FromName As String, Filepath As String, Attachment As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
 Dim x As String
 Dim stri As String
 
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail per program start
    If Winsock1.State = sckClosed Then ' Check to see if socet is closed
        DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
        first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
        Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
        Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
        Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
        Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
        Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
        '******
        sixone = "MIME-Version: 1.0" + vbCrLf
        sixtwo = "Content-Type: application/octet-stream;name=""Attachment "" " + vbCrLf
        sixthree = "Content-Transfer-Encoding: base64" + vbCrLf + "Content-Disposition: attachment;" & " filename=" & Attachment & vbCrLf
        sixfour = vbCrLf
     
        'Open file in binary format ,Encode it  and then relay the data in Base 64 stream
        
        Open Trim(Filepath & Attachment) For Binary As #1
        i = 1
        stri = ""
        While Not EOF(1)
            x = String(1, " ")
            Get #1, , x
            stri = stri + x
            i = i + 1
        Wend
        Close #1
        sixfive = UUEncode(stri)
        
        'the Encoding routine is too slow for attachments in excess of 20 KB. but the idea is
        'to show it works. You can grab some fast C MIME encoding DLL from the net and
        'then encode it
        '******
        Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
        Ninth = "X-Mailer: SMTP" + vbCrLf ' What program sent the e-mail, customize this
        Eighth = Fourth + Third + Ninth + Fifth + Sixth + sixone + sixtwo + sixthree + sixfour + sixfive ' Combine for proper SMTP sending
        Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
        Winsock1.RemoteHost = MailServerName ' Set the server address
        Winsock1.RemotePort = 25 ' Set the SMTP Port
        Winsock1.Connect ' Start connection
        WaitFor ("220")
        Winsock1.SendData ("HELO worldcomputers.com" + vbCrLf)
        WaitFor ("250")
        Winsock1.SendData (first)
        WaitFor ("250")
        Winsock1.SendData (Second)
        WaitFor ("250")
        Winsock1.SendData ("data" + vbCrLf)
        WaitFor ("354")
        Winsock1.SendData (Eighth + vbCrLf)
        Winsock1.SendData (Seventh + vbCrLf)
        Winsock1.SendData ("." + vbCrLf)
        WaitFor ("250")
        Winsock1.SendData ("quit" + vbCrLf)
        WaitFor ("221")
        Winsock1.Close
    Else
        MsgBox (Str(Winsock1.State))
    End If
End Sub


Private Sub WaitFor(ResponseCode As String)
    Start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = Start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**

            If Tmr > 50 Then ' Time in seconds to wait
                MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
                Exit Sub
            End If
        Wend
       While Left(Response, 3) <> ResponseCode
           DoEvents
               If Tmr > 50 Then
                    MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
                    Exit Sub
                End If

            Wend
            Response = "" ' Sent response code to blank **IMPORTANT**
        End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData Response ' Check for incoming response *IMPORTANT*
End Sub
Private Function UUEncode(ByVal text) As String
    Dim a1 As Integer
    Dim a2 As Integer
    Dim a3 As Integer
    Dim LineChars As Integer
    Dim OutStream As String
    Dim CharTable As String
    Dim i As Long
    Dim c As Long
    CharTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    c = 1
    i = Int(4 * Len(text) / 3)
    OutStream = Space(i + 2 * (i \ 72 + 1))
    LineChars = 0
    For i = 1 To Len(text) - 2 Step 3
        a1 = Asc(Mid(text, i, 1))
        a2 = Asc(Mid(text, i + 1, 1))
        a3 = Asc(Mid(text, i + 2, 1))
        Mid(OutStream, c, 4) = Mid(CharTable, (a1 And &HFC) \ 4 + 1, 1) _
        & Mid(CharTable, (16 * (a1 And &H3) + (a2 And &HF0) \ 16) + 1, 1) _
        & Mid(CharTable, (4 * (a2 And &HF) + (a3 And &HC0) \ 64) + 1, 1) _
        & Mid(CharTable, (a3 And &H3F) + 1, 1)
        c = c + 4
        LineChars = LineChars + 4
        If LineChars >= 72 Then
            LineChars = 0
            Mid(OutStream, c, 2) = vbCrLf
            c = c + 2
        End If
    Next
    Select Case Len(text) Mod 3
    Case 1
    a1 = Asc(Mid(text, i, 1))
    a2 = 0
    a3 = 0
    Mid(OutStream, c, 4) = Mid(CharTable, (a1 And &HFC) \ 4 + 1, 1) _
    & Mid(CharTable, (16 * (a1 And &H3) + (a2 And &HF0) \ 16) + 1, 1) _
    & "==" & vbCrLf
    c = c + 6
    Case 2
    a1 = Asc(Mid(text, i, 1))
    a2 = Asc(Mid(text, i + 1, 1))
    a3 = 0
    Mid(OutStream, c, 4) = Mid(CharTable, (a1 And &HFC) \ 4 + 1, 1) _
    & Mid(CharTable, (16 * (a1 And &H3) + (a2 And &HF0) \ 16) + 1, 1) _
    & Mid(CharTable, (4 * (a2 And &HF) + (a3 And &HC0) \ 64) + 1, 1) _
    & "=" & vbCrLf
    c = c + 6
    End Select
    UUEncode = Left(OutStream, c - 1)
End Function


Private Sub Command1_Click()
'put YOUR parameters in all of the below..leave text1.text, text2.text and text3.text and C:\ alone
'make sure you have a local file called info.ini in your C:\ path. to attach
' by putting these values you will be able to send a TEST email to yourself.
'Open your Email Client and recieve the email in your INBOX once you send it

SendEmail "YOUR ISP ", "YOUR EMAIL ", "C:\", "info.ini", "YOUR EMAIL ", "TO EMAIL ", Text1.text, Text2.text, Text3.text
End Sub



