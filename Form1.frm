VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atomic Clock"
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Update Clock"
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   3240
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2280
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "?"
      Top             =   120
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BUsed As Boolean
Dim myTime As TimeF


Private Sub Check1_Click()
    If Not BUsed Then
        BUsed = True
        Winsock1.Connect "time.gov", 13
    End If
End Sub

Private Sub Command1_Click()
    Time = Text1.Text
End Sub

Private Sub Form_Load()
    If Not BUsed Then
        BUsed = True
        Winsock1.Connect "time.gov", 13
    End If
End Sub

Private Sub Timer1_Timer()
    If Text1.Text <> "?" Then
        myTime.S = myTime.S + 1
        If myTime.S = 60 Then
            myTime.M = myTime.M + 1
            myTime.S = 0
        End If
        If myTime.M = 60 Then
            myTime.H = myTime.H + 1
            myTime.M = 0
        End If
        If myTime.H = 24 Then
            myTime.H = 0
        End If
        Text1.Text = ITT(myTime.H) & ":" & ITT(myTime.M) & ":" & ITT(myTime.S)
    End If
End Sub

Private Sub Timer2_Timer()
    Label1.Caption = Label1.Caption + 1
    If (Label1.Caption >= 1440) And (Not BUsed) Then
        BUsed = True
        Winsock1.Connect "time.gov", 13
        Label1.Caption = 0
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim TheTime As String
    Winsock1.GetData TheTime, vbDate
    For x = 1 To Len(TheTime)
        If Mid(TheTime, x, 1) = ":" Then
            myTime.H = CByte(Mid(TheTime, x - 2, 2))
            myTime.M = CByte(Mid(TheTime, x + 1, 2))
            myTime.S = CByte(Mid(TheTime, x + 4, 2))
            Exit For
        End If
    Next x
    Winsock1.Close
    BUsed = False
    myTime = GMTtoLT(myTime)
    Text1.Text = ITT(myTime.H) & ":" & ITT(myTime.M) & ":" & ITT(myTime.S)
End Sub

Private Function ITT(Value As Integer) As String
    If Value < 10 Then
        ITT = "0" & Value
    Else
        ITT = Value
    End If
End Function
