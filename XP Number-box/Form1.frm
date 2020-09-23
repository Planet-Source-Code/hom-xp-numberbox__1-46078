VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Number Box & Balloon"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2475
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   24000
      Top             =   13200
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   270
      Left            =   1200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   615
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   327681
      Value           =   2
      BuddyControl    =   "Text1"
      BuddyDispid     =   196611
      OrigLeft        =   1200
      OrigTop         =   615
      OrigRight       =   1455
      OrigBottom      =   885
      Max             =   999
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   265
      Left            =   1155
      ScaleHeight     =   270
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   615
      Width           =   135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   295
      Left            =   840
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "0"
      Top             =   600
      Width           =   360
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   295
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   4
      Top             =   600
      Width           =   270
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function OnlyNumberS1(TextName As TextBox)
        TextNameForTimer = TextName
        ShowBalloonTip TextName.hwnd, "Unacceptable character", "You can only type a number here", etiError
        Beep
        Timer1.Enabled = True
End Function

Private Sub text1_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 9) Then
        KeyAscii = 0
        OnlyNumberS1 Text1
    End If
End Sub
Private Sub Timer1_Timer()
    Form1.SetFocus
    Timer1.Enabled = False
End Sub
Private Sub UpDown1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
       Text1.SelLength = 0
End Sub
Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
       Text1.SelLength = Len(Text1.Text)
End Sub
