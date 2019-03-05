VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VMI File Editor"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5520
      TabIndex        =   31
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   30
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   29
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Date"
      Height          =   2775
      Left            =   3360
      TabIndex        =   14
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox CboSec 
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Text            =   "Combo7"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox CboMin 
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         Text            =   "Combo6"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox CboHour 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Text            =   "Combo5"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox CboWeek 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Text            =   "Combo4"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox CboDay 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Text            =   "Combo3"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox CboMonth 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Text            =   "Combo2"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox CboYear 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0061
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Second"
         Height          =   255
         Left            =   1680
         TabIndex        =   27
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Minute"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Day of Week"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Copy Protected"
      Height          =   855
      Left            =   1680
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
      Begin VB.OptionButton OptCopy 
         Caption         =   "Yes"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptCopy 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
      Begin VB.OptionButton OptType 
         Caption         =   "Data"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Game"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      MaxLength       =   12
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MaxLength       =   32
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   32
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "VMU Name (12 characters)"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   ".VMS (8 characters)"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (32 characters)"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discrition (32 characters)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Load()
  Dim StrBuf1 As String * 32
  Dim StrBuf2 As String * 8
  Dim StrBuf3 As String * 12
  Dim Buffer As Long
  Dim Buffer2 As Integer
  Dim Bite As Byte
  Dim CheckSum As Long

  Open App.Path & "\sonic.vmi" For Binary As 1
    'checksum
    Get 1, , Buffer
    
    'text stuff
    Get 1, , StrBuf1
    Text1.Text = StrBuf1
    Get 1, , StrBuf1
    Text2.Text = StrBuf1
    
    'creation date/time
    Get 1, , Buffer2
    CboYear.Text = Buffer2
    Get 1, , Bite
    CboMonth.Text = Bite
    Get 1, , Bite
    CboDay.Text = Bite
    
    Get 1, , Bite
    CboHour.Text = Bite
    Get 1, , Bite
    CboMin.Text = Bite
    Get 1, , Bite
    CboSec.Text = Bite
    Get 1, , Bite
    CboWeek.Text = Bite
  
    'extra stuff
    Get 1, , Buffer2
    Get 1, , Buffer2
    
    'get naming information
    Get 1, , StrBuf2
    Text3.Text = StrBuf2
    Get 1, , StrBuf3
    Text4.Text = StrBuf3
    
    'get flags
    Get 1, , Buffer2
    If Buffer2 And &H1001 And &H2001 Then OptCopy(0).Value = 1 Else OptCopy(1).Value = 1
    If Buffer2 And &H1002 And &H2002 Then OptType(0).Value = 1 Else OptType(1).Value = 1
    
    'some other stuff
    Get 1, , Buffer2
    
    'filesize
    Get 1, , Buffer
 
  Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim StrBuf1 As String * 32
  Dim StrBuf2 As String * 8
  Dim StrBuf3 As String * 12
  Dim Buffer As Long
  Dim Buffer2 As Integer
  Dim Bite As Byte
  
  Open App.Path & "\Edited.VMI" For Binary As 1
    Buffer = 0
    Put 1, , Buffer

    StrBuf1 = Text1.Text
    Put 1, , StrBuf1
    StrBuf1 = Text2.Text
    Put 1, , StrBuf1
    
    Buffer2 = Val(CboYear.Text)
    Put 1, , Buffer2
    Bite = Val(CboMonth.Text)
    Put 1, , Bite
    Bite = Val(CboDay.Text)
    Put 1, , Bite

    Bite = Val(CboHour.Text)
    Put 1, , Bite
    Bite = Val(CboMin.Text)
    Put 1, , Bite
    Bite = Val(CboSec.Text)
    Put 1, , Bite
    Bite = Val(CboWeek.Text)
    Put 1, , Bite
    
    Buffer2 = 0
    Put 1, , Buffer2
    Buffer2 = 1
    Put 1, , Buffer2
    
    StrBuf2 = Text3.Text
    Put 1, , StrBuf2
    StrBuf3 = Text4.Text
    Put 1, , StrBuf3

    Buffer2 = 0
    If OptCopy(0).Value = True Then Buffer2 = Buffer2 Or &H1
    If OptType(0).Value = True Then Buffer2 = Buffer2 Or &H2
    Put 1, , Buffer2

    Buffer2 = 9
    Put 1, , Buffer2

    Buffer = 1
    Put 1, , Buffer

  Close
End Sub
