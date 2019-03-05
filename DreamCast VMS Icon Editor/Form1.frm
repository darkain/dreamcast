VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VMU/VMS Icon Editor"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptPage 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   29
      Top             =   4440
      Width           =   375
   End
   Begin VB.OptionButton OptPage 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   28
      Top             =   4200
      Width           =   375
   End
   Begin VB.OptionButton OptPage 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   27
      Top             =   3960
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   4320
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   4
      Top             =   120
      Width           =   3840
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   255
      Left            =   5760
      TabIndex        =   23
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Invert"
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton CmdInvert 
      Caption         =   "Invert"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Repaint"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   120
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   120
      Width           =   3840
   End
   Begin VB.Label Btn1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Left"
      Height          =   255
      Left            =   6960
      TabIndex        =   26
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Btn2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right"
      Height          =   255
      Left            =   7560
      TabIndex        =   25
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X: 0    Y: 0"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F"
      Height          =   255
      Index           =   15
      Left            =   8280
      TabIndex        =   21
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E"
      Height          =   255
      Index           =   14
      Left            =   8280
      TabIndex        =   20
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D"
      Height          =   255
      Index           =   13
      Left            =   8280
      TabIndex        =   19
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      Height          =   255
      Index           =   12
      Left            =   8280
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      Height          =   255
      Index           =   11
      Left            =   8280
      TabIndex        =   17
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      Height          =   255
      Index           =   10
      Left            =   8280
      TabIndex        =   16
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Index           =   9
      Left            =   8280
      TabIndex        =   15
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   8280
      TabIndex        =   14
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   8280
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   8280
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   9
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Picture3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X: 0    Y: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BW_Tile(0 To 31, 0 To 31) As Boolean
Dim Cl_Tile(0 To 2, 0 To 31, 0 To 31) As Byte
Dim CL_Palet(0 To 15) As Long

Dim MouseLeft As Byte
Dim MouseRight As Byte
Dim MouseIsDown As Boolean

Dim SysIcon As Boolean
Dim CurPage As Byte

Dim NewCommand As String




Private Sub PaintIcons()
  Dim i1 As Integer
  Dim i2 As Integer
  
  'paint black-n-white icon
  If SysIcon Then
    For i1 = 0 To 31
      For i2 = 0 To 31
        If BW_Tile(i1, i2) Then
          Picture1.Line (i1 * 8, i2 * 8)-(i1 * 8 + 7, i2 * 8 + 7), RGB(0, 0, 0), BF
        Else
          Picture1.Line (i1 * 8, i2 * 8)-(i1 * 8 + 7, i2 * 8 + 7), RGB(255, 255, 255), BF
        End If
      Next i2
    Next i1
  Else
    Picture1.Line (0, 0)-(256, 256), &HFFFFFF, BF
  End If

  'paint colour icon
  For i1 = 0 To 31
    For i2 = 0 To 31
      Picture2.Line (i1 * 8, i2 * 8)-(i1 * 8 + 7, i2 * 8 + 7), CL_Palet(Cl_Tile(CurPage, i1, i2)), BF
    Next i2
  Next i1
  
  'update palet selector thingy
  For i1 = 0 To 15
    Picture3(i1).BackColor = CL_Palet(i1)
  Next i1
  
  Btn1.BackColor = CL_Palet(MouseLeft)
  Btn2.BackColor = CL_Palet(MouseRight)
End Sub

Private Function HiNym(a As Byte) As Byte
  HiNym = a And &H70
  HiNym = HiNym \ &H10
  If (a And &H80) Then HiNym = HiNym + 8
  
  If HiNym > 15 Then HiNym = 15
End Function

Private Function LoNym(a As Byte) As Byte
  LoNym = a And &H1F And &H2F
End Function

Private Function Nym2Bite(a As Byte, b As Byte) As Byte
  Dim c As Byte
  c = b And &H1F And &H2F
  c = c * &H10
  c = c Or (a And &H1F And &H2F)
  Nym2Bite = c
End Function

Private Function Cnv16To32Bit(colour As Integer) As Long
  Dim colour1 As Integer
  Dim colour2 As Integer
  Dim colour3 As Integer
  Dim colour4 As Integer
  
  colour1 = (colour And &H100F And &H200F) * &H10
  colour2 = (colour And &H10F0 And &H20F0)
  colour3 = (colour And &H1F00 And &H2F00) \ &H10
  colour4 = (colour And &H7000 And &H7000) \ &H100
  If (colour And &H8000) Then colour4 = colour4 + &H80
  Cnv16To32Bit = RGB(colour3, colour2, colour1)
End Function

Private Function Cnv32To16Bit(colour As Long) As Integer
  Dim colour1 As Long
  Dim colour2 As Long
  Dim colour3 As Long
  Dim colour4 As Long
  
  colour1 = (colour And &H100000F0 And &H200000F0) * &H10
  colour2 = (colour And &H1000F000 And &H2000F000) \ &H100
  colour3 = (colour And &H10F00000 And &H20F00000) \ &H100000
  colour4 = 0 '(colour And &H80000000 And &H80000000) ' \ &H10000
  Cnv32To16Bit = colour1 Or colour2 Or colour3 Or colour4
End Function


Private Sub CmdInvert_Click()
  Dim i As Integer
  
   For i = 0 To 15
    CL_Palet(i) = Not CL_Palet(i)
    CL_Palet(i) = CL_Palet(i) And &H10FFFFFF And &H20FFFFFF
    Picture3(i).BackColor = CL_Palet(i)
  Next i
  
  PaintIcons
End Sub

Private Sub Command1_Click()
  Dim i1 As Integer
  Dim i2 As Integer
  
  For i1 = 0 To 31
    For i2 = 0 To 31
      BW_Tile(i1, i2) = 0
    Next i2
  Next i1
  
  PaintIcons
End Sub

Private Sub Command2_Click()
  PaintIcons
End Sub

Private Sub Command3_Click()
  Dim i1 As Integer
  Dim i2 As Integer
  
  For i1 = 0 To 31
    For i2 = 0 To 31
      BW_Tile(i1, i2) = Not BW_Tile(i1, i2)
    Next i2
  Next i1
  
  PaintIcons
End Sub

Private Sub Command4_Click()
  Dim i1 As Integer
  Dim i2 As Integer
  
  For i1 = 0 To 31
    For i2 = 0 To 31
      Cl_Tile(0, i1, i2) = 0
    Next i2
  Next i1
  
  PaintIcons
End Sub

Private Sub Form_Load()
  Dim Buffer As Long
  Dim Buffer2 As Integer
  Dim BW_Offset As Long
  Dim Cl_Offset As Long
  
  Dim i As Integer
  Dim i1 As Integer
  Dim i2 As Integer
  
  Dim Bite As Byte
  
  MouseLeft = 0
  MouseRight = 1
  
  NewCommand = Command
  
  If Left$(NewCommand, 1) = Chr$(34) Then
    NewCommand = Mid$(NewCommand, 2)
  End If
  
  If (Right$(Command, 1)) = Chr$(34) Then
    NewCommand = Left$(NewCommand, Len(NewCommand) - 1)
  End If
  
  On Error GoTo errhan
  Open NewCommand For Binary As 1
    'Text offset
    Get 1, , Buffer
    Get 1, , Buffer
    Get 1, , Buffer
    Get 1, , Buffer
    
    'Icon Offsets
    Get 1, , BW_Offset
    Get 1, , Cl_Offset
    If BW_Offset = 32 And Cl_Offset = 160 Then
      SysIcon = True
    Else
      Cl_Offset = &H60
      BW_Offset = 1
    End If

    
    'Load black-n-white icon
    Get 1, BW_Offset, Bite
    For i1 = 0 To 31
      For i2 = 0 To 3
        Get 1, , Bite
        BW_Tile((i2 * 8) + 7, i1) = Bite And &H11 And &H21
        BW_Tile((i2 * 8) + 6, i1) = Bite And &H12 And &H22
        BW_Tile((i2 * 8) + 5, i1) = Bite And &H14 And &H24
        BW_Tile((i2 * 8) + 4, i1) = Bite And &H18 And &H28
        BW_Tile((i2 * 8) + 3, i1) = Bite And &H10
        BW_Tile((i2 * 8) + 2, i1) = Bite And &H20
        BW_Tile((i2 * 8) + 1, i1) = Bite And &H40
        BW_Tile((i2 * 8) + 0, i1) = Bite And &H80
      Next i2
    Next i1
  
  
  'Load colour icon
  Get 1, Cl_Offset, Bite
  
  'load palet
  For i = 0 To 15
    Get 1, , Buffer2
    CL_Palet(i) = Cnv16To32Bit(Buffer2)
    Picture3(i).BackColor = CL_Palet(i)
  Next i
  
  'load data
'  If SysIcon Then
'    For i1 = 0 To 31
'      For i2 = 0 To 15
'        Get 1, , Bite
'        Cl_Tile(0, i2 * 2 + 1, i1) = LoNym(Bite)
'        Cl_Tile(0, i2 * 2, i1) = HiNym(Bite)
'        'Cl_Tile
'      Next i2
'    Next i1
'  Else
    For i = 0 To 2
      For i1 = 0 To 31
        For i2 = 0 To 15
          Get 1, , Bite
          Cl_Tile(i, i2 * 2 + 1, i1) = LoNym(Bite)
          Cl_Tile(i, i2 * 2, i1) = HiNym(Bite)
        Next i2
      Next i1
    Next i
'  End If
  
  Close
  
  PaintIcons
  Exit Sub
  
errhan:
  MsgBox "Error opening file: " & Command
  End
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Buffer As Long
  Dim Buffer2 As Integer
  Dim BW_Offset As Long
  Dim Cl_Offset As Long
  
  Dim i As Integer
  Dim i1 As Integer
  Dim i2 As Integer
  
  Dim Bite As Byte
  
  Open App.Path & "\output.vms" For Binary As 1
    'Header Info
    Put 1, , "Custon VMS Icon "
    
    'Icon Locations
    Buffer = &H60 + (512 * 3) + 32
    Put 1, , Buffer     'black n white
    Buffer = &H60
    Put 1, , Buffer     'colour
    
    Bite = 0  'dreamcast filename info
    For i = 0 To 23
      Put 1, , Bite
    Next i
    
    Bite = 0  'app ID
    For i = 0 To 15
      Put 1, , Bite
    Next i
    
    'number of icons
    Buffer2 = 2
    Put 1, , Buffer2
    
    'icon animation speed
    Buffer2 = &HF00
    Put 1, , Buffer2
    
    'grafix eyecatch
    Buffer2 = 0
    Put 1, , Buffer2
    
    'CRC
    Buffer2 = 0
    Put 1, , Buffer2
    
    'size of non-headed file
    Buffer = 0
    Put 1, , Buffer
    
    'reserved space
    Bite = 0
    For i = 0 To 19
      Put 1, , Bite
    Next i
    
    'save colour palet info
    For i = 0 To 15
      Buffer2 = Cnv32To16Bit(CL_Palet(i))
      Put 1, , Buffer2
    Next i
    
    For i = 0 To 2
      For i1 = 0 To 31
        For i2 = 0 To 15
          Bite = Nym2Bite(Cl_Tile(i, i2 * 2 + 1, i1), Cl_Tile(i, i2 * 2, i1))
          Put 1, , Bite
        Next i2
      Next i1
    Next i
   
   
   'save b-n-w icon
    For i1 = 0 To 31
      For i2 = 0 To 3
'        Get 1, , Bite
'        BW_Tile((i2 * 8) + 7, i1) = Bite And &H11 And &H21
'        BW_Tile((i2 * 8) + 6, i1) = Bite And &H12 And &H22
'        BW_Tile((i2 * 8) + 5, i1) = Bite And &H14 And &H24
'        BW_Tile((i2 * 8) + 4, i1) = Bite And &H18 And &H28
'        BW_Tile((i2 * 8) + 3, i1) = Bite And &H10
'        BW_Tile((i2 * 8) + 2, i1) = Bite And &H20
'        BW_Tile((i2 * 8) + 1, i1) = Bite And &H40
'        BW_Tile((i2 * 8) + 0, i1) = Bite And &H80
        Bite = 0
        If BW_Tile((i2 * 8) + 0, i1) Then Bite = Bite Or &H80
        If BW_Tile((i2 * 8) + 1, i1) Then Bite = Bite Or &H40
        If BW_Tile((i2 * 8) + 2, i1) Then Bite = Bite Or &H20
        If BW_Tile((i2 * 8) + 3, i1) Then Bite = Bite Or &H10
        If BW_Tile((i2 * 8) + 4, i1) Then Bite = Bite Or &H8
        If BW_Tile((i2 * 8) + 5, i1) Then Bite = Bite Or &H4
        If BW_Tile((i2 * 8) + 6, i1) Then Bite = Bite Or &H2
        If BW_Tile((i2 * 8) + 7, i1) Then Bite = Bite Or &H1
        Put 1, , Bite
      Next i2
    Next i1
    
  Close
End Sub

Private Sub OptPage_Click(Index As Integer)
  CurPage = Index
  PaintIcons
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not SysIcon Then Exit Sub
  
  Dim xx As Integer
  Dim yy As Integer
  Dim X1 As Integer
  Dim Y1 As Integer
  
  MouseIsDown = True
  X1 = X \ 8
  Y1 = Y \ 8
  
  If X1 < 0 Then X1 = 0
  If X1 > 31 Then X1 = 31
  If Y1 < 0 Then Y1 = 0
  If Y1 > 31 Then Y1 = 31
  
  xx = X1 * 8
  yy = Y1 * 8
  
  If Button = 1 Then
    Picture1.Line (xx, yy)-(xx + 7, yy + 7), RGB(0, 0, 0), BF
    BW_Tile(X1, Y1) = True
  ElseIf Button = 2 Then
    Picture1.Line (xx, yy)-(xx + 7, yy + 7), RGB(255, 255, 255), BF
    BW_Tile(X1, Y1) = False
  End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not SysIcon Then Exit Sub
  
  If X < 0 Then X = 0
  If X > 255 Then X = 255
  If Y < 0 Then Y = 0
  If Y > 255 Then Y = 255
  Label1.Caption = "X: " & (X \ 8) & "    Y: " & (Y \ 8)
  
  If MouseIsDown Then
    Call Picture1_MouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseIsDown = False
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim xx As Integer
  Dim yy As Integer
  Dim X1 As Integer
  Dim Y1 As Integer
  
  MouseIsDown = True
  X1 = X \ 8
  Y1 = Y \ 8
  
  If X1 < 0 Then X1 = 0
  If X1 > 31 Then X1 = 31
  If Y1 < 0 Then Y1 = 0
  If Y1 > 31 Then Y1 = 31
  
  xx = X1 * 8
  yy = Y1 * 8
  
  If Button = 1 Then
    Picture2.Line (xx, yy)-(xx + 7, yy + 7), CL_Palet(MouseLeft), BF
    Cl_Tile(0, X1, Y1) = MouseLeft
  ElseIf Button = 2 Then
    Picture2.Line (xx, yy)-(xx + 7, yy + 7), CL_Palet(MouseRight), BF
    Cl_Tile(0, X1, Y1) = MouseRight
  End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If X < 0 Then X = 0
  If X > 255 Then X = 255
  If Y < 0 Then Y = 0
  If Y > 255 Then Y = 255
  Label2.Caption = "X: " & (X \ 8) & "    Y: " & (Y \ 8)
  
  If MouseIsDown Then
    Call Picture2_MouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseIsDown = False
End Sub

Private Sub Picture3_DblClick(Index As Integer)
  Dlg.Color = CL_Palet(Index)
  Dlg.Flags = cdlCCFullOpen Or cdlCCRGBInit
  Dlg.ShowColor
  
  CL_Palet(Index) = Dlg.Color
  Picture3(Index).BackColor = Dlg.Color
  PaintIcons
End Sub

Private Sub Picture3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MouseLeft = Index
    Btn1.BackColor = CL_Palet(Index)
  Else
    MouseRight = Index
    Btn2.BackColor = CL_Palet(Index)
  End If
End Sub
