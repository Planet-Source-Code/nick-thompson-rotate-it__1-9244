VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotate-It - by Nick Thompson"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framedestoptions 
      Caption         =   "Destination Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1440
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox txtheight 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   44
         Text            =   "100"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtwidth 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   43
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optds 
         Caption         =   "Manual"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optds 
         Caption         =   "Source Diagonal"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optds 
         Caption         =   "Same as Source"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdrgbok 
         Caption         =   "OK"
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
         Left            =   6840
         TabIndex        =   36
         Top             =   840
         Width           =   495
      End
      Begin VB.HScrollBar BLUE 
         Height          =   255
         LargeChange     =   10
         Left            =   4680
         Max             =   255
         TabIndex        =   31
         Top             =   960
         Value           =   255
         Width           =   2055
      End
      Begin VB.HScrollBar GREEN 
         Height          =   255
         LargeChange     =   10
         Left            =   4680
         Max             =   255
         TabIndex        =   30
         Top             =   600
         Value           =   255
         Width           =   2055
      End
      Begin VB.HScrollBar RED 
         Height          =   255
         LargeChange     =   10
         Left            =   4680
         Max             =   255
         TabIndex        =   29
         Top             =   240
         Value           =   255
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   42
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Size of Destination"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblrgbresult 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6840
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblblue 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblgreen 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   33
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblred 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame frameloadsave 
      Caption         =   "Load and Save Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      TabIndex        =   7
      Top             =   0
      Width           =   7455
      Begin VB.OptionButton optdest 
         Caption         =   "Show Multi Destination"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton optdest 
         Caption         =   "Show Single Destination"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   50
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.TextBox txtspace 
         Height          =   285
         Left            =   4680
         TabIndex        =   47
         Text            =   "50"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkadd 
         Caption         =   "Add to same dest"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5400
         TabIndex        =   46
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdsavedir 
         Caption         =   "Save As"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdload 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chksave 
         Caption         =   "Save Pictures"
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkcls 
         Caption         =   "Clear Destination First"
         Height          =   255
         Left            =   5400
         TabIndex        =   21
         Top             =   480
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label lblspace 
         Caption         =   "Spacing"
         Height          =   255
         Left            =   3960
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "No save file name selected yet, click Save As"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   960
         Width           =   5655
      End
   End
   Begin VB.Frame framertype 
      Caption         =   "Rotation Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton optnr 
         Caption         =   "Many Rotations (custom number)"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optnr 
         Caption         =   "Single Rotation (custom angle)"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame framedirection 
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
      Begin VB.OptionButton optdir 
         Caption         =   "Anti-Clockwise"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optdir 
         Caption         =   "Clockwise"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame framecustom 
      Caption         =   "Custom Angle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
      Begin VB.TextBox txtcusa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Text            =   "20"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "0 - 360°"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4320
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files"
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   19
      Top             =   1320
      Width           =   4575
      Begin VB.CommandButton cmddestopt 
         Caption         =   "Destination Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdcls 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdstart 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame framedest 
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   4920
      TabIndex        =   0
      Top             =   2280
      Width           =   4575
      Begin VB.VScrollBar VScroll2 
         Height          =   3855
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   255
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   240
         Width           =   4095
      End
      Begin VB.PictureBox picsdest 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   480
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   289
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.PictureBox picdest 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   480
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   185
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.Frame framecustomr 
      Caption         =   "Custom Rotations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtcusr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "36"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblcusr 
         Caption         =   "10°"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Each"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame framesource 
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   4575
      Begin VB.VScrollBar VScroll1 
         Height          =   3855
         Left            =   120
         TabIndex        =   53
         Top             =   480
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   240
         Width           =   4095
      End
      Begin VB.PictureBox picsource 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1650
         Left            =   480
         ScaleHeight     =   110
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   150
         TabIndex        =   4
         Top             =   600
         Width           =   2250
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      X1              =   4800
      X2              =   4800
      Y1              =   1440
      Y2              =   6720
   End
   Begin VB.Label lblinfo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6840
      Width           =   9375
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim dontsave As Byte
  Dim saveas As String
  Dim destr As Integer
  Dim destg As Integer
  Dim destb As Integer
  Dim started As Byte
  Dim customra As Single
  Dim customra2 As Single
  Dim angle As Double
  Dim sangle As Double
  Dim cname As Integer
  Const Pi = 3.14159265359

  
  ' bmp_rotate(pic1, pic2, theta)
     ' Rotate the image in a picture box.
     '   pic1 is the picture box with the bitmap to rotate
     '   pic2 is the picture box to receive the rotated bitmap
     '   theta is the angle of rotation
     '
     Sub bmp_rotate(pic1 As Control, pic2 As Control, theta As Double)
       Dim c1x As Integer  ' Center of pic1.
       Dim c1y As Integer  '   "
       Dim c2x As Integer  ' Center of pic2.
       Dim c2y As Integer  '   "
       Dim a As Single     ' Angle of c2 to p2.
       Dim r As Integer    ' Radius from c2 to p2.

       Dim p1x As Integer  ' Position on pic1.
       Dim p1y As Integer  '   "
       Dim p2x As Integer  ' Position on pic2.
       Dim p2y As Integer  '   "
       Dim n As Integer    ' Max width or height of pic2.

       ' Compute the centers.
       c1x = pic1.ScaleWidth / 2
       c1y = pic1.ScaleHeight / 2
       c2x = pic2.ScaleWidth / 2
       c2y = pic2.ScaleHeight / 2

       ' Compute the image size.
       n = pic2.ScaleWidth
       If n < pic2.ScaleHeight Then n = pic2.ScaleHeight
       n = n / 2 - 1
       ' For each pixel position on pic2.
       For p2x = 0 To n
          For p2y = 0 To n
             ' Compute polar coordinate of p2.
             If p2x = 0 Then
               a = Pi / 2
             Else
               a = Atn(p2y / p2x)
             End If
             r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)

             ' Compute rotated position of p1.
             p1x = r * Cos(a + theta)
             p1y = r * Sin(a + theta)

             ' Copy pixels, 4 quadrants at once.
             c0& = pic1.Point(c1x + p1x, c1y + p1y)
             c1& = pic1.Point(c1x - p1x, c1y - p1y)

             c2& = pic1.Point(c1x + p1y, c1y - p1x)
             c3& = pic1.Point(c1x - p1y, c1y + p1x)
             If c0& <> -1 Then pic2.PSet (c2x + p2x, c2y + p2y), c0&
             If c1& <> -1 Then pic2.PSet (c2x - p2x, c2y - p2y), c1&
             If c2& <> -1 Then pic2.PSet (c2x + p2y, c2y - p2x), c2&
             If c3& <> -1 Then pic2.PSet (c2x - p2y, c2y + p2x), c3&
          Next
          ' Allow pending Windows messages to be processed.
          t% = DoEvents()
          If started = 0 Then Exit Sub
       Next
End Sub


Private Sub BLUE_Change()
destb = BLUE.Value
lblblue.Caption = destb
lblrgbresult.BackColor = RGB(destr, destg, destb)
End Sub

Private Sub BLUE_Scroll()
destb = BLUE.Value
lblblue.Caption = destb
lblrgbresult.BackColor = RGB(destr, destg, destb)

End Sub

Private Sub chkadd_Click()
  VScroll2.Value = 1
  HScroll2.Value = 1
  
If chkadd.Value = 1 Then
  picsdest.Height = picdest.Height
  picsdest.Width = (txtspace.Text * txtcusr.Text) * 15
  txtspace.Visible = True
  lblspace.Visible = True
  optdest(1).Visible = True
  optdest(1).Value = True
  picsdest.Visible = True
  picdest.Visible = False
Else
  txtspace.Visible = False
  lblspace.Visible = False
  optdest(1).Visible = False
  optdest(1).Value = False
  picsdest.Visible = False
  picdest.Visible = True
End If
End Sub

Private Sub cmdcls_Click()
If started = 1 Then Exit Sub
picdest.Cls
picsdest.Cls
End Sub

Private Sub cmddestopt_Click()
If frameloadsave.Visible = True Then
  cmddestopt.Caption = "Load/Save Options"
  framedestoptions.Visible = True
  frameloadsave.Visible = False
Else
  cmddestopt.Caption = "Destination Options"
  framedestoptions.Visible = False
  frameloadsave.Visible = True
End If
End Sub

Private Sub cmdexit_Click()
End

End Sub







Private Sub cmdload_Click()
CD1.ShowOpen
On Error GoTo handler
picsource.Picture = LoadPicture(CD1.filename)
If optds(0) = True Then
  picdest.Width = picsource.Width
  picdest.Height = picsource.Height
  txtwidth.Enabled = False
  txtheight.Enabled = False
  txtwidth.Text = picdest.Width / 15
  txtheight.Text = picdest.Height / 15
ElseIf optds(1) = True Then
  picdest.Width = Sqr((picsource.Width * picsource.Width) + (picsource.Height * picsource.Height))
  picdest.Height = picdest.Width
  txtwidth.Enabled = False
  txtheight.Enabled = False
  txtwidth.Text = picdest.Width / 15
  txtheight.Text = picdest.Height / 15
End If
VScroll1.Value = 0
VScroll2.Value = 0
HScroll1.Value = 0
HScroll2.Value = 0
VScroll1.Max = picsource.Height / 15
HScroll1.Max = picsource.Width / 15
VScroll2.Max = picdest.Height / 15
HScroll2.Max = picdest.Width / 15

Exit Sub
handler:
MsgBox "File not found or not compatible", vbOKOnly, "Load Error"
End Sub

Private Sub cmdrgbok_Click()
picdest.BackColor = RGB(destr, destg, destb)
picsdest.BackColor = RGB(destr, destg, destb)
End Sub

Private Sub cmdsavedir_Click()
CD1.ShowSave
saveas = CD1.filename
Label7.Caption = CD1.FileTitle & ".bmp"
End Sub

Private Sub cmdstart_Click()

  cname = 0
  If saveas = "" And chksave.Value = 1 Then
    Beep
    dontsave = 1
    MsgBox "This picture will not be saved because you have not selected a name for it", vbOKOnly, "Rotate-It"
  Else
    dontsave = 0
  End If
  If started = 1 Then
    started = 0
    cmdstart.Caption = "Start"
    optdir(0).Enabled = True
    optdir(1).Enabled = True
    cmdrgbok.Enabled = True
    Exit Sub
  End If
  started = 1
  cmdrgbok.Enabled = False
  cmdstart.Caption = "Stop"
  chksave.Enabled = False
  chkcls.Enabled = False
  optdir(0).Enabled = False
  optdir(1).Enabled = False
  If optnr(1).Value = True Then 'one rotation
    If chkcls.Value = 1 Then picdest.Cls
    angle = txtcusa.Text * (Pi / 180)
    If optdir(0) = True Then angle = angle - (2 * angle)
    Call bmp_rotate(picsource, picdest, angle)
    If chksave.Value = 1 And dontsave = 0 And chkadd.Value = 0 Then
      SavePicture picdest.Image, saveas & ".bmp"
    End If
  Else 'Many rotations
    
    chkadd.Enabled = False
    If chkadd.Value = 1 Then
      picsdest.Height = picdest.Height
      picsdest.Width = (txtspace.Text * txtcusr.Text) * 15
    End If
    cname = 1
    nrot = txtcusr.Text 'set number of rotations
    customra2 = customra 'stop user changing setting half way through
    angle = 0 'angle 0 for first rotation
    If chkcls.Value = 1 Then
      picdest.Cls 'clear if wanted
      picsdest.Cls
    End If
    Call bmp_rotate(picsource, picdest, angle) 'rotate it
    If chksave.Value = 1 And dontsave = 0 And chkadd.Value = 0 Then
      SavePicture picdest.Image, saveas & cname & ".bmp"
    End If
    If chkadd.Value = 1 Then
 '     success = BitBlt(picsdest.hDC, cname * txtspace.Text, 0, picdest.Width / 15, picdest.Height / 15, picdest.hDC, 0, 0, SRCCOPY)
      
      success = BitBlt(picsdest.hDC, 0, 0, picdest.Width / 15, picdest.Height / 15, picdest.hDC, 0, 0, SRCCOPY)
    End If
    nrot = nrot - 1 'count down
    If started = 0 Then Exit Sub
    For sangle = customra2 * (Pi / 180) To 2 * Pi Step customra2 * (Pi / 180)
      cname = cname + 1
      If nrot = 0 Then
        started = 0
        chkcls.Enabled = True
        chksave.Enabled = True
        cmdstart.Caption = "Start"
        started = 0
        optdir(0).Enabled = True
        optdir(1).Enabled = True
        cmdrgbok.Enabled = True
        chkadd.Enabled = True
        picsdest.Refresh
        If chksave.Value = 1 And dontsave = 0 And chkadd.Value = 1 Then
          SavePicture picsdest.Image, saveas & ".bmp"
        End If
        Exit Sub
      End If
      If chkcls.Value = 1 Then picdest.Cls
      If optdir(0) = True Then
        angle = sangle - (2 * sangle)
      Else
        angle = sangle
      End If
      Call bmp_rotate(picsource, picdest, angle)
      If chksave.Value = 1 And dontsave = 0 And chkadd.Value = 0 Then
        SavePicture picdest.Image, saveas & cname & ".bmp"
      End If
      If chkadd.Value = 1 Then
        success = BitBlt(picsdest.hDC, (cname - 1) * txtspace.Text, 0, picdest.Width / 15, picdest.Height / 15, picdest.hDC, 0, 0, SRCCOPY)
      
'      success = BitBlt(picsdest.hDC, 0, 0, picdest.Width / 15, picdest.Height / 15, picdest.hDC, 0, 0, SRCCOPY)
    End If
  
      nrot = nrot - 1
      If started = 0 Then Exit Sub
    Next
     chkadd.Enabled = True
 
  End If
  picsdest.Refresh
  If chksave.Value = 1 And dontsave = 0 And chkadd.Value = 1 Then
    SavePicture picsdest.Image, saveas & ".bmp"
  End If
  chkcls.Enabled = True
  chksave.Enabled = True
  cmdstart.Caption = "Start"
  started = 0
  optdir(0).Enabled = True
  optdir(1).Enabled = True
  Beep
  cmdrgbok.Enabled = True

End Sub






Private Sub Form_Load()
  saveas = ""
  frmmain.Left = (Screen.Width - frmmain.Width) / 2
  frmmain.Top = (Screen.Height - frmmain.Height) / 2
  started = 0
  framedestoptions.Top = frameloadsave.Top
  framedestoptions.Left = frameloadsave.Left
  destr = 255
  destg = 255
  destb = 255
    picdest.Width = Sqr((picsource.Width * picsource.Width) + (picsource.Height * picsource.Height))
  picdest.Height = picdest.Width
  txtwidth.Enabled = False
  txtheight.Enabled = False
  txtwidth.Text = picdest.Width / 15
  txtheight.Text = picdest.Height / 15
  VScroll1.Max = picsource.Height / 15
  HScroll1.Max = picsource.Width / 15
  VScroll2.Max = picdest.Height / 15
  HScroll2.Max = picdest.Width / 15
End Sub



Private Sub HScroll1_Change()
picsource.Left = 480 - (HScroll1.Value * 15)

End Sub

Private Sub HScroll1_Scroll()
picsource.Left = 480 - (HScroll1.Value * 15)

End Sub

Private Sub HScroll2_Change()
If picdest.Visible = True Then picdest.Left = 480 - (HScroll2.Value * 15)
If picsdest.Visible = True Then picsdest.Left = 480 - (HScroll2.Value * 15)
End Sub

Private Sub HScroll2_Scroll()
If picdest.Visible = True Then picdest.Left = 480 - (HScroll2.Value * 15)
If picsdest.Visible = True Then picsdest.Left = 480 - (HScroll2.Value * 15)

End Sub
Private Sub GREEN_Change()
destg = GREEN.Value
lblgreen.Caption = destg
lblrgbresult.BackColor = RGB(destr, destg, destb)
End Sub

Private Sub GREEN_Scroll()
destg = GREEN.Value
lblgreen.Caption = destg
lblrgbresult.BackColor = RGB(destr, destg, destb)

End Sub

Private Sub optdest_Click(Index As Integer)
If optdest(0).Value = True Then
  picsdest.Visible = False
  picdest.Visible = True
  VScroll2.Value = 0
  HScroll2.Value = 0
  VScroll2.Max = picdest.Height / 15
  HScroll2.Max = picdest.Width / 15
ElseIf optdest(1) = True Then
  picsdest.Height = picdest.Height
  picsdest.Width = (txtspace.Text * txtcusr.Text) * 15
  picsdest.Visible = True
  picdest.Visible = False
  VScroll2.Max = picsdest.Height / 15
  HScroll2.Max = picsdest.Width / 15
End If
End Sub



Private Sub optds_Click(Index As Integer)
If optds(0) = True Then
  picdest.Width = picsource.Width
  picdest.Height = picsource.Height
  txtwidth.Enabled = False
  txtheight.Enabled = False
  txtwidth.Text = picdest.Width / 15
  txtheight.Text = picdest.Height / 15
  picsdest.Height = picdest.Height
  picsdest.Width = (txtspace.Text * txtcusr.Text) * 15
ElseIf optds(1) = True Then
  picdest.Width = Sqr((picsource.Width * picsource.Width) + (picsource.Height * picsource.Height))
  picdest.Height = picdest.Width
  txtwidth.Enabled = False
  txtheight.Enabled = False
  txtwidth.Text = picdest.Width / 15
  txtheight.Text = picdest.Height / 15
  picsdest.Height = picdest.Height
  picsdest.Width = (txtspace.Text * txtcusr.Text) * 15
ElseIf optds(2) = True Then
  txtwidth.Enabled = True
  txtheight.Enabled = True
  txtwidth.Text = picdest.Width / 15
  txtheight.Text = picdest.Height / 15
End If
VScroll2.Value = 0
HScroll2.Value = 0
VScroll2.Max = picdest.Height / 15
HScroll2.Max = picdest.Width / 15

End Sub

Private Sub optnr_Click(Index As Integer)

If optnr(1) = True Then
  If chkadd.Value = 1 Then
    VScroll2.Value = 0
    HScroll2.Value = 0
    VScroll2.Max = picdest.Height / 15
    HScroll2.Max = picdest.Width / 15
  End If
  txtspace.Visible = False
  lblspace.Visible = False
  optdest(1).Visible = False
  optdest(1).Value = False
  picsdest.Visible = False
  picdest.Visible = True
chkadd.Enabled = False
  framecustom.Visible = True
  framecustomr.Visible = False
  If txtcusa.Text = "" Then
    cmdstart.Enabled = False
    lblinfo.Caption = "Please enter a number greater than 0 and less than 360"
    Exit Sub
  End If
  If IsNumeric(txtcusa.Text) Then
    If txtcusa.Text <= 0 Or txtcusa >= 360 Then
      txtcusa.ForeColor = RGB(255, 0, 0)
      cmdstart.Enabled = False
      lblinfo.Caption = "Number must be positive, greater than 0 and less than 360"
      Exit Sub
    End If
    txtcusa.ForeColor = RGB(0, 0, 0)
    lblinfo.Caption = "No Errors"
    cmdstart.Enabled = True
  Else
    txtcusa.ForeColor = RGB(255, 0, 0)
    cmdstart.Enabled = False
    lblinfo.Caption = "Not a valid number"
  End If

Else
  If chkadd.Value = 1 Then
    picsdest.Height = picdest.Height
    picsdest.Width = (txtspace.Text * txtcusr.Text) * 15
    txtspace.Visible = True
    lblspace.Visible = True
    optdest(1).Visible = True
    optdest(1).Value = True
    picsdest.Visible = True
    picdest.Visible = False
  End If
  chkadd.Enabled = True
  framecustom.Visible = False
  framecustomr.Visible = True
  If txtcusr.Text = "" Or txtcusr.Text = "0" Then
    txtcusr.ForeColor = RGB(255, 0, 0)
    lblcusr.Caption = "N/A"
    cmdstart.Enabled = False
    Exit Sub
  End If
  If IsNumeric(txtcusr.Text) Then
    If txtcusr.Text < 1 Then
      lblcusr.Caption = "N/A"
      cmdstart.Enabled = False
      Exit Sub
    End If
    If txtcusr.Text = Int(txtcusr.Text) Then
      lblinfo.Caption = "No Errors"
      customra = 360 / txtcusr.Text
      lblcusr.Caption = customra & "°"
      txtcusr.ForeColor = RGB(0, 0, 0)
     cmdstart.Enabled = True
   Else
      txtcusr.ForeColor = RGB(255, 0, 0)
      lblinfo.Caption = "Number of rotations must be an integer (1 - 999)"
      lblcusr.Caption = "N/A"
      cmdstart.Enabled = False
      Exit Sub
    End If
  Else
    lblcusr.Caption = "N/A"
    txtcusr.ForeColor = RGB(255, 0, 0)
    lblinfo.Caption = "Not a valid number"
    cmdstart.Enabled = False
  End If
  End If
End Sub

Private Sub RED_Change()
destr = RED.Value
lblred.Caption = destr
lblrgbresult.BackColor = RGB(destr, destg, destb)
End Sub

Private Sub RED_Scroll()
destr = RED.Value
lblred.Caption = destr
lblrgbresult.BackColor = RGB(destr, destg, destb)

End Sub

Private Sub txtcusa_Change()
  If txtcusa.Text = "" Then
    cmdstart.Enabled = False
    lblinfo.Caption = "Please enter a number greater than 0 and less than 360"
    Exit Sub
  End If
  If IsNumeric(txtcusa.Text) Then
    If txtcusa.Text <= 0 Or txtcusa >= 360 Then
      txtcusa.ForeColor = RGB(255, 0, 0)
      cmdstart.Enabled = False
      lblinfo.Caption = "Number must be positive, greater than 0 and less than 360"
      Exit Sub
    End If
    txtcusa.ForeColor = RGB(0, 0, 0)
    lblinfo.Caption = "No Errors"
    cmdstart.Enabled = True
  Else
    txtcusa.ForeColor = RGB(255, 0, 0)
    cmdstart.Enabled = False
    lblinfo.Caption = "Not a valid number"
  End If


End Sub

Private Sub txtcusr_Change()
  If txtcusr.Text = "" Or txtcusr.Text = "0" Then
    txtcusr.ForeColor = RGB(255, 0, 0)
    lblcusr.Caption = "N/A"
    cmdstart.Enabled = False
    Exit Sub
  End If
  If IsNumeric(txtcusr.Text) Then
    If txtcusr.Text < 1 Then
      lblcusr.Caption = "N/A"
      cmdstart.Enabled = False
      Exit Sub
    End If
    If txtcusr.Text = Int(txtcusr.Text) Then
      lblinfo.Caption = "No Errors"
      customra = 360 / txtcusr.Text
      lblcusr.Caption = customra & "°"
      txtcusr.ForeColor = RGB(0, 0, 0)
     cmdstart.Enabled = True
   Else
      txtcusr.ForeColor = RGB(255, 0, 0)
      lblinfo.Caption = "Number of rotations must be an integer (1 - 999)"
      lblcusr.Caption = "N/A"
      cmdstart.Enabled = False
      Exit Sub
    End If
  Else
    lblcusr.Caption = "N/A"
    txtcusr.ForeColor = RGB(255, 0, 0)
    lblinfo.Caption = "Not a valid number"
    cmdstart.Enabled = False
  End If
End Sub

Private Sub txtheight_Change()

If txtheight.Text = "" Then Exit Sub
If IsNumeric(txtheight.Text) Then
  If txtheight > 30000 Then
    lblinfo.Caption = "Number too large"
    txtheight.ForeColor = RGB(255, 0, 0)
    Exit Sub
  End If
  If txtheight.Text = Int(txtheight.Text) Then
    picdest.Height = txtheight.Text * 15
    txtheight.ForeColor = RGB(0, 0, 0)
  End If
End If
VScroll2.Value = 0
HScroll2.Value = 0
VScroll2.Max = picdest.Height / 15
HScroll2.Max = picdest.Width / 15
 picsdest.Height = picdest.Height
  picsdest.Width = (txtspace.Text * txtcusr.Text) * 15

End Sub

Private Sub txtwidth_Change()

If txtwidth.Text = "" Then Exit Sub
If IsNumeric(txtwidth.Text) Then
  If txtwidth > 30000 Then
    lblinfo.Caption = "Number too large"
    txtwidth.ForeColor = RGB(255, 0, 0)
    Exit Sub
  End If
  If txtwidth.Text = Int(txtwidth.Text) Then
    picdest.Width = txtwidth.Text * 15
    txtwidth.ForeColor = RGB(0, 0, 0)
  End If
End If
VScroll2.Value = 0
HScroll2.Value = 0
VScroll2.Max = picdest.Height / 15
HScroll2.Max = picdest.Width / 15
picsdest.Height = picdest.Height
picsdest.Width = (txtspace.Text * txtcusr.Text) * 15

End Sub

Private Sub VScroll1_Change()
picsource.Top = 600 - (VScroll1.Value * 15)

End Sub

Private Sub VScroll1_Scroll()
picsource.Top = 600 - (VScroll1.Value * 15)

End Sub

Private Sub VScroll2_Change()
picdest.Top = 600 - (VScroll2.Value * 15)
picsdest.Top = 600 - (VScroll2.Value * 15)
End Sub

Private Sub VScroll2_Scroll()
picdest.Top = 600 - (VScroll2.Value * 15)
picsdest.Top = 600 - (VScroll2.Value * 15)

End Sub
