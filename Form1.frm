VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8B8B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calligraphy Draw by Ken Foster"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShadow 
      BackColor       =   &H00FFC4C4&
      Caption         =   "Shadow"
      Height          =   285
      Left            =   1230
      TabIndex        =   39
      Top             =   3240
      Value           =   1  'Checked
      Width           =   885
   End
   Begin Project1.CandyButton cmdUndo 
      Height          =   405
      Left            =   1200
      TabIndex        =   35
      Top             =   2130
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Undo Last 0"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   12632319
      ColorButtonUp   =   255
      ColorButtonDown =   8421631
      BorderBrightness=   0
      ColorBright     =   12632319
      ColorScheme     =   0
   End
   Begin Project1.CandyButton cmdClear 
      Height          =   405
      Left            =   1200
      TabIndex        =   34
      Top             =   2730
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   12648384
      ColorButtonUp   =   65280
      ColorButtonDown =   8454016
      BorderBrightness=   0
      ColorBright     =   12648384
      ColorScheme     =   0
   End
   Begin Project1.ucPanel ucPanel4 
      Height          =   810
      Left            =   30
      TabIndex        =   29
      Top             =   6075
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1429
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Picture Size in Pixels"
      ColorBottom     =   16744576
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   255
         Left            =   1410
         TabIndex        =   33
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   165
         Left            =   270
         TabIndex        =   32
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1140
         TabIndex        =   31
         Top             =   525
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   30
         TabIndex        =   30
         Top             =   525
         Width           =   975
      End
   End
   Begin Project1.ucPanel ucPanel3 
      Height          =   2040
      Left            =   30
      TabIndex        =   26
      Top             =   4005
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   3598
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Save Select"
      ColorBottom     =   16744576
      Begin Project1.CandyButton cmdSaveJpeg 
         Height          =   465
         Left            =   225
         TabIndex        =   37
         Top             =   1500
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Save Jpeg"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         ColorScheme     =   0
      End
      Begin Project1.CandyButton cmdSaveBitmap 
         Height          =   465
         Left            =   210
         TabIndex        =   36
         Top             =   960
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Save Bitmap"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         ColorScheme     =   0
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   45
         TabIndex        =   27
         Top             =   585
         Width           =   2010
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Filename"
         Height          =   225
         Left            =   60
         TabIndex        =   28
         Top             =   360
         Width           =   675
      End
   End
   Begin Project1.ucPanel ucPanel2 
      Height          =   1875
      Left            =   30
      TabIndex        =   19
      Top             =   2085
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   3307
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Color Select"
      ColorBottom     =   16744576
      Begin VB.Label LabelColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   555
         Width           =   750
      End
      Begin VB.Label LabelColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   195
         TabIndex        =   24
         Top             =   1035
         Width           =   750
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Shape Fill"
         Height          =   240
         Left            =   225
         TabIndex        =   23
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Outline"
         Height          =   195
         Left            =   285
         TabIndex        =   22
         Top             =   840
         Width           =   675
      End
      Begin VB.Label LabelColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   195
         TabIndex        =   21
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Background"
         Height          =   240
         Left            =   150
         TabIndex        =   20
         Top             =   1350
         Width           =   1365
      End
   End
   Begin Project1.ucPanel ucPanel1 
      Height          =   2025
      Left            =   30
      TabIndex        =   10
      Top             =   30
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   3572
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Brush Controls"
      ColorBottom     =   16744576
      Begin VB.ComboBox cboRotation 
         Height          =   315
         ItemData        =   "Form1.frx":030A
         Left            =   1335
         List            =   "Form1.frx":0326
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   705
      End
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1335
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "5"
         Top             =   315
         Width           =   480
      End
      Begin VB.ComboBox cboAngle 
         Height          =   315
         ItemData        =   "Form1.frx":034E
         Left            =   1320
         List            =   "Form1.frx":036A
         TabIndex        =   12
         Text            =   "cboAngle"
         Top             =   1110
         Width           =   735
      End
      Begin VB.ComboBox cboWidth 
         Height          =   315
         ItemData        =   "Form1.frx":039C
         Left            =   1320
         List            =   "Form1.frx":03BE
         TabIndex        =   11
         Text            =   "cboWidth"
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Border Direction:"
         Height          =   240
         Index           =   5
         Left            =   45
         TabIndex        =   18
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Border Size: "
         Height          =   240
         Index           =   4
         Left            =   345
         TabIndex        =   17
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         Height          =   240
         Left            =   735
         TabIndex        =   16
         Top             =   1155
         Width           =   480
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Brush Width:"
         Height          =   210
         Left            =   255
         TabIndex        =   15
         Top             =   1590
         Width           =   960
      End
   End
   Begin VB.ComboBox cboGrid 
      Height          =   315
      ItemData        =   "Form1.frx":03E6
      Left            =   885
      List            =   "Form1.frx":03F6
      TabIndex        =   9
      Text            =   "10"
      Top             =   6930
      Width           =   660
   End
   Begin VB.CheckBox chkLine 
      BackColor       =   &H00FFC4C4&
      Caption         =   "Show Grid"
      Height          =   375
      Left            =   1230
      TabIndex        =   7
      Top             =   3555
      Value           =   1  'Checked
      Width           =   930
   End
   Begin VB.PictureBox picundo 
      Height          =   1320
      Index           =   3
      Left            =   12090
      ScaleHeight     =   1260
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   4380
      Width           =   1095
   End
   Begin VB.PictureBox picundo 
      Height          =   1380
      Index           =   2
      Left            =   11955
      ScaleHeight     =   1320
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   2880
      Width           =   1125
   End
   Begin VB.PictureBox picundo 
      Height          =   1245
      Index           =   1
      Left            =   11925
      ScaleHeight     =   1185
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   1575
      Width           =   1215
   End
   Begin VB.PictureBox picundo 
      Height          =   1410
      Index           =   0
      Left            =   11910
      ScaleHeight     =   1350
      ScaleWidth      =   1125
      TabIndex        =   2
      Top             =   60
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   2250
      Top             =   6405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   5
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   2220
      MouseIcon       =   "Form1.frx":040A
      MousePointer    =   99  'Custom
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   579
      TabIndex        =   0
      Top             =   30
      Width           =   8715
   End
   Begin VB.PictureBox p2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   2040
      Left            =   2220
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   577
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   8715
   End
   Begin VB.PictureBox pGrid 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2040
      Left            =   2220
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   576
      TabIndex        =   8
      Top             =   15
      Visible         =   0   'False
      Width           =   8700
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   10710
      MousePointer    =   15  'Size All
      TabIndex        =   6
      Top             =   2130
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   7005
      Picture         =   "Form1.frx":0CD4
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   1740
      Left            =   3075
      Picture         =   "Form1.frx":1439
      Stretch         =   -1  'True
      Top             =   2820
      Width           =   4620
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Size"
      Height          =   210
      Left            =   150
      TabIndex        =   38
      Top             =   6990
      Width           =   705
   End
   Begin VB.Line Line1 
      X1              =   2175
      X2              =   2175
      Y1              =   2055
      Y2              =   7740
   End
   Begin VB.Image Image3 
      Height          =   5295
      Left            =   2175
      Picture         =   "Form1.frx":233E
      Top             =   2025
      Width           =   7755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Originial code by LaVolpe
'FYI: 3D Shapes Using Regions
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=58678&lngWId=1
'I've just modified and tweak it to work with a freehand drawing program
   
Option Explicit
   
   Const AC_SRC_OVER = &H0
   
   Private Type BLENDFUNCTION
   BlendOp As Byte
   BlendFlags As Byte
   SourceConstantAlpha As Byte
   AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" _
(ByVal hDC As Long, _
ByVal lInt As Long, _
ByVal lInt As Long, _
ByVal lInt As Long, _
ByVal lInt As Long, _
ByVal hDC As Long, _
ByVal lInt As Long, _
ByVal lInt As Long, _
ByVal lInt As Long, _
ByVal lInt As Long, _
ByVal BLENDFUNCT As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" _
(Destination As Any, _
Source As Any, _
ByVal Length As Long)

Dim BlendVal As Integer
Dim BF As BLENDFUNCTION, lBF As Long

' APIs used in DoThreeDedge()...
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (ByRef lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long

' APIs used in DoSample()...
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' APIs used in CreateSampleBrush()...
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Type RECT
Left As Long
top As Long
Right As Long
Bottom As Long
End Type

Dim undoCt As Integer   'undo counter

Private Sub chkShadow_Click()
   DoSample
End Sub

Private Sub Form_Load()
   ' initially set a 3D rotation value
   cboRotation.ListIndex = 7
   cboAngle.Text = -5
   cboWidth.Text = 6
   Label9.Caption = p1.ScaleWidth
   Label10.Caption = p1.ScaleHeight
End Sub

Private Sub Form_Resize()
   DrawGradient Me.hDC, Me.Width, Me.Height, vbWhite, &HFF8080, 1

   p2.Width = p1.Width
   p2.Height = p1.Height
   p2.top = p1.top
   p2.Left = p1.Left
   pGrid.Width = p1.Width
   pGrid.Height = p1.Height
   
   cboGrid_Click
   chkLine_Click
End Sub

Private Sub chkLine_Click()
   DoSample
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      Label8.Left = Label8.Left + X
      p1.Width = Label8.Left - 2000
      Label8.top = Label8.top + Y
      p1.Height = Label8.top - 50
   End If
   Label9.Caption = p1.ScaleWidth
   Label10.Caption = p1.ScaleHeight
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Form_Resize
   p1.Picture = p1.Image
End Sub

Private Sub p1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If undoCt < 0 Then undoCt = 0
   If undoCt > 3 Then   'shift undo pictureboxes to the left
      picundo(0).Picture = picundo(1).Picture
      picundo(1).Picture = picundo(2).Picture
      picundo(2).Picture = picundo(3).Picture
      undoCt = 3
   End If
   p2.Picture = p2.Image   'render picture
   picundo(undoCt).Picture = p2.Picture
   undoCt = undoCt + 1
   cmdUndo.Caption = "Undo Last " & undoCt
   p1.MousePointer = 99
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      p1.Line (X, Y)-(X + cboWidth.Text, Y + cboAngle.Text)
      p2.Line (X, Y)-(X + cboWidth.Text, Y + cboAngle.Text)
   End If
End Sub

Private Sub p1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   DoSample
End Sub

Private Sub DoSample()
   
   Dim newWidth As Long, newHeight As Long
   Dim winRgn As Long  ' combined mainRgn and extRgn
   Dim extRgn As Long ' clipping region used for painting; this is the 3D edge
   Dim testRgn As Long ' base shaped region from sample images
   Dim outlineBrush As Long, fillBrush1 As Long, fillBrush2 As Long
   
   ' ensure a valid 3D border size is passed
   If Val(txtSize) < 1 Then txtSize = "1"
   If Val(txtSize) > 255 Then txtSize = 255
   txtSize = Int(Val(txtSize.Text))
   
   ' create the base shaped region
   testRgn = CreateShapedRegion2(p2.Image.handle)
   ' create the new window region & return the winRgn & extRgn pointers
   winRgn = DoThreeDedge(testRgn, extRgn, Val(txtSize), Val(cboRotation.Text), newWidth, newHeight)
   DeleteObject testRgn    ' the original shaped region is no longer needed
   
   ' blank out so ready for painting
   p1.Picture = LoadPicture()
   
   If chkShadow.Value = Checked Then
   ' select the 3D edge region as the clipping region so we can fill it
   SelectClipRgn p1.hDC, extRgn
   ' solid 3D fill color
   fillBrush1 = CreateSolidBrush(LabelColor(1).BackColor)
   ' fill the 3D edge portion
   FillRgn p1.hDC, extRgn, fillBrush1
   ' remove the clipping region so we can draw our borders
   SelectClipRgn p1.hDC, ByVal 0&
   End If
   
   ' create brush & draw border around the entire region
   outlineBrush = CreateSolidBrush(LabelColor(1).BackColor)
  ' FrameRgn p1.hDC, winRgn, outlineBrush, 1, 1
   
   ' now we'll remove the 3D edge portion & draw another border (inner border)
   CombineRgn extRgn, winRgn, extRgn, 4
   
   SelectClipRgn p1.hDC, extRgn
   fillBrush2 = CreateSolidBrush(LabelColor(0).BackColor)
   FillRgn p1.hDC, extRgn, fillBrush2
   'remove the clipping region so we can draw other stuff if desired
   SelectClipRgn p1.hDC, ByVal 0&
   
   ' here we will frame the inner border
   FrameRgn p1.hDC, extRgn, outlineBrush, 1, 1
   ' clean up.  The only region not deleted is the winRgn 'cause we will assign
   ' it to our test window via SetWindowRgn
   DeleteObject extRgn
   DeleteObject outlineBrush
   DeleteObject fillBrush1
   DeleteObject fillBrush2
   If chkLine.Value = Checked Then
      p1.Refresh
      pGrid.Refresh
      BlendVal = 30
      DoAlphablend pGrid, p1, 30
   Else
      p1.Refresh
      pGrid.Refresh
      BlendVal = 0
      DoAlphablend pGrid, p1, 0
   End If
   p1.Refresh
End Sub

Private Function DoThreeDedge(ByVal inRegion As Long, ThreeDRgn As Long, _
   ByVal ThreeDsize As Byte, ByVal ThreeDangle As Integer, _
   RegionCX As Long, RegionCY As Long) As Long
   
   ' regions to be created & modified
   Dim pRgn As Long    ' this will be the return value
   Dim cRgn As Long    ' temp region used for creating the 3D effect
   
   ' rectangles used to align the new regions with the old
   Dim pRect As RECT
   Dim wRect As RECT
   
   ' used for offsets when aligning regions
   Dim cxAdjust As Long, cyAdjust As Long
   Dim i As Long   ' loop variable
   ' direction of the 3D effect
   Dim Xdir As Long, Ydir As Long
   
   ' Create temporary regions...
   
   ' make a copy of the passed region
   pRgn = CreateRectRgnIndirect(pRect)
   CombineRgn pRgn, inRegion, pRgn, 5
   If ThreeDsize = 0 Then Exit Function
   
   ' make another copy to be ultimately used for the clipping region
   ThreeDRgn = CreateRectRgnIndirect(pRect)
   
   ' make a 3rd copy to be used for creating the 3D effect
   cRgn = CreateRectRgnIndirect(pRect)
   CombineRgn cRgn, inRegion, cRgn, 5
   
   ' set the direction of X,Y depending on the passed rotation
   Select Case ThreeDangle
      Case 0, 360
         Xdir = 1: Ydir = 0
      Case 45
         Xdir = 1: Ydir = -1
      Case 90
         Xdir = 0: Ydir = -1
      Case 135
         Xdir = -1: Ydir = -1
      Case 180
         Xdir = -1: Ydir = 0
      Case 225
         Xdir = -1: Ydir = 1
      Case 270
         Xdir = 0: Ydir = 1
      Case Else
         Xdir = 1: Ydir = 1
   End Select
   
   ' we create the 3D effect by simply sliding the original region
   ' a pixel at a time in the appropriate direction enough times to
   ' satisify the size requested. Simple, huh?
   For i = 1 To ThreeDsize
      OffsetRgn cRgn, Xdir, Ydir                  ' slide region
      CombineRgn pRgn, pRgn, cRgn, 2              ' add to overall region
      Next
      DeleteObject cRgn           ' no longer needed, delete now
      
      ' get the X,Y position of the combined region
      GetRgnBox pRgn, pRect
      ' get the original position of the passed region
      GetRgnBox inRegion, wRect
      
      ' create the actual 3D region
      CombineRgn ThreeDRgn, pRgn, inRegion, 4
      
      ' when the 3D border is above or left to the original region, we need to
      ' do some offsetting to align everything up
      
      ' set the offsets. Not all passed regions will have a 0,0 top/left corner
      If Xdir < 0 Then cxAdjust = pRect.Left Else cxAdjust = wRect.Left
      If Ydir < 0 Then cyAdjust = pRect.top Else cyAdjust = wRect.top
      
      ' perform the offsets
      OffsetRgn pRgn, -cxAdjust + wRect.Left, -cyAdjust + wRect.top
      OffsetRgn ThreeDRgn, -cxAdjust + wRect.Left, -cyAdjust + wRect.top
      
      ' return the calculated region width & height
      RegionCX = wRect.Right + ThreeDsize * Abs(Xdir) + Abs(wRect.Left)
      RegionCY = wRect.Bottom + ThreeDsize * Abs(Ydir) + Abs(wRect.top)
      
      ' return the new comined region
      DoThreeDedge = pRgn
      
   End Function

Private Function CreateSampleBrush(Style As Long) As Long
   ' Function returns a handle to a bitmap brush.
   ' Note: Win95, I believe, will only use the first 8x8 of the bitmap
   
   ' I used a 16x80 swatch with 6 stacked sample bitmaps
   
   Dim xOffset As Long
   Dim tDC As Long, hOldBmp As Long, hNewBmp As Long
   
   ' create temp DC & bitmap.
   ' If used outside of this form, we would use
   ' the return value of GetDC(GetDesktopWindow()) vs Me.hDC below
   tDC = CreateCompatibleDC(Me.hDC)
   hNewBmp = CreateCompatibleBitmap(Me.hDC, 16, 16)
   
   ' select fresh bitmap into our temp DC & Blt over the appropriate 16x16 bits
   hOldBmp = SelectObject(tDC, hNewBmp)
   
   ' remove the bitmap & replace original, then delete the DC
   SelectObject tDC, hOldBmp
   DeleteDC tDC
   
   ' create the brush & then delete the temp bitmap
   CreateSampleBrush = CreatePatternBrush(hNewBmp)
   DeleteObject hNewBmp
   
End Function

Private Sub cboRotation_Click()
   ' option to change degrees of 3D border
   DoSample
End Sub

Private Sub cboGrid_Click()
   Dim dx As Integer
   pGrid.Cls
   For dx = cboGrid.Text To pGrid.Height Step cboGrid.Text
      pGrid.Line (0, dx)-(pGrid.ScaleWidth, dx), vbBlack
   Next dx
   For dx = cboGrid.Text To pGrid.Width Step cboGrid.Text
      pGrid.Line (dx, 0)-(dx, pGrid.ScaleHeight)
   Next dx
   DoSample
End Sub

Private Sub cmdClear_Click()
   Dim cX As Integer
   p1.Picture = LoadPicture()
   p2.Picture = LoadPicture()
   For cX = 0 To 3
      picundo(cX).Picture = LoadPicture()
   Next cX
   undoCt = 0
   cmdUndo.Caption = "Undo Last 0"
   Form_Resize
End Sub

Private Sub cmdSaveBitmap_Click()
   Dim rsp As String
   If Text1.Text = "" Then
      MsgBox "Enter a Filename", , "No filename"
      Exit Sub
   End If
   p1.Picture = p1.Image  'render picture
   'check if file exists already
   If Dir(App.Path & "\" & Text1.Text & ".bmp") = "" Then
      SavePicture p1.Picture, App.Path & "\Images\" & Text1.Text & ".bmp"
      MsgBox "Picture saved at " & App.Path & "\Images\" & Text1.Text & ".bmp", , "Save a Bitmap"
   Else
      rsp = MsgBox("File exists. Do you want to overwrite?", vbYesNo)
      If rsp = vbNo Then GoTo here
      SavePicture p1.Picture, App.Path & "\Images\" & Text1.Text & ".bmp"
      MsgBox "Picture saved at " & App.Path & "\Images\" & Text1.Text & ".bmp", , "Picture Saved"
here:
   End If
   
   Text1.Text = ""
End Sub

Private Sub cmdSaveJPEG_Click()
   Dim rsp As String
   If Text1.Text = "" Then
      MsgBox "Enter a Filename", , "No filename"
      Exit Sub
   End If
   p1.Picture = p1.Image  'render picture
   'check if file exists already
   If Dir(App.Path & "\" & Text1.Text & ".jpg") = "" Then
      SaveJPEG App.Path & "\Images\" & Text1.Text & ".jpg", p1, Me, True, 90
      MsgBox "Picture saved at " & App.Path & "\Images\" & Text1.Text & ".jpg", , "Save as jpg"
   Else
      rsp = MsgBox("File exists. Do you want to overwrite?", vbYesNo)
      If rsp = vbNo Then GoTo here
      SaveJPEG App.Path & "\Images\" & Text1.Text & ".jpg", p1, Me, True, 90
      MsgBox "Picture saved at " & App.Path & "\Images\" & Text1.Text & ".jpg", , "Picture Saved"
here:
   End If
   Text1.Text = ""
End Sub

Private Sub cmdUndo_Click()
   p2.Cls
   If cmdUndo.Caption = "Undo Last 0" Then Exit Sub   'no undo's left
   undoCt = undoCt - 1   'countdown undo counter
   If undoCt < 0 Then undoCt = 0
   If undoCt > 3 Then undoCt = 3
   p2.Picture = picundo(undoCt).Picture
   cmdUndo.Caption = "Undo Last " & undoCt    'show undo count
   picundo(undoCt).Picture = LoadPicture()    'clear picturebox
   p1.Picture = p2.Picture
   chkLine_Click
End Sub

Private Function SaveJPEG(ByVal Filename As String, pic As PictureBox, PForm As Form, Optional ByVal Overwrite As Boolean = True, Optional ByVal Quality As Byte = 90) As Boolean
   Dim JPEGclass As cJpeg
   Dim m_Picture As IPictureDisp
   Dim m_DC As Long
   Dim m_Millimeter As Single
   m_Millimeter = PForm.ScaleX(100, vbPixels, vbMillimeters)
   Set m_Picture = pic
   m_DC = pic.hDC
   'this is not my code....from PSC
   'initialize class
   Set JPEGclass = New cJpeg
   'check there is image to save and the filename string is not empty
   If m_DC <> 0 And LenB(Filename) > 0 Then
      'check for valid quality
      If Quality < 1 Then Quality = 1
      If Quality > 100 Then Quality = 100
      'set quality
      JPEGclass.Quality = Quality
      'save in full color
      JPEGclass.SetSamplingFrequencies 1, 1, 1, 1, 1, 1
      'copy image from hDC
      If JPEGclass.SampleHDC(m_DC, CLng(m_Picture.Width / m_Millimeter), CLng(m_Picture.Height / m_Millimeter)) = 0 Then
         'if overwrite is set and file exists, delete the file
         If Overwrite And LenB(Dir$(Filename)) > 0 Then Kill Filename
         'save file and return True if success
         SaveJPEG = JPEGclass.SaveFile(Filename) = 0
      End If
   End If
   'clear memory
   Set JPEGclass = Nothing
End Function

Private Sub LabelColor_Click(Index As Integer)
   ' get custom solid color & update the test form if needed
   GetColor Index
   p1.BackColor = LabelColor(2).BackColor
   DoSample
End Sub

Private Sub GetColor(lblIndex As Integer)
   ' called when user needs to select color from color dialog
   With dlgCommon
      .Flags = cdlCCRGBInit
      .Color = LabelColor(lblIndex).BackColor
   End With
   On Error Resume Next
   dlgCommon.ShowColor
   If Err.Number = 0 Then LabelColor(lblIndex).BackColor = dlgCommon.Color
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
   ' when user hits enter in text box, update sample form
   If KeyAscii = vbKeyReturn Then
      DoSample
      KeyAscii = 0
   End If
End Sub

Private Sub DoAlphablend(SrcPicBox As PictureBox, DestPicBox As PictureBox, AlphaVal As Integer)
   With BF
      .BlendOp = AC_SRC_OVER
      .BlendFlags = 0
      .SourceConstantAlpha = AlphaVal
      .AlphaFormat = 0
   End With
   'copy the BLENDFUNCTION-structure to a Long
   RtlMoveMemory lBF, BF, 4
   
   'AlphaBlend the picture from Picture1 over the picture of Picture2
   AlphaBlend DestPicBox.hDC, 0, 0, DestPicBox.ScaleWidth, DestPicBox.ScaleHeight, SrcPicBox.hDC, 0, 0, SrcPicBox.ScaleWidth, SrcPicBox.ScaleHeight, lBF
End Sub

