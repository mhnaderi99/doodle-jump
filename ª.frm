VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1500
      Left            =   5880
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   3
      Top             =   720
      Width           =   1050
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   7500
      Left            =   120
      ScaleHeight     =   500
      ScaleMode       =   0  'User
      ScaleWidth      =   346
      TabIndex        =   1
      Top             =   240
      Width           =   5250
      Begin VB.Image Image11 
         Height          =   555
         Left            =   1320
         Picture         =   "ª.frx":0000
         Top             =   4560
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Image Image10 
         Height          =   555
         Left            =   1320
         Picture         =   "ª.frx":048F
         Top             =   4560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Image9 
         Height          =   300
         Left            =   2400
         Picture         =   "ª.frx":0B3D
         Top             =   4680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   300
         Left            =   2400
         Picture         =   "ª.frx":1003
         Top             =   4680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   180
         Left            =   2400
         Picture         =   "ª.frx":14C9
         Top             =   3840
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   19
         Left            =   2880
         Picture         =   "ª.frx":253A
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   18
         Left            =   2880
         Picture         =   "ª.frx":2BD9
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   17
         Left            =   2880
         Picture         =   "ª.frx":3278
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   16
         Left            =   2880
         Picture         =   "ª.frx":3917
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   15
         Left            =   2880
         Picture         =   "ª.frx":3FB6
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   14
         Left            =   2880
         Picture         =   "ª.frx":4655
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   13
         Left            =   2880
         Picture         =   "ª.frx":4CF4
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   12
         Left            =   2880
         Picture         =   "ª.frx":5393
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   11
         Left            =   2880
         Picture         =   "ª.frx":5A32
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   10
         Left            =   2880
         Picture         =   "ª.frx":60D1
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   9
         Left            =   2880
         Picture         =   "ª.frx":6770
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   8
         Left            =   2880
         Picture         =   "ª.frx":6E0F
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   7
         Left            =   2880
         Picture         =   "ª.frx":74AE
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   6
         Left            =   2880
         Picture         =   "ª.frx":7B4D
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   5
         Left            =   2880
         Picture         =   "ª.frx":81EC
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   4
         Left            =   2880
         Picture         =   "ª.frx":888B
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   3
         Left            =   2880
         Picture         =   "ª.frx":8F2A
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   2
         Left            =   2880
         Picture         =   "ª.frx":95C9
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   1
         Left            =   2880
         Picture         =   "ª.frx":9C68
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   225
         Index           =   20
         Left            =   2880
         Picture         =   "ª.frx":A307
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image5 
         Height          =   225
         Index           =   4
         Left            =   1200
         Picture         =   "ª.frx":A9A6
         Top             =   3840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Image5 
         Height          =   225
         Index           =   3
         Left            =   1200
         Picture         =   "ª.frx":B0A3
         Top             =   3840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Image5 
         Height          =   225
         Index           =   2
         Left            =   1200
         Picture         =   "ª.frx":B7A0
         Top             =   3840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Image5 
         Height          =   225
         Index           =   1
         Left            =   1200
         Picture         =   "ª.frx":BE9D
         Top             =   3840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Image5 
         Height          =   225
         Index           =   5
         Left            =   1200
         Picture         =   "ª.frx":C59A
         Top             =   3840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Image4 
         Height          =   1320
         Index           =   47
         Left            =   3720
         Picture         =   "ª.frx":CC97
         Top             =   2280
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Image Image4 
         Height          =   1320
         Index           =   46
         Left            =   3720
         Picture         =   "ª.frx":E842
         Top             =   2280
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Image Image4 
         Height          =   735
         Index           =   45
         Left            =   3120
         Picture         =   "ª.frx":103ED
         Top             =   1560
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Image Image4 
         Height          =   735
         Index           =   44
         Left            =   3120
         Picture         =   "ª.frx":10C33
         Top             =   1560
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Image Image4 
         Height          =   1065
         Index           =   43
         Left            =   1560
         Picture         =   "ª.frx":11479
         Top             =   1440
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Image Image4 
         Height          =   1065
         Index           =   42
         Left            =   1560
         Picture         =   "ª.frx":12AD3
         Top             =   1440
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Image Image4 
         Height          =   750
         Index           =   41
         Left            =   4560
         Picture         =   "ª.frx":1412D
         Top             =   0
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image Image4 
         Height          =   750
         Index           =   40
         Left            =   4560
         Picture         =   "ª.frx":14AF0
         Top             =   0
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image Image4 
         Height          =   1380
         Index           =   39
         Left            =   4080
         Picture         =   "ª.frx":154B3
         Top             =   840
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Image Image4 
         Height          =   1380
         Index           =   38
         Left            =   4080
         Picture         =   "ª.frx":16CB3
         Top             =   840
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Image Image4 
         Height          =   1440
         Index           =   37
         Left            =   0
         Picture         =   "ª.frx":184B3
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Image Image4 
         Height          =   1440
         Index           =   36
         Left            =   0
         Picture         =   "ª.frx":1A81E
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Image Image4 
         Height          =   675
         Index           =   35
         Left            =   3000
         Picture         =   "ª.frx":1CB89
         Top             =   840
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Image Image4 
         Height          =   675
         Index           =   34
         Left            =   3000
         Picture         =   "ª.frx":1D786
         Top             =   840
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Image Image4 
         Height          =   795
         Index           =   33
         Left            =   1440
         Picture         =   "ª.frx":1E383
         Top             =   720
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Image Image4 
         Height          =   795
         Index           =   32
         Left            =   1440
         Picture         =   "ª.frx":1F4B9
         Top             =   720
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Image Image4 
         Height          =   825
         Index           =   31
         Left            =   0
         Picture         =   "ª.frx":205EF
         Top             =   360
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Image Image4 
         Height          =   825
         Index           =   30
         Left            =   0
         Picture         =   "ª.frx":21873
         Top             =   360
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   29
         Left            =   3720
         Picture         =   "ª.frx":22AF7
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   28
         Left            =   3720
         Picture         =   "ª.frx":23352
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image Image4 
         Height          =   1080
         Index           =   27
         Left            =   2040
         Picture         =   "ª.frx":23BAD
         Top             =   -120
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Image Image4 
         Height          =   1080
         Index           =   26
         Left            =   2040
         Picture         =   "ª.frx":251FF
         Top             =   -120
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   25
         Left            =   600
         Picture         =   "ª.frx":26851
         Top             =   -120
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   48
         Left            =   600
         Picture         =   "ª.frx":27405
         Top             =   -120
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Image Image4 
         Height          =   1320
         Index           =   23
         Left            =   3720
         Picture         =   "ª.frx":27FB9
         Top             =   2280
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Image Image4 
         Height          =   735
         Index           =   22
         Left            =   3120
         Picture         =   "ª.frx":29B64
         Top             =   1560
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Image Image4 
         Height          =   1065
         Index           =   21
         Left            =   1560
         Picture         =   "ª.frx":2A3AA
         Top             =   1440
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Image Image4 
         Height          =   750
         Index           =   20
         Left            =   4560
         Picture         =   "ª.frx":2BA04
         Top             =   0
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image Image4 
         Height          =   1380
         Index           =   19
         Left            =   4080
         Picture         =   "ª.frx":2C3C7
         Top             =   840
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Image Image4 
         Height          =   1440
         Index           =   18
         Left            =   0
         Picture         =   "ª.frx":2DBC7
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Image Image4 
         Height          =   675
         Index           =   17
         Left            =   3000
         Picture         =   "ª.frx":2FF32
         Top             =   840
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Image Image4 
         Height          =   795
         Index           =   16
         Left            =   1440
         Picture         =   "ª.frx":30B2F
         Top             =   720
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Image Image4 
         Height          =   825
         Index           =   15
         Left            =   0
         Picture         =   "ª.frx":31C65
         Top             =   360
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   14
         Left            =   3720
         Picture         =   "ª.frx":32EE9
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image Image4 
         Height          =   1080
         Index           =   13
         Left            =   2040
         Picture         =   "ª.frx":33744
         Top             =   -120
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   24
         Left            =   600
         Picture         =   "ª.frx":34D96
         Top             =   -120
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Image Image4 
         Height          =   795
         Index           =   12
         Left            =   1440
         Picture         =   "ª.frx":3594A
         Top             =   720
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Image Image4 
         Height          =   735
         Index           =   11
         Left            =   3120
         Picture         =   "ª.frx":36A80
         Top             =   1560
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   10
         Left            =   3720
         Picture         =   "ª.frx":372C6
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image Image4 
         Height          =   675
         Index           =   9
         Left            =   3000
         Picture         =   "ª.frx":37B21
         Top             =   840
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   8
         Left            =   600
         Picture         =   "ª.frx":3871E
         Top             =   -120
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Image Image4 
         Height          =   825
         Index           =   7
         Left            =   0
         Picture         =   "ª.frx":392D2
         Top             =   360
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Image Image4 
         Height          =   1440
         Index           =   6
         Left            =   0
         Picture         =   "ª.frx":3A556
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Image Image4 
         Height          =   750
         Index           =   5
         Left            =   4560
         Picture         =   "ª.frx":3C8C1
         Top             =   0
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image Image4 
         Height          =   1320
         Index           =   4
         Left            =   3720
         Picture         =   "ª.frx":3D284
         Top             =   2280
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Image Image4 
         Height          =   1380
         Index           =   3
         Left            =   4080
         Picture         =   "ª.frx":3EE2F
         Top             =   840
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Image Image4 
         Height          =   1080
         Index           =   2
         Left            =   2040
         Picture         =   "ª.frx":4062F
         Top             =   -120
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Image Image4 
         Height          =   1065
         Index           =   1
         Left            =   1560
         Picture         =   "ª.frx":41C81
         Top             =   1440
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image Image1 
         Height          =   595
         Left            =   2160
         Picture         =   "ª.frx":432DB
         Stretch         =   -1  'True
         Top             =   6360
         Width           =   600
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Top             =   7320
         Width           =   5895
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Next page  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1050
      Left            =   1680
      Picture         =   "ª.frx":45279
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   1050
      Left            =   2700
      Picture         =   "ª.frx":47217
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(500, 2, 1000)
Dim b(100, 2, 1000)
Dim c(100, 2, 1000)
Dim dd(20, 2, 1000): Dim ee(20)
Dim sh(10, 2, 1000): Dim ss(10)
Dim q(48, 2, 1000): Dim qq(48)
Dim i, yy, vy, ay, xx, vx, ax, h, page, brown, t, g, v, counter, he, d, r, t2, k, s, l, m, n, f, e2, time, p, s2, time2, p2, q2, vv, shatel
Private Sub Form_Load()
  Randomize Timer
  page = 1
  brown = RGB(120, 60, 60)
  h = 0.1
  vy = -50
  ay = 15
  vx = 0
  ax = 0
  xx = Image1.Left
  yy = Image1.Top
  t = 15
  g = 3
  f = 1
  u = 0
  v = 60
  r = Int(Picture1.ScaleHeight / 5)
  t2 = Int(t / 5)
  k = 0
  l = 0
  m = 20
  s2 = 10
  q2 = 48
  vv = 1
  counter = 0
  time = 1000
  shatel = 0
  time2 = 1000
  ay2 = ay
  vy2 = -(vy)
End Sub

Private Sub Label1_Click()
  counter = counter + 1
  If counter = 1 Then
    Picture1.SetFocus
    For i = 1 To t2
      For k = 0 To 4
        s = s + 1
        For j = 0 To 1000
          a(s, 1, j) = Int(Rnd * Picture1.ScaleWidth)
          a(s, 2, j) = Int(Rnd * Picture1.ScaleHeight / 5) + r * k
        Next j
      Next k
    Next i

    For i = 1 To g
      For j = 1 To 1000
        b(i, 1, j) = Int(Rnd * Picture1.ScaleWidth + 100) - 200
        b(i, 2, j) = Int(Rnd * Picture1.ScaleHeight)
        Image5(i).Visible = True
      Next j
    Next i

    For i = 1 To f
      For j = 1 To 1000
        e = Int(Rnd * t) + 1
        c(i, 1, j) = a(e, 1, j) + 25
        c(i, 2, j) = a(e, 2, j) - 5
        Image7.Visible = True
      Next j
    Next i

    For i = 1 To m
      e = Int(Rnd * t) + 1
      e2 = Int(Rnd * 1000) + 1
      ee(i) = e2
      dd(i, 1, ee(i)) = a(e, 1, ee(i)) + 25
      dd(i, 2, ee(i)) = a(e, 2, ee(i)) - 10
    Next i

    For i = 1 To s2
      e = Int(Rnd * t) + 1
      e2 = Int(Rnd * 1000) + 1
      ss(i) = e2
      sh(i, 1, e2) = a(e, 1, e2) + 25
      sh(i, 2, e2) = a(e, 2, e2) - 10
    Next i

    For i = 1 To q2
      e = Int(Rnd * 1000) + 1
      e2 = Int(Rnd * (Picture1.ScaleWidth - Image4(i).Width)) + 1
      e3 = Int(Rnd * (Picture1.ScaleWidth - Image4(i).Height)) + 1
      qq(i) = e
      q(i, 1, e) = e2
      q(i, 2, e) = e3
    Next i

    For i = 1 To t
      Image6(i).Left = a(i, 1, 1)
      Image6(i).Top = a(i, 2, page)
      Image6(i).Visible = True
    Next i

    For i = 1 To g
      Image5(i).Left = b(i, 1, 1)
      Image5(i).Top = b(i, 2, 1)
    Next i

    For i = 1 To f
      Image7.Left = c(i, 1, 1) - Image7.Width / 2
      Image7.Top = c(i, 2, 1)
    Next i

    For i = 1 To m
      If ee(i) = page Then
        Image8.Left = dd(i, 1, ee(i))
        Image8.Top = dd(i, 2, ee(i))
        Image8.Visible = True
      End If
    Next i

    For i = 1 To s2
      If ss(i) = page Then
        Image10.Left = sh(i, 1, ss(i))
        Image10.Top = sh(i, 2, ss(i))
      End If
    Next i

    For i = 1 To q2
      If qq(i) = page Then
        Image4(i).Left = q(i, 1, page)
        Image4(i).Top = q(i, 2, page)
        Image4(i).Visible = True
      Else
        Image4(i).Visible = False
      End If
    Next i

  End If
  
  If counter Mod 2 = 1 Then
    Timer1.Enabled = True
    Label1.Caption = "Pause"
  Else
    Timer1.Enabled = False
    Label1.Caption = "START"
  End If
  
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If p = 1 Then time = time + 1
  If KeyCode = vbKeyRight Then
    Image1.Picture = Image3.Picture
    For i = 1 To 1000
      xx = xx + 0.01
      vx = 10
    Next i
    
    Timer1.Enabled = True
  
  End If
  
  If KeyCode = vbKeyLeft Then
    Image1.Picture = Image2.Picture
    For i = 1 To 1000
      xx = xx - 0.01: vx = -10
    Next i
    Timer1.Enabled = True
  End If
  
  If KeyCode = vbKeySpace Then
    Shape4.Top = Image1.Top - Shape4.Height
    Shape4.Left = Image1.Left + Image1.Width / 2
    Shape4.Visible = True
    Timer2.Enabled = True
  End If
  
End Sub

Private Sub Timer1_Timer()

  For i = 1 To t
    Picture2.Line ((a(i, 1, page)) / 5, (a(i, 2, page)) / 5)-((a(i, 1, page)) / 5 + 20, (a(i, 2, page)) / 5), vbWhite
    Picture2.Line ((a(i, 1, page + 1)) / 5, (a(i, 2, page + 1)) / 5)-((a(i, 1, page + 1)) / 5 + 20, (a(i, 2, page + 1)) / 5), vbGreen
  Next i

  For i = 1 To g
    Picture2.Line ((b(i, 1, page)) / 5, (b(i, 2, page)) / 5)-((b(i, 1, page)) / 5 + 20, (b(i, 2, page)) / 5), vbWhite
    Picture2.Line ((b(i, 1, page + 1)) / 5, (b(i, 2, page + 1)) / 5)-((b(i, 1, page + 1)) / 5 + 20, (b(i, 2, page + 1)) / 5), brown
  Next i

  For i = 1 To f
    Picture2.FillColor = vbWhite: Picture2.Circle ((c(i, 1, page)) / 5, (c(i, 2, page) / 5)), 1, vbWhite
    Picture2.FillColor = vbBlue: Picture2.Circle ((c(i, 1, page + 1)) / 5, (c(i, 2, page + 1) / 5)), 1, vbBlack
  Next i

  For i = 1 To m
    
    If ee(i) = page + 1 Then
      Picture2.FillColor = vbCyan
      Picture2.Circle ((dd(i, 1, page + 1)) / 5, (dd(i, 2, page + 1)) / 5), 2, vbCyan
    End If
    
    If ee(i) = page Then
      Picture2.FillColor = vbWhite
      Picture2.Circle ((dd(i, 1, page)) / 5, (dd(i, 2, page)) / 5), 2, vbWhite
    End If
    
  Next i


  For i = 1 To s2
  
    If ss(i) = page + 1 Then
      Picture2.FillColor = vbYellow
      Picture2.Circle ((sh(i, 1, page + 1)) / 5, (sh(i, 2, page + 1)) / 5), 2, vbYellow
    End If
    
    If ss(i) = page Then
      Picture2.FillColor = vbWhite
      Picture2.Circle ((sh(i, 1, page)) / 5, (sh(i, 2, page)) / 5), 2, vbWhite
    End If
    
Next i

  If p = 1 Then
    time = time + 1
    Image9.Left = Image1.Left + Image1.Width / 2 - Image8.Width / 2
    Image9.Top = Image1.Top - Image8.Height
    Image9.Visible = True
    Image8.Visible = False
  End If
  
  If time < 300 Then ay = 0: vy = -(v)
  If time >= 300 Then
    ay = 15
    p = 0
    time = 1000
    Image8.Visible = False
    Image9.Visible = False
  End If
  
  If p2 = 1 Then
    time2 = time2 + 1
    If Image1.Picture = Image2.Picture Then Image11.Left = Image1.Left + Image1.Width
    If Image1.Picture = Image3.Picture Then Image11.Left = Image1.Left - Image11.Width
    Image11.Top = Image1.Top
    Image11.Visible = True
  End If
  
  If time2 < 300 Then
    ay = 0
    vy = -(2 * v)
  End If
  
  If time2 >= 300 Then
    ay = 15
    p2 = 0
    time2 = 1000
    Image10.Visible = False
    Image11.Visible = False
  End If
  
  s = 0
  
  d = (page - 1) * Picture1.ScaleHeight + (Picture1.ScaleHeight - Int(Image1.Top)) - Shape2.Height
  
  If d > he Then Label2.Caption = Int(d)
  
  he = Int(Val(Label2.Caption))
  
  If xx < -(Image1.Width) Then xx = Picture1.ScaleWidth
  If xx > Picture1.ScaleWidth Then xx = -(Image1.Width)
  
  If yy < -(Image1.Height) Then
    page = page + 1
    shatel = 0
    Shape2.Visible = False
  
  For i = 1 To t
    Image6(i).Left = a(i, 1, page)
    Image6(i).Top = a(i, 2, page)
  Next i

  For i = 1 To g
    Image5(i).Left = b(i, 1, page)
    Image5(i).Top = b(i, 2, page)
    Image5(i).Visible = True
  Next i

  For i = 1 To f
    Image7.Left = c(i, 1, page) - Image7.Width / 2
    Image7.Top = c(i, 2, page)
    Image7.Visible = True
  Next i

  For i = 1 To m
    If ee(i) = page - 1 Then Image8.Visible = False
    If ee(i) = page Then
      Image8.Left = dd(i, 1, page)
      Image8.Top = dd(i, 2, page)
      Image8.Visible = True
    End If
  Next i

  For i = 1 To s2
    If ss(i) = page - 1 Then Image10.Visible = False
    If ss(i) = page Then
      Image10.Left = sh(i, 1, ss(i))
      Image10.Top = sh(i, 2, ss(i))
      Image10.Visible = True
    End If
  Next i

  For i = 2 To q2
    If qq(i) = page Then Image4(qq(i)).Visible = True
  Next i

  yy = Picture1.ScaleHeight - 100
  
  End If

  If yy >= Picture1.ScaleHeight Then
    score = Val(Label2.Caption)
    page = page - 1
    Timer1.Enabled = False
    MsgBox "You lose" & Chr$(13) & Chr$(13) & Chr$(13) & "Your score :  " & score
    End
  End If

  For i = 1 To t
    vx = vx + ax * h
    If vy > 0 And yy + Image1.Height > a(i, 2, page) - 3 And yy + Image1.Height < a(i, 2, page) + 3 And xx + Image1.Width / 2 < a(i, 1, page) + 50 And xx + Image1.Width / 2 > a(i, 1, page) Then vy = -(v)
  Next i
  
  If yy + Image1.Height > Shape2.Top - 5 And yy + Image1.Height < Shape2.Top + 5 And Shape2.Visible = True Then
    vy = -(v)
  Else
    vx = vx + ax * h
    vy = vy + ay * h
    vy2 = vy2 + ay2 * h
  End If
  
  xx = xx + vx * h
  yy = yy + vy * h

  For i = 1 To g
    If vy > 0 And yy + Image1.Height > b(i, 2, page) - 3 And yy + Image1.Height < b(i, 2, page) + 3 And xx + Image1.Width / 2 < b(i, 1, page) + 50 And xx + Image1.Width / 2 > b(i, 1, page) Then Image5(i).Visible = False
  Next i

  For i = 1 To f
    If vy > 0 And yy + Image1.Height > c(i, 2, page) - 3 And yy + Image1.Height < c(i, 2, page) + 3 And xx + Image1.Width / 2 < c(i, 1, page) + 20 And xx + Image1.Width / 2 > c(i, 1, page) - 20 Then vy = -110
  Next i

  For i = 1 To m
    If ee(i) = page Then Image8.Left = dd(i, 1, page): Image8.Top = dd(i, 2, page): Image8.Visible = True
  Next i

  For i = 1 To m
    If ee(i) = page And yy + Image1.Height > dd(i, 2, page) - 3 And yy + Image1.Height < dd(i, 2, page) + 3 And xx + Image1.Width / 2 > dd(i, 1, page) - 20 And xx + Image1.Width / 2 < dd(i, 1, page) + 20 Then
      p = 1
      time = 1
      Image8.Visible = False
    End If
  Next i

  For i = 1 To s2
    If ss(i) = page And shatel <> 1 Then
      Image10.Left = sh(i, 1, page)
      Image10.Top = sh(i, 2, page)
      Image10.Visible = True
    End If
  Next i

  For i = 1 To s2
    If ss(i) = page And yy + Image1.Height > sh(i, 2, page) - 3 And yy + Image1.Height < sh(i, 2, page) + 3 And xx + Image1.Width / 2 > sh(i, 1, page) - 20 And xx + Image1.Width / 2 < sh(i, 1, page) + 20 Then
      p2 = 1
      time2 = 1
      Image10.Visible = False
      shatel = 1
    End If
  Next i

  For i = 1 To q2
    If qq(i) = page Then
      If q(i, 1, page) + Image4(1).Width >= Picture1.ScaleWidth Then vv = -(1)
      If q(i, 1, page) <= 0 Then vv = 1
      q(i, 1, page) = q(i, 1, page) + vv
    End If
  Next i

  For i = 1 To q2
    If Image4(i).Visible = True And qq(i) = page Then
      Image4(i).Left = q(i, 1, page)
      Image4(i).Top = q(i, 2, page)
      Image4(i).Visible = True
    Else
      Image4(i).Visible = False
    End If
  Next i

  For i = 1 To q2
    If vy <= 0 And Image4(i).Visible = True And yy + 5 > Image4(i).Top + Image4(i).Height - 1 And yy + 5 < Image4(i).Top + Image4(i).Height + 1 And xx + Image1.Width / 2 > Image4(i).Left And xx + Image1.Width / 2 < Image4(i).Left + Image4(i).Width Then
      Label2.Caption = "Bakhti"
      MsgBox "You lose" & Chr$(13) & Chr$(13) & Chr$(13) & "Your score :  " & score
      End
    End If
    
    If vy <= 0 And Image4(i).Visible = True And yy > Image4(i).Top And yy < Image4(i).Top + Image4(i).Height And xx + Image1.Width = Image4(i).Left Then Label2.Caption = "Bakhti": MsgBox "Bakhti": End
    If vy <= 0 And Image4(i).Visible = True And yy > Image4(i).Top And yy < Image4(i).Top + Image4(i).Height And xx = Image4(i).Left + Image4(i).Width Then Label2.Caption = "Bakhti": MsgBox "Bakhti": End
    If Image4(i).Visible = True And yy + Image1.Height > Image4(i).Top - 3 And yy + Image1.Height < Image4(i).Top + 3 And xx + Image1.Width >= Image4(i).Left And xx + Image1.Width <= Image4(i).Left + Image4(i).Width Then
      Image4(i).Visible = False
      vy = -100
    End If
  Next i

  Image1.Left = xx
  Image1.Top = yy
  
End Sub

Private Sub Timer2_Timer()

  Shape4.Top = Shape4.Top - 10
  
  If Shape4.Top <= -(Shape4.Height) Then
    Shape4.Visible = False
    Timer2.Enabled = False
  End If
  
  For i = 1 To q2
    If Image4(i).Visible = True And Shape4.Top > Image4(i).Top + Image4(i).Width - 5 And Shape4.Top < Image4(i).Top + Image4(i).Width + 5 And Shape4.Left > Image4(i).Left And Shape4.Left < Image4(i).Left + Image4(i).Width Then
      Image4(i).Visible = False
    End If
  Next i
  
End Sub
