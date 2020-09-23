VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7725
   ClientLeft      =   12120
   ClientTop       =   8670
   ClientWidth     =   10035
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FileList 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   6360
      Width           =   2175
   End
   Begin VB.DirListBox DirList 
      Appearance      =   0  'Flat
      Height          =   990
      Left            =   135
      TabIndex        =   26
      Top             =   5280
      Width           =   2175
   End
   Begin VB.DriveListBox DriveList 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      TabIndex        =   25
      Top             =   4920
      Width           =   2175
   End
   Begin VB.ComboBox PatternCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   7080
      Width           =   2175
   End
   Begin VB.PictureBox picHidden 
      BackColor       =   &H00000000&
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   4755
      TabIndex        =   22
      Top             =   9120
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2650
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   2400
      ScaleHeight     =   2535
      ScaleWidth      =   5175
      TabIndex        =   23
      Top             =   4920
      Width           =   5175
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "www.ganzaborn.altervista.org"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   7785
      TabIndex        =   32
      ToolTipText     =   "Close INFO and Link"
      Top             =   5100
      Width           =   2130
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   7800
      Picture         =   "Form1.frx":0ECA
      Stretch         =   -1  'True
      Top             =   1785
      Width           =   2100
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5250
      TabIndex        =   31
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":Byte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2400
      TabIndex        =   30
      Top             =   7440
      Width           =   450
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6870
      TabIndex        =   29
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6975
      TabIndex        =   28
      Top             =   7440
      Width           =   585
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6540
      TabIndex        =   21
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":Megapixel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6645
      TabIndex        =   20
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   700
      TabIndex        =   19
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixels:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   570
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   1300
      Left            =   90
      TabIndex        =   17
      Top             =   3400
      Width           =   7500
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BASE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALTEZZA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1440
      TabIndex        =   15
      Top             =   2160
      Width           =   750
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASPECT RATIO ADATTATO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   2880
      TabIndex        =   14
      Top             =   2160
      Width           =   2130
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORMATO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   6780
      TabIndex        =   13
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASPECT RATIO IMMAGINE MODIFICATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   3540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASPECT RATIO ORIGINALE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   4080
      TabIndex        =   11
      Top             =   600
      Width           =   2085
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALTEZZA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   2640
      TabIndex        =   10
      Top             =   600
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BASE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASPECT RATIO IMMAGINE ORIGINALE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3420
   End
   Begin VB.Menu LINGUA 
      Caption         =   "LINGUA"
      Begin VB.Menu English 
         Caption         =   "English"
      End
      Begin VB.Menu Español 
         Caption         =   "Español"
      End
      Begin VB.Menu Italiano 
         Caption         =   "Italiano"
      End
   End
   Begin VB.Menu INFO 
      Caption         =   "INFO"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Const SRCCOPY = &HCC0020
Private ImageTypes(4) As String

Public Enum TipoArrotonda
    Difetto = 0
    Eccesso = 1
    Matematico = 2
End Enum

Function Arrotonda(Valore As Double, Arrotondamento As Double, Optional Direzione As TipoArrotonda = Eccesso) As Double
    On Error Resume Next
    Dim Temp As Double
    Temp = Valore / Arrotondamento
    If Int(Temp) = Temp Then
        Arrotonda = Valore
    Else
        Select Case Direzione
            Case TipoArrotonda.Difetto
                Temp = Int(Temp)
            Case TipoArrotonda.Eccesso
                Temp = Int(Temp) + 1
            Case TipoArrotonda.Matematico
                Temp = CDbl(Format(Temp, "0"))
        End Select
        Arrotonda = Temp * Arrotondamento
    End If
End Function

Private Sub English_Click()
LINGUA.Caption = "LANGUAGE"
Label1.Caption = UCase("aspect ratio original image")
Label2.Caption = UCase("base")
Label3.Caption = UCase("height")
Label4.Caption = UCase("aspect ratio original")
Label6.Caption = UCase("aspect ratio modified image")
Label10.Caption = UCase("BASE")
Label9.Caption = UCase("height")
Label8.Caption = UCase("aspect ratio adjusted")
Label7.Caption = UCase("FORMAT")
Combo1.Clear
Combo1.AddItem UCase("cut")
Combo1.AddItem "1.19:1 - Movietone '20"
Combo1.AddItem "1.25:1 - 5:4"
Combo1.AddItem "1.33:1 - 4:3"
Combo1.AddItem "1.37:1 - Cinema 1932-1953"
Combo1.AddItem "1.43:1 - IMAX"
Combo1.AddItem "1.50:1 - 3:2"
Combo1.AddItem "1.56:1 - 14:9"
Combo1.AddItem "1.66:1 - 5:3"
Combo1.AddItem "1.75:1 - Vistavision Cinema"
Combo1.AddItem "1.78:1 - 16:9"
Combo1.AddItem "1.85:1 - US/UK Wide Cinema"
Combo1.AddItem "2.00:1 - SuperScope"
Combo1.AddItem "2.20:1 - Todd-AO"
Combo1.AddItem "2.35:1 - Cinemascope"
Combo1.AddItem "2.39:1 - Cinemascope Panavision"
Combo1.AddItem "2.40:1 - Cinemascope Panavision"
Combo1.AddItem "2.55:1 - Cinemascope 55"
Combo1.AddItem "2.59:1 - Cinerama"
Combo1.AddItem "2.76:1 - MGM Camera 65"
Combo1.AddItem "4.00:1 - Polyvision"
Combo1.ListIndex = 0
End Sub

Private Sub Español_Click()
LINGUA.Caption = "IDIOMA"
Label1.Caption = UCase("relación de aspecto de la imagen original")
Label2.Caption = UCase("BASE")
Label3.Caption = UCase("altura")
Label4.Caption = UCase("relación de aspecto original")
Label6.Caption = UCase("relación de aspecto de la imagen modificada")
Label10.Caption = UCase("BASE")
Label9.Caption = UCase("altura")
Label8.Caption = UCase("relación de aspecto adaptado")
Label7.Caption = UCase("FORMATO")
Combo1.Clear
Combo1.AddItem UCase("corte")
Combo1.AddItem "1.19:1 - Movietone '20"
Combo1.AddItem "1.25:1 - 5:4"
Combo1.AddItem "1.33:1 - 4:3"
Combo1.AddItem "1.37:1 - Cinema 1932-1953"
Combo1.AddItem "1.43:1 - IMAX"
Combo1.AddItem "1.50:1 - 3:2"
Combo1.AddItem "1.56:1 - 14:9"
Combo1.AddItem "1.66:1 - 5:3"
Combo1.AddItem "1.75:1 - Vistavision Cinema"
Combo1.AddItem "1.78:1 - 16:9"
Combo1.AddItem "1.85:1 - US/UK Wide Cinema"
Combo1.AddItem "2.00:1 - SuperScope"
Combo1.AddItem "2.20:1 - Todd-AO"
Combo1.AddItem "2.35:1 - Cinemascope"
Combo1.AddItem "2.39:1 - Cinemascope Panavision"
Combo1.AddItem "2.40:1 - Cinemascope Panavision"
Combo1.AddItem "2.55:1 - Cinemascope 55"
Combo1.AddItem "2.59:1 - Cinerama"
Combo1.AddItem "2.76:1 - MGM Camera 65"
Combo1.AddItem "4.00:1 - Polyvision"
Combo1.ListIndex = 0
End Sub

Private Sub Italiano_Click()
LINGUA.Caption = "LINGUA"
Label1.Caption = UCase("ASPECT RATIO IMMAGINE ORIGINALE")
Label2.Caption = UCase("BASE")
Label3.Caption = UCase("ALTEZZA")
Label4.Caption = UCase("ASPECT RATIO ORIGINALE")
Label6.Caption = UCase("ASPECT RATIO IMMAGINE MODIFICATA")
Label10.Caption = UCase("BASE")
Label9.Caption = UCase("ALTEZZA")
Label8.Caption = UCase("ASPECT RATIO ADATTATO")
Label7.Caption = UCase("FORMATO")
Combo1.Clear
Combo1.AddItem UCase("TAGLIO")
Combo1.AddItem "1.19:1 - Movietone '20"
Combo1.AddItem "1.25:1 - 5:4"
Combo1.AddItem "1.33:1 - 4:3"
Combo1.AddItem "1.37:1 - Cinema 1932-1953"
Combo1.AddItem "1.43:1 - IMAX"
Combo1.AddItem "1.50:1 - 3:2"
Combo1.AddItem "1.56:1 - 14:9"
Combo1.AddItem "1.66:1 - 5:3"
Combo1.AddItem "1.75:1 - Vistavision Cinema"
Combo1.AddItem "1.78:1 - 16:9"
Combo1.AddItem "1.85:1 - US/UK Wide Cinema"
Combo1.AddItem "2.00:1 - SuperScope"
Combo1.AddItem "2.20:1 - Todd-AO"
Combo1.AddItem "2.35:1 - Cinemascope"
Combo1.AddItem "2.39:1 - Cinemascope Panavision"
Combo1.AddItem "2.40:1 - Cinemascope Panavision"
Combo1.AddItem "2.55:1 - Cinemascope 55"
Combo1.AddItem "2.59:1 - Cinerama"
Combo1.AddItem "2.76:1 - MGM Camera 65"
Combo1.AddItem "4.00:1 - Polyvision"
Combo1.ListIndex = 0
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Select Case LINGUA.Caption
Case "LANGUAGE"
If Combo1.Text = "CUT" Then
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Label11.Caption = ""
End If
If Combo1.Text = "1.19:1 - Movietone '20" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.19)
Text8.Text = "Movietone '20"
Label11.Caption = "Movietone format, used in the early sound films in 35 mm, in the late '20s, especially in Europe."
End If
If Combo1.Text = "1.25:1 - 5:4" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.25)
Text8.Text = "5:4 Photo"
Label11.Caption = "The British 405 line TV system used this aspect ratio since its introduction until 1950, when it was changed to the more common 1.33."
End If
If Combo1.Text = "1.33:1 - 4:3" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.33)
Text8.Text = "4:3 - Screen"
Label11.Caption = "35 mm original silent film ratio, commonly known in TV and video as 4:3."
End If
If Combo1.Text = "1.37:1 - Cinema 1932-1953" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.37)
Text8.Text = "Cinema 1932-1953"
Label11.Caption = "35 mm original silent film ratio, commonly known in TV and video as 4:3. Also standard ratio for MPEG-2 video compression. This format is still used in most personal video cameras today. It is the standard 16 mm and Super 35mm ratio."
End If
If Combo1.Text = "1.43:1 - IMAX" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.43)
Text8.Text = "IMAX"
Label11.Caption = "IMAX format. Imax productions use 70 mm wide film (the same as used for 70 mm feature films), but the film runs through the camera and projector sideways."
End If
If Combo1.Text = "1.50:1 - 3:2" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.5)
Text8.Text = "3:2 - Photo"
Label11.Caption = "The aspect ratio of 35 mm film used for still photography when 8 perforations are exposed. Usually called 3:2. Also the native aspect ratio of VistaVision."
End If
If Combo1.Text = "1.56:1 - 14:9" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.56)
Text8.Text = "14:9"
Label11.Caption = "Widescreen aspect ratio 14:9. Often used in shooting commercials etc. as a compromise format between 4:3 (12:9) and 16:9, especially when the output will be used in both standard TV and widescreen."
End If
If Combo1.Text = "1.66:1 - 5:3" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.66)
Text8.Text = "5:3 - Wide super 16 cinema"
Label11.Caption = "35 mm Originally a flat ratio invented by Paramount Pictures, now a standard among several European countries; native Super 16 mm frame ratio. (5:3, sometimes expressed more accurately as 1.67.)"
End If
If Combo1.Text = "1.75:1 - Vistavision Cinema" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.75)
Text8.Text = "Vistavision Cinema"
Label11.Caption = "Early 35 mm widescreen ratio, primarily used by MGM and Warner Bros. between 1953 and 1955, and since abandoned."
End If
If Combo1.Text = "1.78:1 - 16:9" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.78)
Text8.Text = "16:9 - Wide"
Label11.Caption = "Video widescreen standard (16:9), used in high-definition television, one of three ratios specified for MPEG-2 video compression. Also used in some personal video cameras."
End If
If Combo1.Text = "1.85:1 - US/UK Wide Cinema" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.85)
Text8.Text = "US/UK Wide Cinema"
Label11.Caption = "35 mm US and UK widescreen standard for theatrical film. Introduced by Universal Pictures in May, 1953. Projects approximately 3 perforations (perfs) of image space per 4 perf frame; films can be shot in 3-perf to save cost of film stock."
End If
If Combo1.Text = "2.00:1 - SuperScope" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2#)
Text8.Text = "SuperScope"
Label11.Caption = "Original SuperScope ratio, also used in Univisium. Used as a flat ratio for some American studios in the 1950s, abandoned in the 1960s, but recently popularized by the Red One camera system."
End If
If Combo1.Text = "2.20:1 - Todd-AO" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.2)
Text8.Text = "Todd-AO"
Label11.Caption = "70 mm standard. Originally developed for Todd-AO in the 1950s. 2.21:1 is specified for MPEG-2 but not used."
End If
If Combo1.Text = "2.35:1 - Cinemascope" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.35)
Text8.Text = "21:9 - Cinemascope"
Label11.Caption = "35 mm anamorphic prior to 1970, used by CinemaScope (Scope) and early Panavision. The anamorphic standard has subtly changed so that modern anamorphic productions are actually 2.39,[1] but often referred to as 2.35 anyway, due to old convention."
End If
If Combo1.Text = "2.39:1 - Cinemascope Panavision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.39)
Text8.Text = "Cinemascope Panavision"
Label11.Caption = "35 mm anamorphic from 1970 onwards. Sometimes rounded up to 2.40:1[1] Often commercially branded as Panavision format or 'Scope."
End If
If Combo1.Text = "2.40:1 - Cinemascope Panavision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.4)
Text8.Text = "Cinemascope Panavision"
Label11.Caption = "35 mm anamorphic from 1970 onwards."
End If
If Combo1.Text = "2.55:1 - Cinemascope 55" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.55)
Text8.Text = "Cinemascope 55"
Label11.Caption = "Original aspect ratio of CinemaScope before optical sound was added to the film in 1954. This was also the aspect ratio of CinemaScope 55."
End If
If Combo1.Text = "2.59:1 - Cinerama" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.59)
Text8.Text = "Cinerama"
Label11.Caption = "Cinerama at full height (three specially captured 35 mm images projected side-by-side into one composite widescreen image)."
End If
If Combo1.Text = "2.76:1 - MGM Camera 65" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.76)
Text8.Text = "MGM Camera 65"
Label11.Caption = "MGM Camera 65 (65 mm with 1.25x anamorphic squeeze). Used only on a handful of films between 1956 and 1964, such as Ben-Hur (1959)."
End If
If Combo1.Text = "4.00:1 - Polyvision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 4#)
Text8.Text = "Polyvision"
Label11.Caption = "Rare use of Polyvision, three 35 mm 1.33 images projected side by side."
End If
'--------------------------------------------------------------------------------------------------------
Case "IDIOMA"
If Combo1.Text = "CORTE" Then
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Label11.Caption = ""
End If
If Combo1.Text = "1.19:1 - Movietone '20" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.19)
Text8.Text = "Movietone '20"
Label11.Caption = "Movietone formato, utilizado en la primera película sonora en 35 mm, al final de la década de 1920, especialmente en Europa."
End If
If Combo1.Text = "1.25:1 - 5:4" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.25)
Text8.Text = "5:4 Photo"
Label11.Caption = "El sistema de televisión inglés en 405 líneas utiliza esta proporción desde su introducción hasta 1950, cuando se cambió a la 1,33 más comunes."
End If
If Combo1.Text = "1.33:1 - 4:3" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.33)
Text8.Text = "4:3 - Screen"
Label11.Caption = "Proporción original de cine mudo en 35 mm, usado comúnmente para la producción de televisión, donde más se conoce como 4: 3."
End If
If Combo1.Text = "1.37:1 - Cinema 1932-1953" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.37)
Text8.Text = "Cinema 1932-1953"
Label11.Caption = "Cine sonoro en 35 mm, prácticamente universal entre 1932 y 1953. Que se usa en ocasiones para producciones modernas y es el estándar para 16 mm."
End If
If Combo1.Text = "1.43:1 - IMAX" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.43)
Text8.Text = "IMAX"
Label11.Caption = "Formato IMAX. IMAX utiliza producciones cinematográficas de 70 mm, que a diferencia de los convencionales en 70 mm se desplaza horizontalmente, las cámaras de película a un área mayor del negativo."
End If
If Combo1.Text = "1.50:1 - 3:2" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.5)
Text8.Text = "3:2 - Photo"
Label11.Caption = "Relación de aspecto, utilizado para la fotografía en 35 mm, con medida de 24 x 36 mm."
End If
If Combo1.Text = "1.56:1 - 14:9" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.56)
Text8.Text = "14:9"
Label11.Caption = "También llamado 14: 9, a menudo se utiliza para la producción de la publicidad de vídeo, como un compromiso entre el 4: 3 y 16: 9."
End If
If Combo1.Text = "1.66:1 - 5:3" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.66)
Text8.Text = "5:3 - Wide super 16 cinema"
Label11.Caption = "Rapporto European Widescreen standard,  nativo de la película Super 16 mm (5: 3, a veces se expresa como (1,67))."
End If
If Combo1.Text = "1.75:1 - Vistavision Cinema" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.75)
Text8.Text = "Vistavision Cinema"
Label11.Caption = "Formato MetroScope (panorámica experimental en 35 mm), utilizado por la Metro-Goldwyn-Mayer y más tarde abandonado"
End If
If Combo1.Text = "1.78:1 - 16:9" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.78)
Text8.Text = "16:9 - Wide"
Label11.Caption = "Estándar de vídeo alta definición, comúnmente llamado el 16: 9."
End If
If Combo1.Text = "1.85:1 - US/UK Wide Cinema" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.85)
Text8.Text = "US/UK Wide Cinema"
Label11.Caption = "Panorámica relación estándar para las producciones de cine americana y británica. El marco utiliza aproximadamente 3 perforaciones de película en 4."
End If
If Combo1.Text = "2.00:1 - SuperScope" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2#)
Text8.Text = "SuperScope"
Label11.Caption = "Relación original SuperScope, usado también por el Univisium."
End If
If Combo1.Text = "2.20:1 - Todd-AO" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.2)
Text8.Text = "Todd-AO"
Label11.Caption = "Estándar de 70 mm, desarrollado originalmente para el sistema de Todd-AO, en la década de 1950."
End If
If Combo1.Text = "2.35:1 - Cinemascope" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.35)
Text8.Text = "21:9 - Cinemascope"
Label11.Caption = "35 mm anamorfico anterior al 1970, usado en el CinemaScope y en los primeros años del Panavision. El estándar anamorfico ha sido modificado ligeramente de modo que las producciones modernas tengan en realidad una relación de aspecto de 2,39,[1] aunque generalmente vienen llamadas igualmente 2,35, por tradición."
End If
If Combo1.Text = "2.39:1 - Cinemascope Panavision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.39)
Text8.Text = "Cinemascope Panavision"
Label11.Caption = "35 mm anamorfico siguiente al 1970, a veces redondeado a 2,40:1[1] a menudo llamado comercialmente formato Panavision."
End If
If Combo1.Text = "2.40:1 - Cinemascope Panavision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.4)
Text8.Text = "Cinemascope Panavision"
Label11.Caption = "35 mm anamorfico siguiente al 1970."
End If
If Combo1.Text = "2.55:1 - Cinemascope 55" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.55)
Text8.Text = "Cinemascope 55"
Label11.Caption = "La proporción de aspecto original de CinemaScope antes de adición de la pista de sonido óptico. También era la proporción de 55 de CinemaScope."
End If
If Combo1.Text = "2.59:1 - Cinerama" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.59)
Text8.Text = "Cinerama"
Label11.Caption = "Formado Cinerama a altura completa, tres imágenes 35 mm proyectadas de lado a lado sobre la pantalla panorámica."
End If
If Combo1.Text = "2.76:1 - MGM Camera 65" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.76)
Text8.Text = "MGM Camera 65"
Label11.Caption = "Formato de cámara MGM 65 (65 mm con compresión anamórficas de 1,25x)."
End If
If Combo1.Text = "4.00:1 - Polyvision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 4#)
Text8.Text = "Polyvision"
Label11.Caption = "Formato Polyvision, tres imágenes 35 mm con relación 1,33 proyectado de lado a lado."
End If
'--------------------------------------------------------------------------------------------------------
Case "LINGUA"
If Combo1.Text = "TAGLIO" Then
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Label11.Caption = ""
End If
If Combo1.Text = "1.19:1 - Movietone '20" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.19)
Text8.Text = "Movietone '20"
Label11.Caption = "Formato Movietone, usato nei primi film sonori in 35 mm, alla fine degli anni '20, soprattutto in Europa."
End If
If Combo1.Text = "1.25:1 - 5:4" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.25)
Text8.Text = "5:4 Photo"
Label11.Caption = "Il sistema televisivo inglese a 405 linee usava questo rapporto d'aspetto dalla sua introduzione fino al 1950, quando venne modificato nel più comune 1,33."
End If
If Combo1.Text = "1.33:1 - 4:3" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.33)
Text8.Text = "4:3 - Screen"
Label11.Caption = "Rapporto originale del cinema muto in 35 mm, usato comunemente per le produzioni televisive, dove è più noto come 4:3."
End If
If Combo1.Text = "1.37:1 - Cinema 1932-1953" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.37)
Text8.Text = "Cinema 1932-1953"
Label11.Caption = "Cinema sonoro in 35 mm, praticamente di impiego universale tra il 1932 e il 1953. È usato occasionalmente anche per produzioni moderne, e costituisce inoltre lo standard per il 16 mm."
End If
If Combo1.Text = "1.43:1 - IMAX" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.43)
Text8.Text = "IMAX"
Label11.Caption = "Formato IMAX. Le produzioni IMAX usano pellicola da 70 mm, che a differenza delle cineprese convenzionali in 70 mm viene fatta scorrere orizzontalmente, per una maggiore area del negativo."
End If
If Combo1.Text = "1.50:1 - 3:2" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.5)
Text8.Text = "3:2 - Photo"
Label11.Caption = "Rapporto d'aspetto usato per la fotografia in 35 mm, con fotogramma di 24×36 mm."
End If
If Combo1.Text = "1.56:1 - 14:9" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.56)
Text8.Text = "14:9"
Label11.Caption = "Chiamato anche 14:9, è spesso usato per la produzione di filmati pubblicitari, come un formato di compromesso tra il 4:3 e il 16:9."
End If
If Combo1.Text = "1.66:1 - 5:3" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.66)
Text8.Text = "5:3 - Wide super 16 cinema"
Label11.Caption = "Rapporto European Widescreen standard, nativo per la pellicola Super 16 mm (5:3, espresso talvolta come (1,67))."
End If
If Combo1.Text = "1.75:1 - Vistavision Cinema" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.75)
Text8.Text = "Vistavision Cinema"
Label11.Caption = "Formato MetroScope (panoramico sperimentale in 35 mm), usato dalla Metro-Goldwyn-Mayer e in seguito abbandonato."
End If
If Combo1.Text = "1.78:1 - 16:9" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.78)
Text8.Text = "16:9 - Wide"
Label11.Caption = "Rapporto standard per il video ad alta definizione, chiamato comunemente 16:9."
End If
If Combo1.Text = "1.85:1 - US/UK Wide Cinema" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 1.85)
Text8.Text = "US/UK Wide Cinema"
Label11.Caption = "Rapporto panoramico standard per le produzioni cinematografiche americane e inglesi. Il fotogramma usa all'incirca l'altezza di 3 perforazioni di pellicola su 4."
End If
If Combo1.Text = "2.00:1 - SuperScope" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2#)
Text8.Text = "SuperScope"
Label11.Caption = "Rapporto originale SuperScope, usato anche per l'Univisium."
End If
If Combo1.Text = "2.20:1 - Todd-AO" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.2)
Text8.Text = "Todd-AO"
Label11.Caption = "Standard 70 mm, sviluppato in origine per il sistema Todd-AO negli anni 1950."
End If
If Combo1.Text = "2.35:1 - Cinemascope" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.35)
Text8.Text = "21:9 - Cinemascope"
Label11.Caption = "35 mm anamorfico precedente al 1970, usato nel CinemaScope e nei primi anni del Panavision. Lo standard anamorfico è stato modificato leggermente in modo che le produzioni moderne abbiano in realtà un rapporto d'aspetto di 2,39,[1]  anche se vengono di solito chiamate ugualmente 2,35, per tradizione."
End If
If Combo1.Text = "2.39:1 - Cinemascope Panavision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.39)
Text8.Text = "Cinemascope Panavision"
Label11.Caption = "35 mm anamorfico successivo al 1970, a volte arrotondato a 2,40:1[1] Spesso chiamato commercialmente formato Panavision."
End If
If Combo1.Text = "2.40:1 - Cinemascope Panavision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.4)
Text8.Text = "Cinemascope Panavision"
Label11.Caption = "35 mm anamorfico successivo al 1970."
End If
If Combo1.Text = "2.55:1 - Cinemascope 55" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.55)
Text8.Text = "Cinemascope 55"
Label11.Caption = "Rapporto d'aspetto originale del CinemaScope prima dell'aggiunta della traccia audio ottica. Era inoltre il rapporto del CinemaScope 55."
End If
If Combo1.Text = "2.59:1 - Cinerama" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.59)
Text8.Text = "Cinerama"
Label11.Caption = "Formato Cinerama ad altezza piena (tre immagini 35 mm proiettate fianco a fianco sullo schermo panoramico)."
End If
If Combo1.Text = "2.76:1 - MGM Camera 65" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 2.76)
Text8.Text = "MGM Camera 65"
Label11.Caption = "Formato MGM Camera 65 (65 mm con compressione anamorfica 1,25x)."
End If
If Combo1.Text = "4.00:1 - Polyvision" Then
Text5.Text = CInt(Text1.Text)
Text6.Text = CInt(Text1.Text / 4#)
Text8.Text = "Polyvision"
Label11.Caption = "Formato Polyvision, tre immagini 35 mm con rapporto 1,33 proiettate fianco a fianco."
End If
End Select
Text7.Text = dblEuro2Dec(Text5.Text / Text6.Text) & ":1"
End Sub

Private Sub Form_Load()
Form1.Height = 8385
Form1.Width = 7785
Me.Caption = "AspectRatio Calculator" & " v" & App.Major & "." & App.Minor & "." & App.Revision & " - " & "03/03/2010" & " - " & "13/07/2012"
    PatternCombo.AddItem "Graphic (*.gif;*.jpg;*.bmp)"
    PatternCombo.AddItem "JPEG (*.jpg)"
    PatternCombo.AddItem "Bitmaps (*.bmp)"
    PatternCombo.AddItem "GIF (*.gif)"
    PatternCombo.AddItem "All Files (*.*)"
    PatternCombo.ListIndex = 0
    DriveList.Drive = App.Path
    DirList.Path = App.Path
    picHidden.AutoSize = True
    picHidden.Visible = False
Form1.Width = 7785
Italiano_Click
End Sub

Function SetBytes(Bytes) As String

On Error GoTo hell

If Bytes >= 1073741824 Then
    SetBytes = Format(Bytes / 1024 / 1024 / 1024, "#0.00") _
         & " GB"
ElseIf Bytes >= 1048576 Then
    SetBytes = Format(Bytes / 1024 / 1024, "#0.00") & " MB"
ElseIf Bytes >= 1024 Then
    SetBytes = Format(Bytes / 1024, "#0.00") & " KB"
ElseIf Bytes < 1024 Then
    SetBytes = Fix(Bytes) & " Bytes"
End If

Exit Function
hell:
SetBytes = "0 Bytes"
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label16.BackColor = &H80000005
Label16.ForeColor = &H0&
End Sub

Private Sub PatternCombo_Click()
Dim pat As String
Dim p1 As Integer
Dim p2 As Integer
    pat = PatternCombo.List(PatternCombo.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    FileList.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub

Private Sub INFO_Click()
If INFO.Caption = "INFO" Then
Form1.Width = 10155
Form1.WindowState = 0
INFO.Caption = "Close INFO"
    Else
Form1.Width = 7785
Form1.WindowState = 0
INFO.Caption = "INFO"
    End If
End Sub

Private Sub Label16_Click()
Dim ApriPaginaWeb As Long
Form1.Width = 7785
Form1.WindowState = 0
INFO.Caption = "INFO"
    ApriPaginaWeb = ShellExecute(Me.hwnd, vbNullString, Label16.Caption, vbNullString, "c:\", 1)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label16.BackColor = &HFF8080
Label16.ForeColor = &HFFFFFF
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label16.BackColor = &HFF8080
End Sub

Private Sub Text1_Change()
On Error Resume Next
Dim a As Long
Dim ore As String
a = Text1.Text * Text2.Text
Text3.Text = Text1.Text / Text2.Text
Text3.Text = dblEuro2Dec(Text3.Text) & ":1"
Label13.Caption = Format(a, "#,#")
Label13.Caption = Replace(Label13.Caption, ".", ",")
Label15.Caption = FormatNumber(Text1.Text * Text2.Text / 1000000, 1)
Label15.Caption = Replace(Label15.Caption, ",", ".")
Combo1_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr$(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Text2_Change()
On Error Resume Next
Dim a As Long
Dim ore As String
a = Text1.Text * Text2.Text
Text3.Text = Text1.Text / Text2.Text
Text3.Text = dblEuro2Dec(Text3.Text) & ":1"
Label13.Caption = Format(a, "#,#")
Label13.Caption = Replace(Label13.Caption, ".", ",")
Label15.Caption = FormatNumber(Text1.Text * Text2.Text / 1000000, 1)
Label15.Caption = Replace(Label15.Caption, ",", ".")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr$(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub
Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub
Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub
Private Sub Text7_GotFocus()
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)
End Sub
Private Sub Text8_GotFocus()
Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)
End Sub

Private Sub FitPictureToBox(ByVal pic_src As PictureBox, pic_dst As PictureBox)
Dim aspect_src As Single
Dim wid As Single
Dim hgt As Single
    aspect_src = pic_src.ScaleWidth / pic_src.ScaleHeight
    wid = pic_dst.ScaleWidth
    hgt = pic_dst.ScaleHeight
    If wid / hgt > aspect_src Then
        wid = aspect_src * hgt
    Else
        hgt = wid / aspect_src
    End If
    pic_dst.Cls
    On Error Resume Next
    pic_dst.PaintPicture pic_src.Picture, _
        (pic_dst.ScaleWidth - wid) / 2, _
        (pic_dst.ScaleHeight - hgt) / 2, _
        wid, hgt
End Sub

Private Sub DirList_Change()
    FileList.Path = DirList.Path
End Sub

Private Sub DriveList_Change()
On Error GoTo DriveError:
    DirList.Path = DriveList.Drive
    Exit Sub
DriveError:
If LINGUA.Caption = "IDIOMA" Then
MsgBox "UNIDAD VACIA", vbExclamation
End If
If LINGUA.Caption = "LINGUA" Then
MsgBox "DISCO VUOTO", vbExclamation
End If
If LINGUA.Caption = "LANGUAGE" Then
MsgBox "DRIVE EMPTY", vbExclamation
End If
    DriveList.Drive = DirList.Path
    Exit Sub
End Sub

Private Sub FileList_Click()
Dim fname As String
 On Error GoTo LoadPictureError
    ImageTypes(0) = "Unknown"
    ImageTypes(1) = "GIF"
    ImageTypes(2) = "JPEG"
    ImageTypes(3) = "PNG"
    ImageTypes(4) = "BMP"
    ReadImageInfo (FileList.Path & "\" & FileList.FileName)
    Text2.Text = ImageHeight
    Text1.Text = ImageWidth
    Label17.Caption = ImageTypes(ImageType)
    Label18.Caption = FileSize & " :Byte" & " - (" & SetBytes(FileSize) & ")"
    '----------------------------------------------
    fname = FileList.Path & "\" & FileList.FileName
    MousePointer = vbHourglass
    DoEvents
    picHidden.Picture = LoadPicture(fname)
    FitPictureToBox picHidden, picImage
    MousePointer = vbDefault
    Exit Sub
    Combo1_Click
LoadPictureError:
    Beep
    MousePointer = vbDefault
    Caption = "Viewer [Invalid picture]"
    Exit Sub
End Sub

Public Function dblEuro2Dec(dblCifra As Double) As Double
Dim strNumero As String, strIntero As String, sDecimali As String
Dim sChar As String * 1
Dim iPosVirg As Integer, iCnt As Integer
Dim dblNumElab As Double
    ' trasforma il numero in stringa
    strNumero = CStr(dblCifra)
    ' rileva la parte intera del numero
    strIntero = CStr(Int(dblCifra))
    ' SE il separatore decimale è "." lo converto in ","
    ' per compatibilità con tutte le impostazioni internazionali
    strNumero = Replace(strNumero, ".", ",")
    ' rileva la posizione della virgola
    iPosVirg = InStr(strNumero, ",")
    ' SE non ci sono decimali
    If (iPosVirg = 0) Then
       dblEuro2Dec = dblCifra
       Exit Function
    End If
    ' rileva la parte decimale del numero
    sDecimali = Mid(strNumero, iPosVirg + 1, Len(strNumero) - iPosVirg)
    Select Case Len(sDecimali)
      Case Is = 2
        dblEuro2Dec = dblCifra
        Exit Function
      Case Is > 6
        ' tronca a 6 il numeri di decimali
        sDecimali = Left(sDecimali, 6)
    End Select
    ' il numero da elaborare è la variabile "dblNumElab",
    ' cui assegnamo il valore: 0,nnnnn
    dblNumElab = Val("0." & sDecimali)
    ' esegue la funzione di arrotondamento iniziando dal decimale in coda
    ' fino all'ultimo arrotondabile, che è il secondo
    For iCnt = Len(sDecimali) To 3 Step -1
      iPosVirg = InStr(dblNumElab, ",")
      sDecimali = Mid(dblNumElab, iPosVirg + 1, Len(dblNumElab) - iPosVirg)
      If (Len(sDecimali) < 3) Then Exit For
      sChar = Mid(sDecimali, iCnt, 1)
      sDecimali = Left(sDecimali, Len(sDecimali) - 1)
      dblNumElab = Val("0." & sDecimali)
      If (Val(sChar) > 4) Then
        dblNumElab = dblNumElab + (1 / 10 ^ (iCnt - 1))
      End If
    Next iCnt
    ' assegna alla funzione il numero arrotonadato
    dblEuro2Dec = Val(strIntero) + dblNumElab
End Function
