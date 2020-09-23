VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Simple Fade Efect Button"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   715
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   3
      Left            =   360
      TabIndex        =   19
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   8421504
      ButtonStyle     =   "Form1.frx":1E5A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 nextButton 
      Height          =   315
      Left            =   6960
      TabIndex        =   18
      Top             =   360
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      HighlightColor  =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      Aceleration     =   0
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3495
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "Form1.frx":5AF7
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6120
      Top             =   240
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   192
      ButtonStyle     =   "Form1.frx":5E2D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   16576
      ButtonStyle     =   "Form1.frx":9E1B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   8388736
      ButtonStyle     =   "Form1.frx":DF1F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   7680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   8388608
      ButtonStyle     =   "Form1.frx":1225D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   5
      Left            =   2640
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   8421376
      ButtonStyle     =   "Form1.frx":162B2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   6
      Left            =   2640
      TabIndex        =   6
      Top             =   6000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   32768
      ButtonStyle     =   "Form1.frx":1A3D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   6840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   8599650
      ButtonStyle     =   "Form1.frx":1DD26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   8
      Left            =   2640
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   16512
      ButtonStyle     =   "Form1.frx":22027
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   9
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   16384
      ButtonStyle     =   "Form1.frx":25F80
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   10
      Left            =   360
      TabIndex        =   10
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   16512
      ButtonStyle     =   "Form1.frx":29FF7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   11
      Left            =   2640
      TabIndex        =   1
      Top             =   6840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   4194368
      ButtonStyle     =   "Form1.frx":2DC77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   12
      Left            =   2640
      TabIndex        =   0
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ButtonStyle     =   "Form1.frx":31DBB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   13
      Left            =   10440
      TabIndex        =   14
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      HighlightColor  =   16777215
      ForeColor       =   8388608
      ButtonStyle     =   "Form1.frx":33BBB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin Project1.UserControl1 Button 
      Height          =   630
      Index           =   14
      Left            =   12600
      TabIndex        =   15
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1111
      NumeroClips     =   12
      Efect           =   1
      HighlightColor  =   16777215
      ForeColor       =   192
      ButtonStyle     =   "Form1.frx":37C10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Click Me"
   End
   Begin VB.Label Labelmario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pill Colors"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   17
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Always Effect               MouseMoveEffect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   10560
      TabIndex        =   16
      Top             =   2280
      Width           =   3960
   End
   Begin VB.Label Labelmario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pill Buttons Fade Effect By Mario Flores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Color As Integer
Dim Decolorate As Boolean
Dim Acelerate As Integer


'***********************Esta es la misma tecnica que use en el usercontrol**************************
'***************************************************************************************************



Private Sub Form_Activate()
Color = 0
Acelerate = 10
Decolorate = True 'Es Oscuro (Default )!Hay que decolorarlo
nextButton.Left = Me.ScaleWidth - nextButton.Width - 50 ' Posicion al Final de Hoja
nextButton.Top = Me.ScaleHeight - nextButton.Height - 50 ' Posicion al Final de Hoja
End Sub



Private Sub nextButton_Click()
Unload Me
MsgBox "Code by MARIO FLORES G " & vbCrLf & "sistec_de_juarez@hotmail.com", vbInformation, "Exit"
End
End Sub

Private Sub Timer1_Timer()


'***********************Esta es la misma tecnica que use en el usercontrol**************************
'***************************************************************************************************


If Decolorate = True Then

    For i = 0 To Acelerate
        If Color <= 255 Then
           Color = Color + 1
           Labelmario(0).ForeColor = RGB(Color, Color, 255)
           Labelmario(1).ForeColor = RGB(255, Color, Color)
        Else
           Decolorate = Not Decolorate
           Exit Sub
        End If
        
    Next i
        
'Nota: This is not FlickerLess

End If

If Decolorate = False Then
    
    For i = 0 To Acelerate
        If Color > 1 Then
           Color = Color - 1
           Labelmario(0).ForeColor = RGB(Color, Color, 255)
           Labelmario(1).ForeColor = RGB(255, Color, Color)
        Else
           Decolorate = Not Decolorate
           Exit Sub
        End If
    Next i
        
    
    
End If

End Sub

Private Sub Button_Click(Index As Integer)
MsgBox "I'm a Cool Button", vbInformation, "Button # " & Index
End Sub

