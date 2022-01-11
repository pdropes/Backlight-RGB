VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6720
   ClientLeft      =   4395
   ClientTop       =   3930
   ClientWidth     =   7725
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H80000013&
   Icon            =   "RGB_LED.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   13440
      Top             =   4680
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      Height          =   4335
      Left            =   720
      MousePointer    =   2  'Cross
      Picture         =   "RGB_LED.frx":25CA
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   1
      Top             =   960
      Width           =   6255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      BaudRate        =   4800
   End
   Begin VB.Label EEprom 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "EEPROM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F00000&
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000CB000&
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000F0&
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "By Dropes 2022"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Blue_label 
      BackColor       =   &H8000000C&
      Caption         =   "Blue:    0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Green_label 
      BackColor       =   &H8000000C&
      Caption         =   "Green: 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Red_label 
      BackColor       =   &H8000000C&
      Caption         =   "Red:    0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000C&
      Caption         =   "Backlight"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Dim r As Long
Dim g As Long
Dim b As Long

Private Sub Form_Load()

   On Error GoTo ErTrap
    With MSComm1
        .CommPort = 8
        'set the badurate,parity,databits,stopbits for the connection
        .Settings = "4800,N,8,1"
        .InputMode = comInputModeBinary
        .DTREnable = True
        .RTSEnable = True
        'enable the oncomm event for every reveived character
        .RThreshold = 1
        'disable the oncomm event for send characters
        .SThreshold = 0
        'open the serial port
        .PortOpen = True
        .InBufferCount = 0
    End With    'MSComm1
    Exit Sub
    
ErTrap:
    If Err.Number = 8002 Then       ' Error '8002' is "Invalid Port Number"
       MsgBox "Your USB/Serial device is unplugged or is not recognized as COM8", vbExclamation
       Unload Me
    End If
End Sub

Private Sub Form_exit()
    MSComm1.InBufferCount = 0
    MSComm1.PortOpen = False
End Sub

Private Sub Save_Click()
    MSComm1.Output = Chr(3)    'identificador para gravar
    MSComm1.Output = Chr(3)
End Sub

Private Sub Picture1_Mousedown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Save_Backcolor As Long
    'Altera as cores do botão "Save"
    Save_Backcolor = (b * &H10000) + (g * &H100) + r
    EEprom.BackColor = Save_Backcolor
    EEprom.ForeColor = &HFFFFFF - Save_Backcolor

    'Correcção de cor
    b = b / 1.1
    g = g / 1.8
    r = r * 1.8
    If r > 255 Then r = 255
    '6 bit mode
    
    MSComm1.Output = Chr(0)
    MSComm1.Output = Chr(r)
    MSComm1.Output = Chr(1)
    MSComm1.Output = Chr(g)
    MSComm1.Output = Chr(2)
    MSComm1.Output = Chr(b)
End Sub

Private Sub Picture1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Debug.Print "x: " & x & "y: " & y

    Dim pixelColor As Long

    'Get the pixel color
    pixelColor = Picture1.Point(x, y)

    'Convert to RGB
    r = pixelColor And &HFF
    b = (pixelColor And &HFF0000) / &H10000
    g = (pixelColor - (b * &H10000) - r) / &H100

    Red_label = "Red:    " & r
    Green_label = "Green: " & g
    Blue_label = "Blue:    " & b
End Sub

