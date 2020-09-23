VERSION 5.00
Object = "{102225D5-EA25-11D3-886E-00105A154A4D}#1.0#0"; "VPORTAL2.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   8415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6000
      TabIndex        =   11
      Text            =   "0"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Custom"
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Non-sensitive"
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Less sensitive"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sensitive"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   7
      Top             =   1320
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Very sensitive"
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   120
   End
   Begin VB.PictureBox Picture3 
      Height          =   1815
      Left            =   600
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   600
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   4560
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   3360
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   4560
      Width           =   2415
   End
   Begin VPORTAL2LibCtl.VideoPortal VideoPortal1 
      Height          =   3615
      Left            =   7440
      OleObjectBlob   =   "Form1.frx":009F
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "If you have any question feel free to ask I enjoy hearing from you, and specially when you're talking about this program."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   7920
      Width           =   10335
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":00C3
      Height          =   2295
      Left            =   6480
      TabIndex        =   17
      Top             =   4920
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "Number of changes"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "The live camera..."
      Height          =   255
      Left            =   7200
      TabIndex        =   15
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "The changes are recognizing here."
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "The First picture is comparing to the Second picture in the above PicBox."
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   6360
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   1695
   End
   Begin VB.OLE OLE1 
      Class           =   "SoundRec"
      Height          =   615
      Left            =   3960
      OleObjectBlob   =   "Form1.frx":03D8
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You need to have a Logitech digital camera and SDK package.

Dim p As Boolean
Dim b As Long
Dim bTEMP As Long
Dim i As Integer
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub Form_Load()
VideoPortal1.PrepareControl "QCSDK", "HKEY_LOCAL_MACHINE\Software\QCSDK", 0
VideoPortal1.EnableUIElements UIELEMENT_STATUSBAR, 0, 1
VideoPortal1.ConnectCamera2
VideoPortal1.EnablePreview = True
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then b = bTEMP - 25
If Index = 1 Then b = bTEMP
If Index = 2 Then b = bTEMP + 25
If Index = 3 Then b = bTEMP + 50
If Index = 4 Then b = Text2.Text
End Sub

Private Sub Timer1_Timer()
'command 3
Picture1.Picture = Picture2.Picture
VideoPortal1.PictureToFile 0, 24, "c:\aaa.bmp", ""
Picture2.Picture = LoadPicture("c:\aaa.bmp")
'command 2
Picture3.Cls
For y = 0 To 150
    For x = 0 To 200
        If GetPixel(Picture1.hdc, x, y) <> GetPixel(Picture2.hdc, x, y) Then
            SetPixel Picture3.hdc, x, y, GetPixel(Picture2.hdc, x, y)
            a = a + 1
        End If
    Next x
Next y
i = i + 1
If i = 20 Then Form1.Cls: i = 0
If p = False And i > 4 Then
    b = b + a
    If i = 9 Then
        b = b / 5
        b = b + 50
        bTEMP = b
        p = True
        Command1.Visible = False
    End If
End If
Text1.Text = b
Shape1.BackColor = &H8000000F
If a > b And p = True Then Shape1.BackColor = &HFF&: OLE1.DoVerb
Print a
End Sub
