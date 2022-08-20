VERSION 5.00
Object = "{F6D0E404-2D34-11D2-9A46-0080C8763FA4}#1.0#0"; "DATEEDITBOX.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2040
      Width           =   1575
   End
   Begin DATEEDITBOXLib.BHM_DateMaskEditbox BHM_DateMaskEditbox1 
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1085
      _StockProps     =   109
      ForeColor       =   255
      BackColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateFormat      =   "ggg/mm/dd"
      DisplayDateFormat=   "ggg/mm/dd"
      ReadOnly        =   0   'False
      TextAlign       =   0
      CtlStop         =   -1  'True
      CtlIndex        =   0
      CanYearNegative =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BHM_DateMaskEditbox1_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode
End Sub

Private Sub BHM_DateMaskEditbox1_KeyPress(KeyAscii As Integer)
Debug.Print KeyAscii

End Sub
