VERSION 4.00
Begin VB.Form StartForm 
   BackColor       =   &H00008050&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   7650
   FillColor       =   &H00008060&
   Height          =   5370
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   7770
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "stephan.swertvaegher@planetinternet.be"
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Left            =   1980
      TabIndex        =   4
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   825
      Left            =   3075
      TabIndex        =   3
      Top             =   4005
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by Swertvaegher Stephan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   420
      Left            =   1980
      TabIndex        =   2
      Top             =   2115
      Width           =   3660
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed 1999"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   420
      Left            =   1980
      TabIndex        =   1
      Top             =   1800
      Width           =   3660
   End
   Begin VB.Image Image2 
      Height          =   3900
      Left            =   135
      Picture         =   "StartForm.frx":0000
      Top             =   495
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Icon-Album"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   2205
      TabIndex        =   0
      Top             =   900
      Width           =   3165
   End
   Begin VB.Image Image1 
      Height          =   3900
      Left            =   5805
      Picture         =   "StartForm.frx":20CF
      Top             =   495
      Width           =   1710
   End
End
Attribute VB_Name = "StartForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StartForm.Move (Screen.Width - StartForm.Width) / 2, (Screen.Height - StartForm.Height) / 2
Call ColForm(StartForm, 80, 128, 0, 50)
IconBook.Show
IconBook.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H80&
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H80&
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H80&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H80&
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H80&
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H80&
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartForm.Hide
IconBook.Show
IconBook.Enabled = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H50FF&
End Sub
