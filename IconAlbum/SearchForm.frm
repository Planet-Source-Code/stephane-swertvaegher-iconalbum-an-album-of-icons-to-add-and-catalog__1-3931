VERSION 4.00
Begin VB.Form SearchForm 
   Caption         =   "Search Icons"
   ClientHeight    =   3570
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4755
   Height          =   3975
   Icon            =   "SearchForm.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   Top             =   1170
   Width           =   4875
   Begin VB.CommandButton Command1 
      Caption         =   "COPY"
      Height          =   420
      Left            =   3645
      TabIndex        =   3
      Top             =   1125
      Width           =   870
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00008000&
      Height          =   2985
      Left            =   1845
      Pattern         =   "*.ico"
      TabIndex        =   2
      Top             =   495
      Width           =   1500
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00008000&
      Height          =   2505
      Left            =   135
      TabIndex        =   1
      Top             =   990
      Width           =   1545
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   495
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   135
      TabIndex        =   4
      Top             =   45
      Width           =   4470
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3825
      Top             =   540
      Width           =   480
   End
End
Attribute VB_Name = "SearchForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click() 'copy to temp
On Error GoTo mkdir0
MkDir IBpath$ & "\" & "_tempicons"
IconBook.Dir1.Refresh
mkdir0:
FileCopy Label1.Caption, IBpath & "\" & "_tempicons" & "\" & File1.List(File1.ListIndex)
Command1.Enabled = False
IconBook.File1.Refresh
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Label1.Caption = Dir1.Path
CopyEnabled
End Sub

Private Sub Dir1_Click()
File1.Path = Dir1.List(Dir1.ListIndex)
Label1.Caption = Dir1.List(Dir1.ListIndex)
CopyEnabled
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Label1.Caption = Drive1.Drive
CopyEnabled
End Sub

Private Sub File1_Click()
If Right(File1.Path, 1) = "\" Then
Image1.Picture = LoadPicture(File1.Path + File1.List(File1.ListIndex))
Label1.Caption = File1.Path + File1.List(File1.ListIndex)
Else
Image1.Picture = LoadPicture(File1.Path + "\" + File1.List(File1.ListIndex))
Label1.Caption = File1.Path + "\" + File1.List(File1.ListIndex)
End If
CopyEnabled
End Sub

Private Sub Form_Activate()
Command1.Enabled = False
End Sub

Private Sub Form_Load()
SearchForm.Move (Screen.Width - SearchForm.Width) / 2, (Screen.Height - SearchForm.Height) / 2
Drive1.Drive = "c:"
Dir1.Path = "c:\"
File1.Path = "c:\"
Label1.Caption = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
SearchForm.Hide
IconBook.Show
End Sub

Private Sub CopyEnabled()
If File1.ListIndex = -1 Then
Command1.Enabled = False
Image1.Picture = LoadPicture("")
Else
Command1.Enabled = True
End If
End Sub
