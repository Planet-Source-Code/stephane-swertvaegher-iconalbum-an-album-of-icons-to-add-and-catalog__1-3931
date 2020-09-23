VERSION 4.00
Begin VB.Form IconBook 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "IconAlbum     1999 by Swertvaegher Stephan"
   ClientHeight    =   7275
   ClientLeft      =   975
   ClientTop       =   1560
   ClientWidth     =   10950
   Height          =   7680
   Icon            =   "IconBook.frx":0000
   Left            =   915
   LinkTopic       =   "Form1"
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   730
   Top             =   1215
   Width           =   11070
   Begin VB.Frame Frame1 
      Caption         =   "Commands"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1545
      Left            =   90
      TabIndex        =   4
      Top             =   5670
      Width           =   2265
      Begin VB.CommandButton Command2 
         Caption         =   "Search for Icons"
         Height          =   330
         Left            =   315
         TabIndex        =   6
         Top             =   990
         Width           =   1635
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Make new directory"
         Height          =   330
         Left            =   315
         TabIndex        =   5
         Top             =   540
         Width           =   1635
      End
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6765
      Left            =   2430
      ScaleHeight     =   6765
      ScaleWidth      =   8385
      TabIndex        =   2
      Top             =   405
      Width           =   8385
      Begin VB.PictureBox Pic1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   180
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   4980
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   2310
   End
   Begin VB.FileListBox File1 
      Height          =   3570
      Left            =   2340
      Pattern         =   "*.ico"
      TabIndex        =   0
      Top             =   2925
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   45
      TabIndex        =   8
      Top             =   5085
      Width           =   2310
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2430
      TabIndex        =   7
      Top             =   45
      Width           =   8340
   End
   Begin MSComDlg.CommonDialog ComD1 
      Left            =   10395
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".ico"
      Flags           =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Icon As"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Move Icon"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Icon"
      End
   End
End
Attribute VB_Name = "IconBook"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click() 'Make new dir
On Error GoTo mkdir0
Temp$ = InputBox("Type the name of the new directory" & vbCr & "you want to create.", "IconAlbum")
If Temp$ = "" Then Pic2.SetFocus: Exit Sub
MkDir IBpath$ & "\" & Temp$
Dir1.Refresh
Pic2.SetFocus
mkdir0:
End Sub

Private Sub Command2_Click()
SearchForm.Show 1
Pic2.SetFocus
End Sub

Private Sub Dir1_Change()
Dir1.Path = IBpath
End Sub

Private Sub Dir1_Click()
For xx% = 0 To 179
Pic1(xx%).Picture = LoadPicture("")
Pic1(xx%).Visible = False
Next xx%
File1.Path = Dir1.List(Dir1.ListIndex)

If File1.ListCount <= 180 Then
Idx% = File1.ListCount - 1
Else
Idx% = 179
End If
Label1.Caption = Dir1.List(Dir1.ListIndex)
Label2.Caption = "Icons in map:" & vbCr & Idx% + 1
For xx% = 0 To Idx%
If Right(File1.Path, 1) = "\" Then
Pic1(xx%).Picture = LoadPicture(File1.Path + File1.List(xx%))
Else
Pic1(xx%).Picture = LoadPicture(File1.Path + "\" + File1.List(xx%))
End If
Pic1(xx%).Visible = True
Next xx%
End Sub

Private Sub Form_Activate()
'Pic2.SetFocus
End Sub

Private Sub Form_Load()
Call ColForm(Pic2, 128, 148, 96, 50)
IBpath = App.Path
IconBook.Move (Screen.Width - IconBook.Width) / 2, (Screen.Height - IconBook.Height) / 2
For xx = 1 To 179
Load Pic1(xx%)
Next xx
For xx = 0 To 14
Pic1(xx%).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 15).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 30).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 45).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 60).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 75).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 90).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 105).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 120).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 135).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 150).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx% + 165).Left = Pic1(0).Left + (xx% * 36)
Pic1(xx%).Top = Pic1(0).Top
Pic1(xx% + 15).Top = Pic1(0).Top + 36
Pic1(xx% + 30).Top = Pic1(0).Top + 72
Pic1(xx% + 45).Top = Pic1(0).Top + 108
Pic1(xx% + 60).Top = Pic1(0).Top + 144
Pic1(xx% + 75).Top = Pic1(0).Top + 180
Pic1(xx% + 90).Top = Pic1(0).Top + 216
Pic1(xx% + 105).Top = Pic1(0).Top + 252
Pic1(xx% + 120).Top = Pic1(0).Top + 288
Pic1(xx% + 135).Top = Pic1(0).Top + 324
Pic1(xx% + 150).Top = Pic1(0).Top + 360
Pic1(xx% + 165).Top = Pic1(0).Top + 396
Next xx%
Dir1.Path = IBpath
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub mnuDelete_Click()
Temp$ = MsgBox("Are you sure you want to delete the icon: " & vbCr & vbCr & Dir1.List(Dir1.ListIndex) & "\" & File1.List(IconIdx%), vbQuestion + vbYesNo, "IconAlbum - System message")
If Temp$ = vbNo Then Exit Sub
Kill Dir1.List(Dir1.ListIndex) & "\" & File1.List(IconIdx%)
File1.Refresh
Dir1_Click

End Sub

Private Sub mnuMove_Click()
Dim Oldpath$, Newpath$
Oldpath$ = Dir1.List(Dir1.ListIndex)
On Error GoTo mnuMove2
inp$ = InputBox("You want to move the icon:" & vbCr & Dir1.List(Dir1.ListIndex) & "\" & File1.List(IconIdx%) & vbCr & vbCr & "Type the name of the destination map:", "IconAlbum", inp$)
If inp$ = "" Then Exit Sub
FileCopy Oldpath$ & "\" & File1.List(IconIdx%), IBpath & "\" & inp$ & "\" & File1.List(IconIdx%)
Kill Oldpath & "\" & File1.List(IconIdx%)
File1.Refresh
Dir1_Click
Exit Sub
mnuMove2:
If Err = 76 Then
Temp$ = MsgBox("The map " & inp$ & " does not exist !" & vbCr & vbCr & "Do you want me to create it ?", vbExclamation + vbYesNo, "IconAlbum - System message")
If Temp$ = vbNo Then Exit Sub
MkDir IBpath & "\" & inp$
Dir1.Refresh
FileCopy Oldpath$ & "\" & File1.List(IconIdx%), IBpath & "\" & inp$ & "\" & File1.List(IconIdx%)
Kill Oldpath$ & "\" & File1.List(IconIdx%)
File1.Refresh
Dir1_Click

Exit Sub
End If
Mess$ = MsgBox("There's a copy error !", vbCritical + vbOKOnly, "IconAlbum - System Message")
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo NoSave
ComD1.filename = File1.List(IconIdx%)
ComD1.DialogTitle = "Save Icon"
ComD1.ShowSave
SavePicture Pic1(IconIdx%).Picture, ComD1.filename
NoSave:

End Sub

Private Sub Pic1_Click(Index As Integer)
IconIdx% = Index
PopupMenu mnuFile
End Sub

