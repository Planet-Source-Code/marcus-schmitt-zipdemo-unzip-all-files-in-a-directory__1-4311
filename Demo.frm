VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unzip !"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Unzip"
      Height          =   375
      Left            =   2430
      TabIndex        =   3
      Top             =   3960
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   3720
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim AnzDateien%, Archiv$, t%, r&
  'Extract to same directory
  'You can change this directory here:
  ChDir Dir1.Path
  Dim na&, nb&
    For x = 0 To File1.ListCount - 1
    Archiv = File1.List(x)
    VBUnzip Archiv, CurDir, 0, 1, 0, 0, na, nb
    File1.ListIndex = x
  Next x
    File1.ListIndex = -1
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    Command1.Enabled = File1.ListCount > 0
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Sub Form_Load()
    With File1
        .Path = App.Path
        .Pattern = "*.zip"
    End With
End Sub
