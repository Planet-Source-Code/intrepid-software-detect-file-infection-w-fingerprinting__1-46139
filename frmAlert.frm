VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Possible Infection!"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Allow"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restore"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "If you are sure this file was not intentionally modified by yourself click RESTORE to restore the file with its backup."
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6405
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "The following file has a possible infection:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2925
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<filename>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6450
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
WindowState = vbMinimized
FileCopy App.Path & "\Backup" & frmMain.GetNameOnly(Label2.Caption), Label2.Caption
frmMain.lvAddItem Label2.Caption, 3, False
Unload Me
End Sub


Private Sub Command2_Click()
WindowState = vbMinimized
frmMain.RemoveFile Label2.Caption, App.Path & "\Backup" & frmMain.GetNameOnly(Label2.Caption)
frmMain.SecureFile Label2.Caption, App.Path & "\Backup" & frmMain.GetNameOnly(Label2.Caption)
frmMain.lvAddItem Label2.Caption, 3, False
Unload Me
End Sub


