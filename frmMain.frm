VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteration Security - www.intrepid.vze.com"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6765
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2760
      Top             =   2880
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4440
      TabIndex        =   7
      Text            =   "5"
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox chkStartup 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4440
      TabIndex        =   5
      Top             =   5760
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3840
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Secure a file..."
      Filter          =   "*.*"
      InitDir         =   "C:\"
   End
   Begin MSComctlLib.ImageList ilListview 
      Left            =   3240
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6064
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D566
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove file"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Scan file"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Secure a new file"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin prjAlerteration.cSysTray SysTray 
      Left            =   4680
      Top             =   2640
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmMain.frx":DF78
      TrayTip         =   "Alteration Security"
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilListview"
      SmallIcons      =   "ilListview"
      ColHdrIcons     =   "ilListview"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Original Path"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Backup Path"
         Object.Width           =   14112
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "minute(s)."
      Height          =   195
      Left            =   4800
      TabIndex        =   8
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Check every:"
      Height          =   195
      Left            =   3240
      TabIndex        =   6
      Top             =   6000
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Run on startup:"
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngMin As Long
Public Function GetNameOnly(strFile As String) As String
GetNameOnly = Mid(strFile, InStrRev(strFile, "\"), Len(strFile) - InStrRev(strFile, "\") + 1)
GetNameOnly = GetNameOnly
End Function



Public Sub RemoveFile(Filename As String, Backup As String)
Dim a As String
Open "C:\Backup.lst" For Input As #2
Line Input #2, a$
Close #2
a$ = Replace(a$, Filename & "|/\|" & Backup, "")
Open "C:\Backup.lst" For Output As #2
Print #2, a$
Close #2
On Error Resume Next
Kill Backup
End Sub

Public Sub ScanFile(File As String)
Dim frmA As Form, filelen1 As Long, filelen2 As Long
filelen1 = FileLen(File)
filelen2 = FileLen(App.Path & "\Backup" & GetNameOnly(File))
If filelen1 <> filelen2 Then
'INFECTION POSSIBLE!
Set frmA = New frmAlert
frmA.Label2.Caption = File
frmA.Show
Else
lvAddItem File, 3, False
End If

End Sub

Public Sub SecureFile(Filename As String, Backup As String)
Open "C:\Backup.lst" For Append As #1
Print #1, Filename & "|/\|" & Backup
Close #1
FileCopy Filename, Backup
End Sub


Private Sub chkStartup_Click()
Dim x As String, y As String, q As String
x = App.Path & "\" & "AlterS" & ".exe"
y = GetStartupMenu & "AlterS.exe"
'On Error Resume Next
If chkStartup.Value = vbChecked Then
FileCopy x, y
Else
Kill GetStartupMenu & "\AlterS.exe"
End If

End Sub

Public Function GetStartupMenu() As String
    Dim lpStartupPath As String * MAX_PATH
    Dim Pidl As Long
    Dim hResult As Long
    
    hResult = SHGetSpecialFolderLocation(0, CSIDL_COMMON_STARTUP, Pidl)


    If hResult = 0 Then
        hResult = SHGetPathFromIDList(ByVal Pidl, lpStartupPath)


        If hResult = 1 Then
            lpStartupPath = Left(lpStartupPath, InStr(lpStartupPath, Chr(0)) - 1)
            GetStartupMenu = lpStartupPath
            GetStartupMenu = Mid(GetStartupMenu, 1, InStrRev(GetStartupMenu, "\"))
            GetStartupMenu = GetStartupMenu & "Startup\"
        End If
    End If
End Function
Private Sub Command1_Click()
On Error GoTo 1
cd.ShowOpen
If cd.Filename <> "" Then
    lvAddItem cd.Filename
    cd.Filename = ""
End If
Exit Sub
1:
MsgBox Err.Description, vbCritical, "An error accured!"
End Sub

Public Sub lvAddItem(Path As String, Optional Icon As Integer = 2, Optional blnNew As Boolean = True)
On Error Resume Next

Dim item As ListItem, i As Integer
If blnNew = True Then
Set item = lvFiles.ListItems.Add()
item.Text = Path
item.SubItems(1) = App.Path & "\Backup" & GetNameOnly(Path)
item.Icon = Icon
item.SmallIcon = Icon
SecureFile Path, item.SubItems(1)
Else
For i = 1 To lvFiles.ListItems.Count
    If lvFiles.ListItems.item(i).Text = Path Then
    Set item = lvFiles.ListItems.item(i)
    GoTo 1
    End If
Next i
Exit Sub
1:
item.Text = Path
item.SubItems(1) = App.Path & "\Backup" & GetNameOnly(Path)
item.Icon = Icon
item.SmallIcon = Icon
End If
End Sub

Public Sub lvLoadItem(Path As String, Backup As String, Optional Icon As Integer = 2)
On Error Resume Next

Dim item As ListItem, i As Integer
Set item = lvFiles.ListItems.Add()
item.Text = Path
item.SubItems(1) = App.Path & "\Backup" & GetNameOnly(Path)
item.Icon = Icon
item.SmallIcon = Icon
Exit Sub

End Sub

Private Sub Command2_Click()
On Error Resume Next
lvAddItem lvFiles.SelectedItem.Text, 1, False
ScanFile lvFiles.SelectedItem.Text
End Sub

Private Sub Command3_Click()
On Error GoTo 1
RemoveFile lvFiles.SelectedItem.Text, lvFiles.SelectedItem.SubItems(1)
lvFiles.ListItems.Remove (lvFiles.SelectedItem.Index)
1:
End Sub


Private Sub Form_Load()
Dim b() As String, a As String, n As Long
On Error Resume Next
Open "C:\Backup.lst" For Input As #1
Do While EOF(1) = False
Line Input #1, a$
b = Split(a$, "|/\|")
lvLoadItem b(0), b(1)
ScanFile b(0)
DoEvents
Loop
1:
Close #1
Text1.Text = GetSetting(App.Title, "Settings", "CheckEvery", "5")
On Error Resume Next
n = 0
n = FileLen(GetStartupMenu & "\AlterS.exe")
If n > 0 Then
chkStartup.Value = vbChecked
Else
chkStartup.Value = 0
End If
App.TaskVisible = True
End Sub


Private Sub Form_Resize()
If WindowState = vbMinimized Then
Hide
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
SysTray.InTray = False
End
End Sub


Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
Show
End Sub

Private Sub Text1_Change()
SaveSetting App.Title, "Settings", "CheckEvery", CLng(Text1.Text)
Text1.Text = CLng(Text1.Text)
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
lngMin = lngMin + 1
If lngMin >= CLng(Text1.Text) Then
lngMin = 0
    If lvFiles.ListItems.Count = 0 Then Exit Sub
    For i = 1 To lvFiles.ListItems.Count
    ScanFile lvFiles.ListItems(i).Text
    DoEvents
    Next i
End If

End Sub


