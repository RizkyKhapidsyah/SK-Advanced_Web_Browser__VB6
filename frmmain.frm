VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Insane Programmers Web Browser"
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10110
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1800
      Top             =   5520
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   480
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   0
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Urlbox 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   8655
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6615
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11818
            MinWidth        =   864
            Picture         =   "frmmain.frx":0442
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   9495
      ExtentX         =   16748
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0946
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":105A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":131E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":17EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":21F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":267E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2FB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   926
      ButtonWidth     =   1482
      ButtonHeight    =   873
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Back"
            Key             =   "back"
            Description     =   "back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Forward"
            Key             =   "forward"
            Description     =   "forward"
            Object.ToolTipText     =   "&Forward"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Stop"
            Key             =   "stop"
            Description     =   "stop"
            Object.ToolTipText     =   "&Stop"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh"
            Key             =   "refresh"
            Description     =   "refresh"
            Object.ToolTipText     =   "&Refresh"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Home"
            Key             =   "home"
            Description     =   "home"
            Object.ToolTipText     =   "&Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Favorites"
            Key             =   "fav"
            Description     =   "Favorites"
            Object.ToolTipText     =   "&Favorites"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&History"
            Key             =   "history"
            Description     =   "history"
            Object.ToolTipText     =   "&History"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print"
            Key             =   "print"
            Description     =   "print"
            Object.ToolTipText     =   "&Print"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mail"
            Key             =   "mail"
            Description     =   "mail"
            Object.ToolTipText     =   "&Mail"
            ImageIndex      =   6
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "send"
                  Text            =   "&Send Mail"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "check"
                  Text            =   "Check mail"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&About"
            Key             =   "about"
            Description     =   "about"
            Object.ToolTipText     =   "&About"
            ImageIndex      =   9
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Menu wbip 
      Caption         =   "WbIP"
      Enabled         =   0   'False
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mode 
         Caption         =   "&Work mode"
         Begin VB.Menu online 
            Caption         =   "&Online"
            Checked         =   -1  'True
         End
         Begin VB.Menu offline 
            Caption         =   "&Offline"
         End
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu search 
         Caption         =   "&Search Internet"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu show 
         Caption         =   "&Show"
         Begin VB.Menu statusbar 
            Caption         =   "&Statusbar"
            Checked         =   -1  'True
         End
         Begin VB.Menu toolbar 
            Caption         =   "&Toolbar"
            Checked         =   -1  'True
         End
         Begin VB.Menu urls 
            Caption         =   "Urls"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu stopp 
         Caption         =   "&Stop"
      End
      Begin VB.Menu refreshh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu backk 
         Caption         =   "&Go Back"
      End
      Begin VB.Menu forwardd 
         Caption         =   "&Go Forward"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu source 
         Caption         =   "&Source"
      End
      Begin VB.Menu properties 
         Caption         =   "Properties"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mail 
      Caption         =   "&Mail"
      Begin VB.Menu check 
         Caption         =   "&Check Mail"
      End
      Begin VB.Menu send 
         Caption         =   "&Send Mail"
      End
   End
   Begin VB.Menu helpp 
      Caption         =   "&Help"
      Begin VB.Menu help 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu about 
         Caption         =   "&About"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu a 
      Caption         =   ""
   End
   Begin VB.Menu b 
      Caption         =   ""
   End
   Begin VB.Menu d 
      Caption         =   ""
   End
   Begin VB.Menu e 
      Caption         =   ""
   End
   Begin VB.Menu f 
      Caption         =   ""
   End
   Begin VB.Menu g 
      Caption         =   ""
   End
   Begin VB.Menu h 
      Caption         =   ""
   End
   Begin VB.Menu fav 
      Caption         =   ""
      Begin VB.Menu add 
         Caption         =   "&Add"
      End
      Begin VB.Menu viewfav 
         Caption         =   "&View"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub add_Click()
Dim addfav
    addfav = InputBox("Enter website you wish to add to favorites", "Add", "www.hotmail.com")
        If addfav = "" Then
    Exit Sub
        Else:
        favv.List1.AddItem (addfav)
End If
End Sub

Private Sub backk_Click()
    Web.GoBack
End Sub

Private Sub check_Click()
    Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
End Sub

Private Sub copy_Click()
On Error Resume Next
    Clipboard.GetText Web.Document
End Sub

Private Sub cut_Click()
On Error Resume Next
    Clipboard.GetData (Web.Document)
    End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
frmabout.show
On Error Resume Next
    iniPath$ = App.Path & "/web.dll"
    Dim starting
    starting = GetFromINI("main", "home", iniPath$)
    Web.Navigate (starting)
    StatusBar1.Panels(1).Text = "Ready."
    StatusBar1.Panels(2).Text = "Online."
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Web.Width = frmmain.Width - 100
    Web.Height = frmmain.Height - 1900
    Urlbox.Width = frmmain.Width - 3000
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub forwardd_Click()
    Web.GoForward
End Sub



Private Sub help_Click()
MsgBox "Need help than get ahold of me at:" & vbCrLf & "ICQ: 13237186" & vbCrLf & " OR " & vbCrLf & "Email: hax03d@hotmail.com", vbInformation, "Help"
End Sub

Private Sub offline_Click()
If online.Checked = True Then
    online.Checked = False
    offline.Checked = True
    Web.offline = True
    StatusBar1.Panels(2).Text = "Offline"
    StatusBar1.Panels(1).Text = "You can now work offline"
End If
End Sub

Private Sub online_Click()
If offline.Checked = True Then
    online.Checked = True
    Web.offline = False
    offline.Checked = False
    StatusBar1.Panels(1).Text = ""
    StatusBar1.Panels(2).Text = "Online"
End If
End Sub

Private Sub open_Click()
On Error Resume Next
    Com.Filter = "All Internet Files (*.hmt,*.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml"
    Com.ShowOpen
If Com.Filename = "" Then
    Exit Sub
Else
    Web.Navigate (Com.Filename)
End If
End Sub

Private Sub paste_Click()
On Error Resume Next
    Clipboard.SetText (Clipboard.GetData)
End Sub

Private Sub properties_Click()
frmoptions.show
End Sub

Private Sub refreshh_Click()
    Web.Refresh
End Sub

Private Sub save_Click()
    Com.Filter = "lnk file (*.lnk) | *.lnk"
    Com.ShowSave
If Com.Filename = "" Then
    Exit Sub
Else
    Open Com.Filename For Output As #1
     Print #1, Web.Document
   Close #1
End If
End Sub

Private Sub search_Click()
    iniPath$ = App.Path & "/web.dll"
Dim search
    search = GetFromINI("Search", "url", iniPath$)
    Web.Navigate (search)
    Urlbox.Text = search
    Urlbox.AddItem (search)
End Sub

Private Sub send_Click()
Dim subject, person
    person = InputBox("Enter email address", "email")
    subject = InputBox("Enter subject for email", "subject")
    Web.Navigate ("mailto:" & person & "?subject=" & subject)
End Sub

Private Sub source_Click()
On Error Resume Next
Open App.Path & "/source.tmp" For Output As #1
    Print #1, Inet1.OpenURL(Web.LocationURL)
Close #1
    Shell "C:\windows\notepad.exe " & App.Path & "/source.tmp", vbNormalFocus
    Kill App.Path & "/source.tmp"
End Sub

Private Sub statusbar_Click()
If statusbar.Checked = True Then
    statusbar.Checked = False
    StatusBar1.Visible = False
Else
    statusbar.Checked = True
    StatusBar1.Visible = True
End If
End Sub

Private Sub stopp_Click()
    Web.Stop
End Sub

Private Sub Timer1_Timer()
Unload frmabout
Me.WindowState = 2
Timer1.Enabled = False
End Sub

Private Sub toolbar_Click()
If toolbar.Checked = True Then
    toolbar.Checked = False
    Toolbar1.Visible = False
Else
    toolbar.Checked = True
    Toolbar1.Visible = True
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim starting
On Error Resume Next
Select Case Button.Key
    Case "back"
     Web.GoBack
    Case "forward"
     Web.GoForward
    Case "stop"
     Web.Stop
    Case "refresh"
     Web.Refresh
    Case "home"
        iniPath$ = App.Path & "/web.dll"
        starting = GetFromINI("main", "home", iniPath$)
        Web.Navigate (starting)
    Case "fav"
     PopupMenu fav
    Case "print"
     Print Web.Document
    Case "mail"
     PopupMenu mail
    Case "about"
     frmabout.show
    Case "history"
     Call ShellExecute(hwnd, "Open", "C:\Windows\History\", "", App.Path, 1)
End Select
End Sub
Private Sub Urlbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Web.Navigate (Urlbox.Text)
    Urlbox.AddItem (Urlbox.Text)
End If
End Sub

Private Sub urls_Click()
If urls.Checked = True Then
    urls.Checked = False
    Urlbox.Visible = False
Else
    urls.Checked = True
    Urlbox.Visible = True
End If
End Sub

Private Sub viewfav_Click()
    favv.show
End Sub

Private Sub Web_DocumentComplete(ByVal pDisp As Object, Url As Variant)
    StatusBar1.Panels(1).Text = "Document Finished."
End Sub


Private Sub Web_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
    frmmain.Caption = Web.LocationName
    StatusBar1.Panels(1).Text = Web.LocationURL
    Urlbox.Text = Web.LocationURL
End Sub

Private Sub Web_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
Stat.show
Stat.ProgressBar1.Max = ProgressMax
Stat.ProgressBar1.Value = Progress
If Progress = 0 Then
Stat.Hide
Else:
Stat.show
End If
End Sub
