VERSION 5.00
Object = "{3AE562E9-33E9-44B4-BBE5-F4A5AB980FCC}#3.0#0"; "SearchPanel.ocx"
Begin VB.Form frmMain 
   Caption         =   "FireFox Style - Search Bar"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin SearchBoxControl.SearchBox SearchBox3 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      ForeColor       =   255
      Object.Height          =   615
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Example Text"
   End
   Begin SearchBoxControl.SearchBox SearchBox2 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BorderColor     =   255
      BackColor       =   12632319
      Object.Height          =   375
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SearchBoxControl.SearchBox SearchBox1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      Object.Height          =   375
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnablePopup 
      Caption         =   "Disable Popup"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Disable Control"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtDebug 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Event Display:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Example of Customization:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   2235
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Event Trigger Interval:"
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   1650
   End
   Begin VB.Menu mnuPopupMenu 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupMenu_Item 
         Caption         =   "Subject"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPopupMenu_Item 
         Caption         =   "Sender"
         Index           =   1
      End
      Begin VB.Menu mnuPopupMenu_Item 
         Caption         =   "Subject & Sender"
         Index           =   2
      End
      Begin VB.Menu mnuPopupMenu_Item 
         Caption         =   "Message"
         Index           =   3
      End
      Begin VB.Menu mnuPopupMenu_Item 
         Caption         =   "Entire Message"
         Index           =   4
      End
      Begin VB.Menu mnuPopupMenu_Item 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPopupMenu_Item 
         Caption         =   "Save Search as Folder"
         Enabled         =   0   'False
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdEnable_Click()
    SearchBox1.Enabled = Not SearchBox1.Enabled
    cmdEnable.Caption = IIf(SearchBox1.Enabled, "Disable Control", "Enable Control")
End Sub

Private Sub cmdEnablePopup_Click()
    SearchBox1.PopupMenu = Not SearchBox1.PopupMenu
    cmdEnablePopup.Caption = IIf(SearchBox1.PopupMenu, "Disable Popup", "Enable Popup")
End Sub


Private Sub Form_Load()
    txtInterval.Text = CStr(SearchBox1.SearchTiggerInterval)
End Sub

Private Sub txtInterval_Change()
    SearchBox1.SearchTiggerInterval = Val(txtInterval.Text)
End Sub

Private Sub SearchBox1_PopupMenu()
    Call WriteDebug("Event Fired: POPUPMENU")
    PopupMenu mnuPopupMenu, , SearchBox1.Left, (SearchBox1.Top + SearchBox1.Height) - 15
End Sub

Private Sub SearchBox1_Search(SearchString As String)
    Call WriteDebug("Event Fired: SEARCH with Parameters '" & SearchString & "'")
End Sub

Private Sub SearchBox1_SearchCanceled()
    Call WriteDebug("Event Fired: SEARCHCANCELED")
End Sub


Private Sub WriteDebug(ByRef DebugText As String)
    txtDebug.Text = txtDebug.Text & DebugText & vbCrLf
End Sub
