VERSION 5.00
Begin VB.UserControl SearchBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   3  'I-Beam
   ScaleHeight     =   1860
   ScaleWidth      =   4080
   Begin VB.Timer tmrSearchInterval 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.TextBox txtSearchText 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "Click Here To Search"
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image imgSearchNoDropDownDisabled 
      Height          =   240
      Left            =   2880
      Picture         =   "uSearchPanel.ctx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSearchDropDownDisabled 
      Height          =   240
      Left            =   2520
      Picture         =   "uSearchPanel.ctx":03F9
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSearchClearDisabled 
      Height          =   240
      Left            =   3360
      Picture         =   "uSearchPanel.ctx":07F7
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSearchClearHover 
      Height          =   240
      Left            =   3600
      Picture         =   "uSearchPanel.ctx":0BA5
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSearchNoDropDown 
      Height          =   240
      Left            =   2880
      Picture         =   "uSearchPanel.ctx":0E2A
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSearchDropDown 
      Height          =   240
      Left            =   2520
      Picture         =   "uSearchPanel.ctx":122D
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSearchClear 
      Height          =   240
      Left            =   3360
      Picture         =   "uSearchPanel.ctx":1636
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSearchClearButton 
      Height          =   240
      Left            =   3720
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSearchDropDownButton 
      Height          =   240
      Left            =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "SearchBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' Default Property Values.
Private Const m_def_Enabled             As Boolean = True
Private Const m_def_SearchEventInterval As Long = 1000
Private Const m_def_BackStyle           As Long = 0
Private Const m_def_BorderColor         As Long = &HB99D7F
Private Const m_def_BorderStyle         As Long = 0
Private Const m_def_Height              As Long = 345
Private Const m_def_PopupMenu           As Boolean = True
Private Const m_def_Text                As String = vbNullString


' Property Variables.
Private m_Enabled                   As Boolean
Private m_BackStyle                 As Integer
Private m_BorderColor               As Long
Private m_BorderStyle               As Integer
Private m_PopupMenu                 As Boolean


' Event Declarations.
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Search(SearchString As String)
Event SearchCanceled()
Event PopupMenu()


Private m_SearchString          As String
Private m_ClearButtonHover      As Boolean



Private Sub UserControl_Initialize()
    ' Redraw the User Control.
    Call DrawControl
    Call ChangeCloseButtonState
End Sub
Private Sub UserControl_Resize()
    ' Redraw the User Control.
    Call DrawControl
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub UserControl_InitProperties()
    m_BorderColor = m_def_BorderColor
    m_Enabled = m_def_Enabled
    Set UserControl.Font = Ambient.Font
    UserControl.Height = m_def_Height
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_PopupMenu = m_def_PopupMenu
    txtSearchText.Text = m_def_Text
    Set txtSearchText.Font = UserControl.Font
    tmrSearchInterval.Interval = m_def_SearchEventInterval
    
    Call ProcessStateChanges
    Call ChangeCloseButtonState
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    UserControl.Height = PropBag.ReadProperty("Height", m_def_Height)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    m_PopupMenu = PropBag.ReadProperty("PopupMenu", m_def_PopupMenu)
    tmrSearchInterval.Interval = PropBag.ReadProperty("SearchInterval", m_def_SearchEventInterval)
    txtSearchText.Text = PropBag.ReadProperty("Text", m_def_Text)
    Set txtSearchText.Font = UserControl.Font
    txtSearchText.ForeColor = UserControl.ForeColor
    txtSearchText.BackColor = UserControl.BackColor
    txtSearchText.Enabled = m_Enabled

    ' Redraw the User Control.
    Call DrawControl
    Call ProcessStateChanges
    Call ChangeCloseButtonState
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Height", UserControl.Height, m_def_Height)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, True)
    Call PropBag.WriteProperty("PopupMenu", m_PopupMenu, m_def_PopupMenu)
    Call PropBag.WriteProperty("Text", txtSearchText.Text, m_def_Text)
    Call PropBag.WriteProperty("SearchInterval", tmrSearchInterval.Interval, m_def_SearchEventInterval)
End Sub


Private Sub tmrSearchInterval_Timer()
    tmrSearchInterval.Enabled = False
    
    ' Check to make sure its different then b4.
    If (m_SearchString <> txtSearchText.Text) Then
        m_SearchString = Trim$(txtSearchText.Text)
        RaiseEvent Search(m_SearchString)
    End If
End Sub


Private Sub txtSearchText_Change()
    ' Call the Search Start code for every change.
    If (LenB(Trim$(txtSearchText.Text)) > 0) Then
        Call SearchStart
    Else
        Call CancelSearch
    End If
End Sub
Private Sub txtSearchText_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        Call tmrSearchInterval_Timer
        KeyAscii = 0
    End If
End Sub

Private Sub imgSearchDropDownButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the left button has been pressed then trigger the PopupMenu Event
    If (Button = vbLeftButton) And (m_PopupMenu) And (m_Enabled) Then RaiseEvent PopupMenu
End Sub

Private Sub imgSearchClearButton_Click()
    ' Initiate the Clear Search.
    If (m_Enabled) Then Call CancelSearch
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Height
Public Property Get Height() As Long
    Height = UserControl.Height
End Property
Public Property Let Height(ByVal New_Height As Long)
    UserControl.Height = New_Height
    PropertyChanged "Height"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    txtSearchText.BackColor = New_BackColor
    PropertyChanged "BackColor"
    ' Redraw the User Control.
    Call DrawControl
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=0,0,0,&HB99D7F
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    ' Redraw the User Control.
    Call DrawControl
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    txtSearchText.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    txtSearchText.Enabled = m_Enabled
    PropertyChanged "Enabled"
    ' Initiate the Clear Search.
    Call ProcessStateChanges
    Call ChangeCloseButtonState
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Set txtSearchText.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property
Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    ' Redraw the User Control.
    Call DrawControl
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property
Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,True
Public Property Get PopupMenu() As Boolean
    PopupMenu = m_PopupMenu
End Property
Public Property Let PopupMenu(ByVal New_PopupMenu As Boolean)
    m_PopupMenu = New_PopupMenu
    PropertyChanged "PopupMenu"
    Call ProcessStateChanges
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Text() As String
    Text = txtSearchText.Text
End Property
Public Property Let Text(ByVal New_Text As String)
    txtSearchText.Text = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1000
Public Property Get SearchTiggerInterval() As String
    SearchTiggerInterval = tmrSearchInterval.Interval
End Property
Public Property Let SearchTiggerInterval(ByVal New_Interval As String)
    tmrSearchInterval.Interval = New_Interval
    PropertyChanged "SearchInterval"
End Property


Private Sub DrawControl()
    ' Clear the control graphics.
    UserControl.Cls
    
    ' Draw the controls borders.
    UserControl.Line (1, 0)-(ScaleWidth, 0), m_BorderColor
    UserControl.Line (1, ScaleHeight - 10)-(ScaleWidth, ScaleHeight - 10), m_BorderColor
    UserControl.Line (0, 0)-(0, ScaleHeight - 10), m_BorderColor
    UserControl.Line (ScaleWidth - 10, 0)-(ScaleWidth - 10, ScaleHeight - 10), m_BorderColor

    ' .
    imgSearchDropDownButton.Move 15, (UserControl.Height / 2) - (imgSearchDropDownButton.Height / 2)
    
    ' .
    imgSearchClearButton.Move (UserControl.Width - imgSearchClearButton.Width - 60), (UserControl.Height / 2) - (imgSearchClearButton.Height / 2)

    ' .
    txtSearchText.Move imgSearchDropDownButton.Left + imgSearchDropDownButton.Width + 30, (UserControl.Height / 2) - (txtSearchText.Height / 2)
    txtSearchText.Width = (UserControl.Width - 60) - (15 + txtSearchText.Left) - (imgSearchClearButton.Width + 15)
    txtSearchText.Height = UserControl.TextHeight("TEST")
End Sub


Private Sub SearchStart()
    ' Show the Clear Search Button and initiate the Search Event Timer.
    imgSearchClearButton.Visible = True
    tmrSearchInterval.Enabled = False
    tmrSearchInterval.Enabled = True
End Sub
Private Sub CancelSearch()
    ' Hide the Clear Search Button and clear the textbox.
    imgSearchClearButton.Visible = False
    txtSearchText.Text = vbNullString
    tmrSearchInterval.Enabled = False
    ' Raise the clear search event.
    RaiseEvent SearchCanceled
End Sub

Private Sub ProcessStateChanges()
    If (m_PopupMenu) Then
        Set imgSearchDropDownButton.Picture = IIf(m_Enabled, imgSearchDropDown.Picture, imgSearchDropDownDisabled.Picture)
    Else
        Set imgSearchDropDownButton.Picture = IIf(m_Enabled, imgSearchNoDropDown.Picture, imgSearchNoDropDownDisabled.Picture)
    End If
    txtSearchText.Enabled = m_Enabled
    UserControl.MousePointer = IIf(m_Enabled, vbIbeam, vbNormal)
End Sub

Private Sub ChangeCloseButtonState()
    If (m_Enabled) Then
        Set imgSearchClearButton.Picture = IIf(m_ClearButtonHover, imgSearchClearHover.Picture, imgSearchClear.Picture)
    Else
        Set imgSearchClearButton.Picture = imgSearchClearDisabled.Picture
    End If
End Sub

