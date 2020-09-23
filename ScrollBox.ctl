VERSION 5.00
Begin VB.UserControl ScrollBox 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   ClipBehavior    =   0  'None
   ScaleHeight     =   285
   ScaleWidth      =   750
   ToolboxBitmap   =   "ScrollBox.ctx":0000
   Begin VB.VScrollBar vscScroll 
      Height          =   285
      Left            =   615
      Max             =   0
      Min             =   32767
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.HScrollBar hscScroll 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   285
      Width           =   615
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   0
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "ScrollBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default Property Values:
Const m_def_Min = 0
Const m_def_Max = 32767
Const m_def_Value = 0
Const m_def_SmallChange = 1
Const m_def_LargeChange = 1
Const m_def_ScrollBar = 0

'Property Variables:
Dim m_Min As Integer
Dim m_Max As Integer
Dim m_Value As Integer
Dim m_SmallChange As Integer
Dim m_LargeChange As Integer
Dim m_ScrollBar As Long

'Event Declarations:
Event Change() 'MappingInfo=txtValue,txtValue,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Sub hscScroll_Change()
    txtValue.Text = hscScroll.Value
    Value = hscScroll.Value
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "." Then
        KeyAscii = 0
    End If
    
    If IsNumeric(Chr(KeyAscii)) = False Then
        If (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyDelete) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    If m_ScrollBar = 0 Then
        UserControl.Width = txtValue.Width + vscScroll.Width
        UserControl.Height = txtValue.Height
    ElseIf m_ScrollBar = 1 Then
        UserControl.Width = txtValue.Width
        UserControl.Height = txtValue.Height + hscScroll.Height
    End If
End Sub

Private Sub UserControl_Resize()
    If m_ScrollBar = 0 Then
        UserControl.Width = txtValue.Width + vscScroll.Width
        UserControl.Height = txtValue.Height
    ElseIf m_ScrollBar = 1 Then
        UserControl.Width = txtValue.Width
        UserControl.Height = txtValue.Height + hscScroll.Height
    End If
End Sub

Private Sub UserControl_Show()
    If m_ScrollBar = 0 Then
        UserControl.Width = txtValue.Width + vscScroll.Width
        UserControl.Height = txtValue.Height
    ElseIf m_ScrollBar = 1 Then
        UserControl.Width = txtValue.Width
        UserControl.Height = txtValue.Height + hscScroll.Height
    End If
    
    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue.Text)
End Sub

Private Sub vscScroll_Change()
    txtValue.Text = vscScroll.Value
    Value = vscScroll.Value
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
Attribute Min.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    If New_Min < 0 Then New_Min = 0
    
    m_Min = New_Min
    vscScroll.Max = m_Min
    hscScroll.Min = m_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,32767
Public Property Get Max() As Integer
Attribute Max.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    m_Max = New_Max
    vscScroll.Min = m_Max
    hscScroll.Max = m_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of an object."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    txtValue.Text = CStr(m_Value)
    vscScroll.Value = m_Value
    hscScroll.Value = m_Value
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get SmallChange() As Integer
Attribute SmallChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks a scroll arrow."
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Integer)
    m_SmallChange = New_SmallChange
    vscScroll.SmallChange = m_SmallChange
    hscScroll.SmallChange = m_SmallChange
    PropertyChanged "SmallChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get LargeChange() As Integer
Attribute LargeChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks the scroll bar area."
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Integer)
    m_LargeChange = New_LargeChange
    vscScroll.LargeChange = m_LargeChange
    hscScroll.LargeChange = m_LargeChange
    PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub txtValue_Change()
    On Error Resume Next
    
    If txtValue.Text = "" Then
        txtValue.Text = "0"
        txtValue.SelStart = 0
        txtValue.SelLength = 1
    End If
    
    If CLng(txtValue.Text) > 32767 Then
        txtValue.Text = "32767"
        txtValue.SelStart = 0
        txtValue.SelLength = 5
    End If
    
    m_Value = Val(txtValue.Text)
    Value = m_Value
    RaiseEvent Change
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ScrollBar() As Long
Attribute ScrollBar.VB_Description = "Sets Verticle or Horizontal ScrollBar"
    ScrollBar = m_ScrollBar
End Property

Public Property Let ScrollBar(ByVal New_ScrollBar As Long)
    If New_ScrollBar > 1 Then New_ScrollBar = 1
    If New_ScrollBar < 0 Then New_ScrollBar = 0
    m_ScrollBar = New_ScrollBar
    PropertyChanged "ScrollBar"
    UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    m_SmallChange = m_def_SmallChange
    m_LargeChange = m_def_LargeChange
    m_ScrollBar = m_def_ScrollBar
    
    vscScroll.Min = m_Max
    vscScroll.Max = m_Min
    vscScroll.Value = m_Value
    vscScroll.SmallChange = m_SmallChange
    vscScroll.LargeChange = m_LargeChange
    
    hscScroll.Min = m_Min
    hscScroll.Max = m_Max
    hscScroll.Value = m_Value
    hscScroll.SmallChange = m_SmallChange
    hscScroll.LargeChange = m_LargeChange
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    Min = m_Min
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    Max = m_Max
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Value = m_Value
    m_SmallChange = PropBag.ReadProperty("SmallChange", m_def_SmallChange)
    SmallChange = m_SmallChange
    m_LargeChange = PropBag.ReadProperty("LargeChange", m_def_LargeChange)
    LargeChange = m_LargeChange
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_ScrollBar = PropBag.ReadProperty("ScrollBar", m_def_ScrollBar)
    ScrollBar = m_ScrollBar
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("SmallChange", m_SmallChange, m_def_SmallChange)
    Call PropBag.WriteProperty("LargeChange", m_LargeChange, m_def_LargeChange)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ScrollBar", m_ScrollBar, m_def_ScrollBar)
End Sub

