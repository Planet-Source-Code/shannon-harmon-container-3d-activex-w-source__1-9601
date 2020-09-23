VERSION 5.00
Begin VB.UserControl Container3D 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1710
   ControlContainer=   -1  'True
   ScaleHeight     =   840
   ScaleWidth      =   1710
   ToolboxBitmap   =   "Container3D.ctx":0000
   Begin VB.Timer tmrMouseEnter 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   105
      Top             =   405
   End
End
Attribute VB_Name = "Container3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Enum ControlEffects
  Flat = 0
  Raised = 1
  Raised3D = 2
  Inset = 3
  Bump = 4
  Etched = 5
End Enum

Public Enum AutoSize
  None = 0
  [Auto-size Child] = 1
End Enum

Public Enum BackStyle
  Transparent = 0
  Opaque = 1
End Enum

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10
Private Const BF_MIDDLE = &H800
Private Const BF_SOFT = &H1000
Private Const BF_ADJUST = &H2000
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
    qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

'Default Property Values:
Const m_def_AutoSize = 0
Const m_def_CtrlEffect = 1
'Property Variables:
Dim m_AutoSize As AutoSize
Dim m_CtrlEffect As ControlEffects
'Event Declarations:
Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the mouse first enters the Container3D's focus area."
Event MouseExit()
Attribute MouseExit.VB_Description = "Occurs when the mouse leaves the Container3D's focus area."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Type PointAPI
  x As Long
  Y As Long
End Type

Private Sub tmrMouseEnter_Timer()
  If Not IsMouseOver Then
    tmrMouseEnter.Enabled = False
    RaiseEvent MouseExit
  End If
End Sub

Private Function IsMouseOver() As Boolean
  Dim MousePos As PointAPI
  Dim rc As Long
    
  On Error Resume Next
  rc = GetCursorPos(MousePos)
  If WindowFromPoint(MousePos.x, MousePos.Y) = UserControl.hWnd Then
    IsMouseOver = True
  End If
End Function


Private Sub DrawControl(Optional blnPainting As Boolean = False)
On Error Resume Next

  Dim r As RECT
  Dim oldScale As Long
  oldScale = UserControl.ScaleMode  'Save Old ScaleMode So We Can Set It Back To Previous Later
  UserControl.ScaleMode = vbPixels  'Must Be Set To vbPixels For This To Work
  UserControl.Cls 'Clear Our Control
  
  r.Left = UserControl.ScaleLeft
  r.Top = UserControl.ScaleTop
  r.Right = UserControl.ScaleWidth
  r.Bottom = UserControl.ScaleHeight

  Select Case m_CtrlEffect  'Draw Based On Our CtrlEffect Property
    Case Flat
      'Nothing needed the cls does the trick!
      'Could draw a flat edge but there is no point...
      'DrawEdge hdc, r, EDGE_SUNKEN, BF_FLAT
      
    Case Inset
      DrawEdge hdc, r, EDGE_SUNKEN, BF_RECT
    
    Case Raised3D
      DrawEdge hdc, r, EDGE_RAISED, BF_RECT
      
    Case Bump
      DrawEdge hdc, r, EDGE_BUMP, BF_RECT
    
    Case Etched
      DrawEdge hdc, r, EDGE_ETCHED, BF_RECT
    
    Case Raised
      'No API To Create This With That I Know Of So I Just Draw Lines??? <- Look Into More
      Line (0, 0)-(UserControl.ScaleWidth, 0), vbWhite
      Line (0, 0)-(0, UserControl.ScaleHeight), vbWhite
      Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), vbButtonShadow
      Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), vbButtonShadow
  End Select
  
  If Not blnPainting Then 'Don't Do If Control Has Just Painted Or It Will Put It Into A Loop
    If m_AutoSize = [Auto-size Child] Then  'If AutoSize Is Enabled
      If UserControl.ContainedControls.Count = 1 Then 'Only Do If We Have Only One ContainedControl
        Dim iOffSet As Integer, lWidth As Long, lHeight As Long
        iOffSet = 30  'Offset Space To Align The ContainedControl
        Err.Clear 'Clear Any Errors
        'If The ContainedControl Does Not Have A Width Or Height Property This Will Cause An Error
        lWidth = UserControl.ContainedControls(0).Width
        lHeight = UserControl.ContainedControls(0).Height
        If Err.Number = 0 Then  'If There Were No Errors Size The ContainedControl
          UserControl.ContainedControls(0).Left = iOffSet
          UserControl.ContainedControls(0).Top = iOffSet
          UserControl.ContainedControls(0).Height = UserControl.Height - (iOffSet * 2)
          UserControl.ContainedControls(0).Width = UserControl.Width - (iOffSet * 2)
        End If
      End If
    End If
  End If
  
  UserControl.ScaleMode = oldScale  'Set ScaleMode Back To Previous
  If UserControl.AutoRedraw Then UserControl.Refresh  'Refresh If Needed
End Sub
  

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/Sets the Container3D's to be active or not.  When inactive, all child controls will also become inactive."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

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
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
  MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
  UserControl.MousePointer() = New_MousePointer
  PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CtrlEffect() As ControlEffects
Attribute CtrlEffect.VB_Description = "Returns/Sets the effect used for the Container3D when drawing the Container3D."
  CtrlEffect = m_CtrlEffect
End Property

Public Property Let CtrlEffect(ByVal New_CtrlEffect As ControlEffects)
  On Error Resume Next
  If New_CtrlEffect < 0 Or New_CtrlEffect > 5 Then GoTo PROC_ERR
  
  m_CtrlEffect = New_CtrlEffect
  PropertyChanged "CtrlEffect"
  DrawControl
  
  Exit Property

PROC_ERR:
  On Error GoTo 0
  Err.Raise _
    Number:=(514), _
    Source:="Container3D.CtrlEffect", _
    Description:="Index Out Of Bounds"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
  DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get AutoSize() As AutoSize
Attribute AutoSize.VB_Description = "Returns/Sets Autosize value.  When true, if a single control is a child of the Container3D and has Height & Width Properties it will be size to fit the maximum area of the Container3D.  Ignores if there is more than one child control."
  AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As AutoSize)
  On Error Resume Next
  If New_AutoSize < 0 Or New_AutoSize > 1 Then GoTo PROC_ERR
  
  m_AutoSize = New_AutoSize
  PropertyChanged "AutoSize"
  DrawControl
  Exit Property

PROC_ERR:
  On Error GoTo 0
  Err.Raise _
    Number:=(514), _
    Source:="Container3D.CtrlEffect", _
    Description:="Index Out Of Bounds"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As BackStyle
  BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyle)
  UserControl.BackStyle() = New_BackStyle
  PropertyChanged "BackStyle"
End Property

Private Sub UserControl_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
  Debug.Print "Dragged over"
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_CtrlEffect = m_def_CtrlEffect
  m_AutoSize = m_def_AutoSize
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
  DrawControl
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If tmrMouseEnter.Enabled = False Then
    RaiseEvent MouseEnter
    tmrMouseEnter.Enabled = True
  End If
  RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
  If UserControl.Height < 120 Then UserControl.Height = 120
  If UserControl.Width < 120 Then UserControl.Width = 120
  DrawControl
End Sub

Private Sub UserControl_Paint()
  DrawControl True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
  UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
  m_CtrlEffect = PropBag.ReadProperty("CtrlEffect", m_def_CtrlEffect)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
  UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
  Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
  Call PropBag.WriteProperty("CtrlEffect", m_CtrlEffect, m_def_CtrlEffect)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
  Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
End Sub

