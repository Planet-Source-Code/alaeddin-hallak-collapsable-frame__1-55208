VERSION 5.00
Begin VB.UserControl CollpasableFrame 
   Alignable       =   -1  'True
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2670
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   2670
   ToolboxBitmap   =   "CollapsableFrame.ctx":0000
   Begin VB.Image imgCollapse 
      Height          =   240
      Left            =   1680
      Picture         =   "CollapsableFrame.ctx":0312
      Top             =   1350
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgExpand 
      Height          =   240
      Left            =   735
      Picture         =   "CollapsableFrame.ctx":069C
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Overlay 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      MouseIcon       =   "CollapsableFrame.ctx":0A26
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Click to collapse the frame"
      Top             =   0
      Width           =   2670
   End
   Begin VB.Image imgExpandCollpase 
      Height          =   240
      Left            =   2340
      ToolTipText     =   "Expand section"
      Top             =   50
      Width           =   240
   End
   Begin VB.Label lblFrameCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      MousePointer    =   4  'Icon
      TabIndex        =   0
      Top             =   60
      Width           =   2220
   End
   Begin VB.Shape shpCaptionBack 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   2670
   End
   Begin VB.Shape shpBodyFrame 
      BorderColor     =   &H80000010&
      Height          =   2340
      Left            =   0
      Top             =   255
      Width           =   2625
   End
End
Attribute VB_Name = "CollpasableFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Animating(ByVal bExpand As Boolean)
Public Event Expanded()
Public Event Collapsed()

Public Enum BorderStyleConstants
    [None] = 0
    [Fixed Single] = 1
    [Fixed 3D] = 2
End Enum

Private m_bAllowStateChange As Boolean
Private m_BorderStyle As BorderStyleConstants
Private m_bExpanded As Boolean
Private m_sngExpandedHeight As Single
Private m_bRightToLeft As Boolean
Private m_bShowSelRectangle As Boolean
Private m_bAnimate As Boolean
Private m_bEnabled As Boolean
Private m_intAnimationSpd As Integer    'Min 10, Max 200
Private m_strExpCaptionTooltip As String
Private m_strCollCaptionTooltip As String

Private m_bAnimating As Boolean         'Set to true in order to surpress UserControl resize events while animating

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "Height" Then
    
    End If
End Sub

Private Sub UserControl_GotFocus()
    If m_bShowSelRectangle Then
        shpCaptionBack.BorderStyle = 3
    End If
End Sub

Private Sub UserControl_LostFocus()
    If m_bShowSelRectangle Then
        If m_BorderStyle = [Fixed 3D] Or m_BorderStyle = None Then
            shpCaptionBack.BorderStyle = 0
        Else
            shpCaptionBack.BorderStyle = 1
        End If
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If shpCaptionBack.BorderStyle = 3 Then
        If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
            Expanded = Not Expanded
        ElseIf KeyCode = vbKeyUp And Expanded = True Then
            Expanded = False
        ElseIf KeyCode = vbKeyDown And Expanded = False Then
            Expanded = True
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    ExpandedHeight = shpBodyFrame.Height
    imgExpandCollpase.Picture = imgCollapse.Picture
End Sub

Private Sub UserControl_InitProperties()
    ShowFocusRectangle = True
    AllowStateChange = True
    Enabled = True
    Expanded = True
    ExpandedHeight = 2340
    RightToLeft = Ambient.RightToLeft
    ShowArrow = True
    Animate = True
    AnimationSpeed = 90
    BorderStyle = [Fixed Single]
    Font = Ambient.Font
    CaptionForeColor = Ambient.ForeColor
    ExpandedCaptionTooltipText = "Click to collapse the frame"
    CollapsedCaptionTooltipText = "Click to expand the frame"
    CaptionMousePointer = vbCustom
       
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    DoEvents
    With PropBag
        .WriteProperty "AllowStateChange", m_bAllowStateChange, True
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "Expanded", m_bExpanded, True
        .WriteProperty "ExpandedHeight", m_sngExpandedHeight, 2340
        .WriteProperty "RightToLeft", m_bRightToLeft, False
        .WriteProperty "ShowFocusRectangle", m_bShowSelRectangle, True
        .WriteProperty "CaptionForeColor", lblFrameCaption.ForeColor, vbBlack
        .WriteProperty "CaptionBackColor", shpCaptionBack.BackColor, vbWhite
        .WriteProperty "Caption", lblFrameCaption.Caption, "Caption Text"
        .WriteProperty "Font", Font, Ambient.Font
        .WriteProperty "BodyBackColor", UserControl.BackColor, vbButtonFace
        .WriteProperty "BorderColor", shpBodyFrame.BorderColor, vbButtonShadow
        .WriteProperty "BorderStyle", m_BorderStyle, BorderStyleConstants.[Fixed Single]
        .WriteProperty "BorderWidth", shpBodyFrame.BorderWidth, 1
        .WriteProperty "ShowArrow", imgExpandCollpase.Visible, True
        .WriteProperty "Animate", m_bAnimate, True
        .WriteProperty "AnimationSpeed", m_intAnimationSpd, 90
        .WriteProperty "ExpandedCaptionTooltipText", m_strExpCaptionTooltip, "Click to collapse the frame"
        .WriteProperty "CollapsedCaptionTooltipText", m_strCollCaptionTooltip, "Click to expand the frame"
        .WriteProperty "CaptionMousePointer", Overlay.MousePointer, MousePointerConstants.vbCustom
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    DoEvents
    With PropBag
        AllowStateChange = .ReadProperty("AllowStateChange", True)
        Enabled = .ReadProperty("Enabled", True)
        Expanded = .ReadProperty("Expanded", True)
        ExpandedHeight = .ReadProperty("ExpandedHeight", 2340)
        ShowFocusRectangle = .ReadProperty("ShowFocusRectangle", True)
        RightToLeft = .ReadProperty("RightToLeft", False)
        CaptionForeColor = .ReadProperty("CaptionForeColor", vbBlack)
        CaptionBackColor = .ReadProperty("CaptionBackColor", vbWhite)
        BodyBackColor = .ReadProperty("BodyBackColor", vbButtonFace)
        BorderColor = .ReadProperty("BorderColor", vbButtonShadow)
        BorderStyle = .ReadProperty("BorderStyle", BorderStyleConstants.[Fixed Single])
        BorderWidth = .ReadProperty("BorderWidth", 1)
        Caption = .ReadProperty("Caption", "Caption Text")
        Set Font = .ReadProperty("Font", Font)
        ShowArrow = .ReadProperty("ShowArrow", True)
        Animate = .ReadProperty("Animate", True)
        AnimationSpeed = .ReadProperty("AnimationSpeed", 90)
        ExpandedCaptionTooltipText = .ReadProperty("ExpandedCaptionTooltipText", "Click to collapse the frame")
        CollapsedCaptionTooltipText = .ReadProperty("CollapsedCaptionTooltipText", "Click to expand the frame")
        CaptionMousePointer = .ReadProperty("CaptionMousePointer", MousePointerConstants.vbCustom)
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If Not m_bAnimating Then
        shpCaptionBack.Move 0, 0, UserControl.Width
        
        imgExpandCollpase.Left = IIf(m_bRightToLeft, 90, shpCaptionBack.Width - 330)
        lblFrameCaption.Width = shpCaptionBack.Width - 450
        
        shpBodyFrame.Move 0, shpCaptionBack.Height - 10, UserControl.Width, UserControl.Height - shpCaptionBack.Height
        Overlay.Move 0, 0, UserControl.Width
        
        If UserControl.Height < shpCaptionBack.Height Then UserControl.Height = shpCaptionBack.Height
        
        If m_bExpanded Then
            m_sngExpandedHeight = UserControl.Height
        Else
            UserControl.Height = shpCaptionBack.Height
        End If
    End If
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets whether contained controls are enabled or disabled"
Attribute Enabled.VB_UserMemId = -514
    Enabled = m_bEnabled
End Property

Public Property Get ShowFocusRectangle() As Boolean
Attribute ShowFocusRectangle.VB_Description = "Returns/sets whether a focus rectangle is displayed around the frame's caption when the frame receives focus"
    ShowFocusRectangle = m_bShowSelRectangle
End Property

Public Property Let ShowFocusRectangle(ByVal vNewValue As Boolean)
    m_bShowSelRectangle = vNewValue
End Property

Public Property Get CaptionMousePointer() As MousePointerConstants
   CaptionMousePointer = Overlay.MousePointer
End Property

Public Property Let CaptionMousePointer(ByVal NewPointer As MousePointerConstants)
   Overlay.MousePointer = NewPointer
   PropertyChanged "CaptionMousePointer"
End Property

Public Property Get AllowStateChange() As Boolean
Attribute AllowStateChange.VB_Description = "Returns/sets whether user can change the state of frame"
    AllowStateChange = m_bAllowStateChange
End Property

Public Property Let AllowStateChange(ByVal vNewValue As Boolean)
    If vNewValue = False Then
        ShowArrow = False
        PropertyChanged "ShowArrow"
    End If
    m_bAllowStateChange = vNewValue
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    m_bEnabled = vNewValue
    Dim x As Control
    For Each x In UserControl.ContainedControls
        x.Enabled = m_bEnabled
    Next
End Property

Public Property Let Expanded(Value As Boolean)
Attribute Expanded.VB_Description = "Returns/sets whether the frame is in expanded or in collapsed mode"
    If ChangeState(Value) Then
        m_bExpanded = Value
    End If
End Property

Public Property Get Expanded() As Boolean
    Expanded = m_bExpanded
End Property

Public Property Get ExpandedHeight() As Single
Attribute ExpandedHeight.VB_Description = "Returns/sets the full height of the frame in the expanded state"
    ExpandedHeight = m_sngExpandedHeight
End Property

Public Property Let ExpandedHeight(ByVal vNewValue As Single)
    m_sngExpandedHeight = vNewValue
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed on frame caption"
Attribute Caption.VB_UserMemId = -518
    Caption = lblFrameCaption.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    lblFrameCaption.Caption = vNewValue
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Returns/sets the right to left orientation"
Attribute RightToLeft.VB_UserMemId = -611
    RightToLeft = m_bRightToLeft
End Property

Public Property Let RightToLeft(ByVal vNewValue As Boolean)
    If vNewValue <> m_bRightToLeft Then
        m_bRightToLeft = vNewValue
        FlipControlsHorizontal
    End If
End Property

Public Property Get ExpandedCaptionTooltipText() As String
    ExpandedCaptionTooltipText = m_strExpCaptionTooltip
End Property

Public Property Let ExpandedCaptionTooltipText(ByVal vNewValue As String)
    m_strExpCaptionTooltip = vNewValue
End Property

Public Property Get CollapsedCaptionTooltipText() As String
    CollapsedCaptionTooltipText = m_strCollCaptionTooltip
End Property

Public Property Let CollapsedCaptionTooltipText(ByVal vNewValue As String)
    m_strCollCaptionTooltip = vNewValue
End Property

Public Property Get CaptionBackColor() As OLE_COLOR
Attribute CaptionBackColor.VB_Description = "Returns/sets the back color of the frame header (the caption)"
    CaptionBackColor = shpCaptionBack.BackColor
End Property

Public Property Let CaptionBackColor(ByVal vNewValue As OLE_COLOR)
    shpCaptionBack.BackColor = vNewValue
    lblFrameCaption.BackColor = vNewValue
End Property

Public Property Get CaptionForeColor() As OLE_COLOR
Attribute CaptionForeColor.VB_Description = "Returns/sets the forecolor of the frame header (the caption)"
Attribute CaptionForeColor.VB_UserMemId = -513
    CaptionForeColor = lblFrameCaption.ForeColor
End Property

Public Property Let CaptionForeColor(ByVal vNewValue As OLE_COLOR)
    lblFrameCaption.ForeColor = vNewValue
End Property

Public Property Get BodyBackColor() As OLE_COLOR
Attribute BodyBackColor.VB_Description = "Returns/sets the back color of the frame body"
    BodyBackColor = UserControl.BackColor
End Property

Public Property Let BodyBackColor(ByVal vNewValue As OLE_COLOR)
    UserControl.BackColor = vNewValue
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the border color of the control"
Attribute BorderColor.VB_UserMemId = -503
    BorderColor = shpBodyFrame.BorderColor
End Property

Public Property Let BorderColor(ByVal vNewValue As OLE_COLOR)
    shpBodyFrame.BorderColor = vNewValue
    shpCaptionBack.BorderColor = vNewValue
End Property

Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for the frame"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewValue As BorderStyleConstants)
    If vNewValue = None Then
        shpCaptionBack.BorderStyle = 0
        shpBodyFrame.BorderStyle = 0
        UserControl.BorderStyle = 0
    ElseIf vNewValue = [Fixed Single] Then
        shpCaptionBack.BorderStyle = 1
        shpBodyFrame.BorderStyle = 1
        UserControl.BorderStyle = 0
    Else
        shpCaptionBack.BorderStyle = 0
        shpBodyFrame.BorderStyle = 0
        UserControl.BorderStyle = 1
    End If
    m_BorderStyle = vNewValue
End Property

Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns/sets the width of the frame's border"
Attribute BorderWidth.VB_UserMemId = -505
    BorderWidth = shpBodyFrame.BorderWidth
End Property

Public Property Let BorderWidth(ByVal vNewValue As Integer)
    On Error GoTo ErrHandler
    shpBodyFrame.BorderWidth = vNewValue
    shpCaptionBack.BorderWidth = vNewValue
    
ErrHandler:
    
End Property

Public Property Get ShowArrow() As Boolean
Attribute ShowArrow.VB_Description = "Returns/sets whether the arrow at the top of the frame is displayed"

    ShowArrow = imgExpandCollpase.Visible
    
End Property

Public Property Let ShowArrow(ByVal vNewValue As Boolean)
  
    imgExpandCollpase.Visible = vNewValue And m_bAllowStateChange
    
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns sets the font of the frame"
Attribute Font.VB_UserMemId = -512
    Set Font = lblFrameCaption.Font
End Property

Public Property Set Font(vNewValue As StdFont)
    Set lblFrameCaption.Font = vNewValue
    PropertyChanged "Font"
End Property

Public Property Get Animate() As Boolean
Attribute Animate.VB_Description = "Returns/sets whether expanding or collapsing the frame is done with or without animation"
    Animate = m_bAnimate
End Property

Public Property Let Animate(ByVal vNewValue As Boolean)
    m_bAnimate = vNewValue
End Property

Public Property Get AnimationSpeed() As Integer
Attribute AnimationSpeed.VB_Description = "Returns/sets the speed of animation. (Min: 10, Max: 200)"
    AnimationSpeed = m_intAnimationSpd
End Property

Public Property Let AnimationSpeed(ByVal vNewValue As Integer)
    If vNewValue > 200 Then
        m_intAnimationSpd = 200
    ElseIf vNewValue >= 10 And vNewValue <= 200 Then
        m_intAnimationSpd = vNewValue
    End If
End Property

Private Sub FlipControlsHorizontal()

    Dim ctrl As Control
    On Error Resume Next
    
    UserControl.RightToLeft = Not UserControl.RightToLeft
    
    For Each ctrl In UserControl.Controls
        ctrl.Left = UserControl.Width - ctrl.Left - ctrl.Width
        ctrl.RightToLeft = UserControl.RightToLeft
        If UserControl.RightToLeft Then
            ctrl.Alignment = 1
        Else
            ctrl.Alignment = 0
        End If
    Next ctrl
    
    For Each ctrl In UserControl.ContainedControls
        ctrl.Left = ctrl.Container.Width - ctrl.Left - ctrl.Width
        ctrl.RightToLeft = UserControl.RightToLeft
        If UserControl.RightToLeft Then
            ctrl.Alignment = 1
        Else
            ctrl.Alignment = 0
        End If
    Next ctrl
    
End Sub

Private Sub Overlay_Click()
    If ChangeState(Not m_bExpanded) Then
        m_bExpanded = Not m_bExpanded
        PropertyChanged "Expanded"
    End If
End Sub

Public Sub Expand()
    If ChangeState(True) Then
        m_bExpanded = True
        PropertyChanged "Expanded"
    End If
End Sub

Public Sub ExpandWithoutAnimation()
    Dim tmp As Boolean
    tmp = m_bAnimate
    m_bAnimate = False
    
    If ChangeState(True) Then
        m_bExpanded = True
        PropertyChanged "Expanded"
    End If
    
    m_bAnimate = tmp
End Sub

Public Sub Collapse()
    If ChangeState(False) Then
        m_bExpanded = False
        PropertyChanged "Expanded"
    End If
End Sub

Public Sub CollapseWithoutAnimation()
    Dim tmp As Boolean
    tmp = m_bAnimate
    m_bAnimate = False
    
    If ChangeState(False) Then
        m_bExpanded = False
        PropertyChanged "Expanded"
    End If
    
    m_bAnimate = tmp
End Sub

Private Function ChangeState(ByVal Expand As Boolean) As Boolean
    If Not m_bAnimating And (m_bAllowStateChange Or (Not m_bAllowStateChange And Not Ambient.UserMode)) Then
        m_bAnimating = True
        If Expand Then
            
            If m_bAnimate Then
                AnimateExpand
            Else
                UserControl.Height = m_sngExpandedHeight
            End If
            
            imgExpandCollpase.Picture = imgCollapse.Picture
            Overlay.ToolTipText = m_strExpCaptionTooltip
            shpBodyFrame.Move 0, shpCaptionBack.Height - 10, UserControl.Width, UserControl.Height - shpCaptionBack.Height
            
            RaiseEvent Expanded
        Else
            ExpandedHeight = UserControl.Height
            
            If m_bAnimate Then
                AnimateCollapse
            Else
                UserControl.Height = shpCaptionBack.Height
            End If
            
            imgExpandCollpase.Picture = imgExpand.Picture
            Overlay.ToolTipText = m_strCollCaptionTooltip
            
            RaiseEvent Collapsed
        End If
        ChangeState = True
        m_bAnimating = False
    Else
        ChangeState = False
    End If
End Function

Private Sub AnimateExpand()
   
    Do Until UserControl.Height >= m_sngExpandedHeight
        UserControl.Height = UserControl.Height + m_intAnimationSpd
        UserControl.Refresh
        DoEvents
        RaiseEvent Animating(True)
    Loop
    UserControl.Parent.Refresh
    UserControl.Height = m_sngExpandedHeight
    
End Sub

Private Sub AnimateCollapse()
    
    On Error GoTo ErrHandler
    
    Do Until UserControl.Height <= shpCaptionBack.Height
        UserControl.Height = UserControl.Height - m_intAnimationSpd
        UserControl.Refresh
        DoEvents
        RaiseEvent Animating(False)
    Loop
    
ErrHandler:
    UserControl.Parent.Refresh
    If UserControl.Height < shpCaptionBack.Height Then UserControl.Height = shpCaptionBack.Height
End Sub

Public Property Get Controls() As ContainedControls
    Set Controls = UserControl.ContainedControls
End Property

Public Property Get Height() As Integer
    Height = UserControl.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    If m_bExpanded Then
        UserControl.Height = vNewValue
        PropertyChanged "Height"
    End If
End Property
