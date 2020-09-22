VERSION 5.00
Object = "*\A..\..\..\..\PROGRA~2\MyVB\VBCODE~1\COLLAP~1\CollapsableFrame.vbp"
Begin VB.Form Form1 
   Caption         =   "Collapsable Frame Demo"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check5 
      Caption         =   "ShowFocusRectangle"
      Height          =   345
      Left            =   5370
      TabIndex        =   7
      Top             =   3000
      Value           =   1  'Checked
      Width           =   2115
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Enabled"
      Height          =   345
      Left            =   165
      TabIndex        =   6
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.TextBox txtAnimationSpeed 
      Height          =   330
      Left            =   2850
      TabIndex        =   3
      Top             =   3435
      Width           =   1020
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Animate            Animation Speed:"
      Height          =   345
      Left            =   165
      TabIndex        =   2
      Top             =   3435
      Value           =   1  'Checked
      Width           =   2730
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ShowArrow"
      Height          =   345
      Left            =   3635
      TabIndex        =   1
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Caption         =   "RightToLeft"
      Height          =   345
      Left            =   1900
      TabIndex        =   0
      Top             =   3000
      Width           =   1275
   End
   Begin CollaspableFrame.CollpasableFrame CollpasableFrame1 
      Align           =   1  'Align Top
      Height          =   2835
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   5001
      ExpandedHeight  =   2835
      CaptionForeColor=   -2147483630
      CaptionBackColor=   12632319
      Caption         =   "&Frame's caption goes here, duh!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BodyBackColor   =   14737632
      BorderColor     =   0
      CollapsedCaptionTooltipText=   "ÇÖÛØ åäÇ áÊÕÛíÑ ÇáÅØÇÑ"
      Begin VB.CommandButton Command1 
         Caption         =   "&Expand/Collapse"
         Height          =   420
         Left            =   1027
         TabIndex        =   11
         Top             =   2055
         Width           =   2040
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   1365
         TabIndex        =   10
         Top             =   705
         Width           =   2505
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Expand/Collapse Without Animation"
         Height          =   420
         Left            =   3442
         TabIndex        =   9
         Top             =   2055
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Some &Label:"
         Height          =   285
         Left            =   330
         TabIndex        =   12
         Top             =   735
         Width           =   1020
      End
   End
   Begin VB.Label Label2 
      Caption         =   "(10-150)"
      Height          =   225
      Left            =   3990
      TabIndex        =   4
      Top             =   3480
      Width           =   690
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      Height          =   480
      Left            =   195
      TabIndex        =   5
      Top             =   3930
      Width           =   7305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    CollpasableFrame1.RightToLeft = Check1.Value
End Sub

Private Sub Check2_Click()
   CollpasableFrame1.ShowArrow = Check2.Value
End Sub

Private Sub Check3_Click()
    CollpasableFrame1.Animate = Check3.Value
    txtAnimationSpeed.Enabled = Check3.Value
    Label2.Enabled = Check3.Value
End Sub


Private Sub CollpasableFrame1_Collapsed()
    FlexUI1.UpdateControlPositionSize CollpasableFrame1
End Sub

Private Sub Command2_Click()
    If CollpasableFrame1.Expanded Then
        CollpasableFrame1.CollapseWithoutAnimation
    Else
        CollpasableFrame1.ExpandWithoutAnimation
    End If
End Sub

Private Sub Check4_Click()
    CollpasableFrame1.Enabled = Check4.Value
End Sub

Private Sub Check5_Click()
    CollpasableFrame1.ShowFocusRectangle = Check5.Value
End Sub

Private Sub Command1_Click()
    CollpasableFrame1.Expanded = Not CollpasableFrame1.Expanded
End Sub

Private Sub Form_Load()
    FlexUI1.InitResizer
    txtAnimationSpeed = CollpasableFrame1.AnimationSpeed
    Check4.Value = IIf(CollpasableFrame1.Enabled, 1, 0)
    Check1.Value = IIf(CollpasableFrame1.RightToLeft, 1, 0)
    Check2.Value = IIf(CollpasableFrame1.ShowArrow, 1, 0)
    Check5.Value = IIf(CollpasableFrame1.ShowFocusRectangle, 1, 0)
    Check3.Value = IIf(CollpasableFrame1.Animate, 1, 0)
End Sub

Private Sub Form_Resize()
    FlexUI1.Resize
End Sub

Private Sub txtAnimationSpeed_LostFocus()
    CollpasableFrame1.AnimationSpeed = Val(txtAnimationSpeed.Text)
End Sub
