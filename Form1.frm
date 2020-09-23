VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Scroll Through a Form"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   270
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   1320
      Width           =   2925
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1995
      Left            =   4410
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   810
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   2850
      Width           =   2955
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   840
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2280
      Width           =   2595
   End
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   780
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1770
      Width           =   2595
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   690
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   990
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   540
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   810
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   150
      Width           =   2355
   End
   Begin VB.Menu mnuVE 
      Caption         =   "Vscroll Equals"
   End
   Begin VB.Menu mnuSize 
      Caption         =   "Size Of Form"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
VScroll.Top = 0
VScroll.Height = Form1.Height - 400
VScroll.TabStop = False         'This is to stop that annoying blinking
VScroll.Min = 0
VScroll.Max = 100
VScroll.LargeChange = 3
VScroll.SmallChange = 1
VScroll.Value = VScroll.Min
Text1.Left = 210
Text2.Left = 210
Text3.Left = 210
Text4.Left = 210
Text5.Left = 210
Text6.Left = 210
Text7.Left = 210
Text1.Height = 315
Text2.Height = 315
Text3.Height = 315
Text4.Height = 315
Text5.Height = 315
Text6.Height = 315
Text7.Height = 315
Text1.Width = Form1.Width - 750
Text2.Width = Form1.Width - 750
Text3.Width = Form1.Width - 750
Text4.Width = Form1.Width - 750
Text5.Width = Form1.Width - 750
Text6.Width = Form1.Width - 750
Text7.Width = Form1.Width - 750
Call VScroll_Change             'This makes sure the boxes are organized correctly
End Sub

Private Sub Form_Resize()
VScroll.Top = 0
VScroll.Left = Form1.Width - 375
'VScroll.Height = Form1.Height - 625
VScroll.Height = Form1.Height - 910
Text1.Width = Form1.Width - 750
Text2.Width = Form1.Width - 750
Text3.Width = Form1.Width - 750
Text4.Width = Form1.Width - 750
Text5.Width = Form1.Width - 750
Text6.Width = Form1.Width - 750
Text7.Width = Form1.Width - 750
End Sub

Private Sub mnuSize_Click()
MsgBox "Height = " & Form1.Height & " Width = " & Form1.Width
End Sub

Private Sub mnuVE_Click()
MsgBox "Vscroll = " & VScroll.Value
End Sub

Private Sub Text1_GotFocus()
If (Text1.Top < (VScroll.Value - 100) + 200) Or _
    (Text1.Top > (VScroll.Value - 100) + (Form1.Height - 1200)) Then
    VScroll.Value = 0
End If
End Sub
Private Sub Text2_GotFocus()
If (Text2.Top < (VScroll.Value - 100) + 200) Or _
    (Text2.Top > (VScroll.Value - 100) + (Form1.Height - 1200)) Then
    VScroll.Value = 5
End If
End Sub

Private Sub Text3_GotFocus()
If (Text3.Top < (VScroll.Value - 100) + 200) Or _
    (Text3.Top > (VScroll.Value - 100) + (Form1.Height - 1200)) Then
    VScroll.Value = 9
End If
End Sub

Private Sub Text4_GotFocus()
If (Text4.Top < (VScroll.Value - 100) + 200) Or _
    (Text4.Top > (VScroll.Value - 100) + (Form1.Height - 1200)) Then
    VScroll.Value = 13
End If
End Sub

Private Sub Text5_GotFocus()
If (Text5.Top < (VScroll.Value - 100) + 200) Or _
    (Text5.Top > (VScroll.Value - 100) + (Form1.Height - 1200)) Then
    VScroll.Value = 17
End If
End Sub

Private Sub Text6_GotFocus()
If (Text6.Top < (VScroll.Value - 100) + 200) Or _
    (Text6.Top > (VScroll.Value - 100) + (Form1.Height - 1200)) Then
    VScroll.Value = 21
End If
End Sub

Private Sub Text7_GotFocus()
If (Text7.Top < (VScroll.Value - 100) + 200) Or _
    (Text7.Top > (VScroll.Value - 100) + (Form1.Height - 1200)) Then
    VScroll.Value = 25
End If
End Sub

Private Sub VScroll_Change()
Text1.Top = (VScroll.Value * -100) + 200
Text2.Top = (VScroll.Value * -100) + 600
Text3.Top = (VScroll.Value * -100) + 1000
Text4.Top = (VScroll.Value * -100) + 1400
Text5.Top = (VScroll.Value * -100) + 1800
Text6.Top = (VScroll.Value * -100) + 2200
Text7.Top = (VScroll.Value * -100) + 2600
End Sub

