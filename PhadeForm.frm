VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form PhadeForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phader Example by Chad"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "PhadeForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDynamic 
      Caption         =   "Dynamic Phade (for very fast PCs only)."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   " Phades the text as you scroll the color sliders above. "
      Top             =   2385
      Width           =   3500
   End
   Begin VB.CheckBox chkStrike 
      Caption         =   "Strike"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2580
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox chkUnderline 
      Caption         =   "Underline"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1485
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   750
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtHTML 
      Height          =   735
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   2670
      Width           =   3225
   End
   Begin RichTextLib.RichTextBox txtPreview 
      Height          =   300
      Left            =   60
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1815
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   529
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"PhadeForm.frx":27A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   21
      Top             =   1530
      Width           =   3225
   End
   Begin VB.HScrollBar Scroll9 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1320
      Value           =   255
      Width           =   2500
   End
   Begin VB.HScrollBar Scroll8 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1185
      Width           =   2500
   End
   Begin VB.HScrollBar Scroll7 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1050
      Width           =   2500
   End
   Begin VB.HScrollBar Scroll6 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   840
      Width           =   2500
   End
   Begin VB.HScrollBar Scroll5 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   705
      Value           =   255
      Width           =   2500
   End
   Begin VB.HScrollBar Scroll4 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   570
      Width           =   2500
   End
   Begin VB.HScrollBar Scroll3 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   2500
   End
   Begin VB.HScrollBar Scroll2 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   225
      Width           =   2500
   End
   Begin VB.HScrollBar Scroll1 
      Height          =   125
      LargeChange     =   25
      Left            =   540
      Max             =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Value           =   255
      Width           =   2500
   End
   Begin VB.Label B3 
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3075
      TabIndex        =   20
      Top             =   1305
      Width           =   495
   End
   Begin VB.Label G3 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3075
      TabIndex        =   19
      Top             =   1170
      Width           =   495
   End
   Begin VB.Label R3 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3075
      TabIndex        =   18
      Top             =   1035
      Width           =   495
   End
   Begin VB.Label B2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3075
      TabIndex        =   17
      Top             =   825
      Width           =   495
   End
   Begin VB.Label G2 
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3075
      TabIndex        =   16
      Top             =   690
      Width           =   495
   End
   Begin VB.Label R2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   125
      Left            =   3075
      TabIndex        =   15
      Top             =   555
      Width           =   500
   End
   Begin VB.Label Color3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   75
      TabIndex        =   8
      Top             =   1050
      Width           =   405
   End
   Begin VB.Label Color2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   75
      TabIndex        =   7
      Top             =   570
      Width           =   405
   End
   Begin VB.Label B1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3075
      TabIndex        =   6
      Top             =   345
      Width           =   495
   End
   Begin VB.Label G1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3075
      TabIndex        =   5
      Top             =   210
      Width           =   495
   End
   Begin VB.Label R1 
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   125
      Left            =   3075
      TabIndex        =   4
      Top             =   75
      Width           =   500
   End
   Begin VB.Label Color1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   405
   End
End
Attribute VB_Name = "PhadeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub FadeIt()
Dim strTags As String
If chkBold.Value = Checked Then strTags$ = strTags$ & "<b>"
If chkItalic.Value = Checked Then strTags$ = strTags$ & "<i>"
If chkUnderline.Value = Checked Then strTags$ = strTags$ & "<u>"
If chkStrike.Value = Checked Then strTags$ = strTags$ & "<s>"
txtHTML.Text = strTags$ & FadeByColor3(Color1.BackColor, Color2.BackColor, Color3.BackColor, txtText, False)
Call FadePreview2(txtPreview, txtHTML)
End Sub

Private Sub chkBold_Click()
 If chkBold.Value = 1 Then
  txtHTML.Text = "<b>" & txtHTML
  Call FadePreview2(txtPreview, txtHTML)
 Else
  txtHTML.Text = Replace(txtHTML, "<b>", "")
  Call FadePreview2(txtPreview, txtHTML)
 End If
End Sub

Private Sub chkBold_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FadeIt
End Sub

Private Sub chkItalic_Click()
 If chkItalic.Value = 1 Then
  txtHTML.Text = "<i>" & txtHTML
  Call FadePreview2(txtPreview, txtHTML)
 End If
End Sub

Private Sub chkStrike_Click()
 If chkStrike.Value = 1 Then
  txtHTML.Text = "<s>" & txtHTML
  Call FadePreview2(txtPreview, txtHTML)
 End If
End Sub

Private Sub chkUnderline_Click()
 If chkUnderline.Value = 1 Then
  txtHTML.Text = "<u>" & txtHTML
  Call FadePreview2(txtPreview, txtHTML)
 End If
End Sub

Private Sub Scroll1_Change()
R1.Caption = Scroll1.Value
Color1.BackColor = RGB2Long(Scroll1.Value, Scroll2.Value, Scroll3.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll1_Scroll()
R1.Caption = Scroll1.Value
Color1.BackColor = RGB2Long(Scroll1.Value, Scroll2.Value, Scroll3.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll2_Change()
G1.Caption = Scroll2.Value
Color1.BackColor = RGB2Long(Scroll1.Value, Scroll2.Value, Scroll3.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll2_Scroll()
G1.Caption = Scroll2.Value
Color1.BackColor = RGB2Long(Scroll1.Value, Scroll2.Value, Scroll3.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll3_Change()
B1.Caption = Scroll3.Value
Color1.BackColor = RGB2Long(Scroll1.Value, Scroll2.Value, Scroll3.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll3_Scroll()
B1.Caption = Scroll3.Value
Color1.BackColor = RGB2Long(Scroll1.Value, Scroll2.Value, Scroll3.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll4_Change()
R2.Caption = Scroll4.Value
Color2.BackColor = RGB2Long(Scroll4.Value, Scroll5.Value, Scroll6.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll4_Scroll()
R2.Caption = Scroll4.Value
Color2.BackColor = RGB2Long(Scroll4.Value, Scroll5.Value, Scroll6.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll5_Change()
G2.Caption = Scroll5.Value
Color2.BackColor = RGB2Long(Scroll4.Value, Scroll5.Value, Scroll6.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll5_Scroll()
G2.Caption = Scroll5.Value
Color2.BackColor = RGB2Long(Scroll4.Value, Scroll5.Value, Scroll6.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll6_Change()
B2.Caption = Scroll6.Value
Color2.BackColor = RGB2Long(Scroll4.Value, Scroll5.Value, Scroll6.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll6_Scroll()
B2.Caption = Scroll6.Value
Color2.BackColor = RGB2Long(Scroll4.Value, Scroll5.Value, Scroll6.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll7_Change()
R3.Caption = Scroll7.Value
Color3.BackColor = RGB2Long(Scroll7.Value, Scroll8.Value, Scroll9.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll7_Scroll()
R3.Caption = Scroll7.Value
Color3.BackColor = RGB2Long(Scroll7.Value, Scroll8.Value, Scroll9.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll8_Change()
G3.Caption = Scroll8.Value
Color3.BackColor = RGB2Long(Scroll7.Value, Scroll8.Value, Scroll9.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll8_Scroll()
G3.Caption = Scroll8.Value
Color3.BackColor = RGB2Long(Scroll7.Value, Scroll8.Value, Scroll9.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll9_Change()
B3.Caption = Scroll9.Value
Color3.BackColor = RGB2Long(Scroll7.Value, Scroll8.Value, Scroll9.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub Scroll9_Scroll()
B3.Caption = Scroll9.Value
Color3.BackColor = RGB2Long(Scroll7.Value, Scroll8.Value, Scroll9.Value)
If chkDynamic.Value = 1 Then Call FadeIt
End Sub

Private Sub txtText_Change()
Call FadeIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
