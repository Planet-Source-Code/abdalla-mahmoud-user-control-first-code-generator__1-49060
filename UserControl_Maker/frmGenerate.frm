VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmGenerate 
   Caption         =   "User Control Code"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   7170
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox Text 
      Height          =   2415
      Left            =   1080
      TabIndex        =   0
      Top             =   2400
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4260
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   50000
      TextRTF         =   $"frmGenerate.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmGenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Call Text.Move(0, 0, ScaleWidth, ScaleHeight)
End Sub
