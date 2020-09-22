VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Control First Code Generator"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDefValue 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   29
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   645
      Left            =   4080
      TabIndex        =   28
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox txtDType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   27
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   345
      Left            =   5040
      TabIndex        =   26
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   25
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtPara 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   23
      Top             =   5520
      Width           =   2415
   End
   Begin VB.ListBox lstDType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      ItemData        =   "frmMain.frx":0000
      Left            =   3960
      List            =   "frmMain.frx":002E
      TabIndex        =   22
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Frame framProps 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   1320
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
      Begin VB.OptionButton optProp3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Friend ."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optProp3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Private ."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optProp3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Public ."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame framProps 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
      Begin VB.OptionButton optProp2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Friend ."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optProp2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Private ."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optProp2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Public ."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame framProps 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
      Begin VB.OptionButton optProp1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Friend ."
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optProp1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Private ."
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optProp1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Public ."
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkProps 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Set ."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox chkProps 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Let ."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CheckBox chkProps 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Get ."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox cboProps 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "PropNameHere"
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Default Value :"
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   30
      Top             =   5880
      Width           =   1050
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Paramaters :"
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   24
      Top             =   5520
      Width           =   885
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   120
      Top             =   5400
      Width           =   5895
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data type :"
      Height          =   195
      Index           =   4
      Left            =   4080
      TabIndex        =   21
      Top             =   1440
      Width           =   780
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   3960
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   3960
      Y1              =   1680
      Y2              =   5400
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scope :"
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   20
      Top             =   1440
      Width           =   555
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   1200
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   19
      Top             =   1440
      Width           =   450
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   3960
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3960
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   1200
      Y1              =   1320
      Y2              =   5400
   End
   Begin VB.Shape Shape1 
      Height          =   4095
      Left            =   120
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Properties :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1020
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   4920
      Index           =   1
      Left            =   120
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   510
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMAin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//Name : UserControl Maker
'//Author : Abdalla Mahmoud
'//E-Mail : la3toot@hotmail.com
'         : la3toot@yahoo.com
'//Purpose : This Program will generate code
'of UserControls Automaticly , Just Enter props
'and other options the Click Generate Enjoy ;)
Option Explicit
Private Type typProperty
    lName     As String
    lGet      As String
    lLet      As String
    lSet      As String
    lData     As String
    lPara     As String
    lDefValue As String
End Type
Private m_Items() As typProperty
Private m_Count   As Long
Private m_CanEdit As Boolean

Private Sub cboProps_Click()
    If cboProps.ListIndex = -1 Then Exit Sub
    m_CanEdit = True
    With m_Items(cboProps.ListIndex + 1)
        chkProps(0).Value = -(.lGet <> vbNullString)
            optProp1(0).Value = -(.lGet = "P")
            optProp1(1).Value = -(.lGet = "V")
            optProp1(2).Value = -(.lGet = "F")
        chkProps(1).Value = -(.lLet <> vbNullString)
            optProp2(0).Value = -(.lLet = "P")
            optProp2(1).Value = -(.lLet = "V")
            optProp2(2).Value = -(.lLet = "F")
        chkProps(2).Value = -(.lSet <> vbNullString)
            optProp3(0).Value = -(.lSet = "P")
            optProp3(1).Value = -(.lSet = "V")
            optProp3(2).Value = -(.lSet = "F")
        txtDType.Text = .lData
        txtPara.Text = .lPara
        txtDefValue.Text = .lDefValue
    End With
    m_CanEdit = False
End Sub

Private Sub chkProps_Click(Index As Integer)
    If cboProps.ListIndex = -1 Or m_CanEdit Then Exit Sub
    Select Case Index
    Case 0
        If chkProps(Index).Value Then
            Select Case True
            Case optProp1(0).Value
                m_Items(cboProps.ListIndex + 1).lGet = "P"
            Case optProp1(1).Value
                m_Items(cboProps.ListIndex + 1).lGet = "V"
            Case optProp1(2).Value
                m_Items(cboProps.ListIndex + 1).lGet = "F"
            Case Else
                optProp1(0).Value = True
            End Select
        Else
            m_Items(cboProps.ListIndex + 1).lGet = vbNullString
        End If
    Case 1
        If chkProps(Index).Value Then
            Select Case True
            Case optProp2(0).Value
                m_Items(cboProps.ListIndex + 1).lLet = "P"
            Case optProp2(1).Value
                m_Items(cboProps.ListIndex + 1).lLet = "V"
            Case optProp2(2).Value
                m_Items(cboProps.ListIndex + 1).lLet = "F"
            Case Else
                optProp2(0).Value = True
            End Select
        Else
            m_Items(cboProps.ListIndex + 1).lLet = vbNullString
        End If
    Case 2
        If chkProps(Index).Value Then
            Select Case True
            Case optProp3(0).Value
                m_Items(cboProps.ListIndex + 1).lSet = "P"
            Case optProp3(1).Value
                m_Items(cboProps.ListIndex + 1).lSet = "V"
            Case optProp3(2).Value
                m_Items(cboProps.ListIndex + 1).lSet = "F"
            Case Else
                optProp3(0).Value = True
            End Select
        Else
            m_Items(cboProps.ListIndex + 1).lSet = vbNullString
        End If
    End Select
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtName.Text) <> vbNullString Then
        Call Add(txtName.Text)
        txtName.Text = vbNullString
    End If
End Sub

Private Sub cmdGenerate_Click()
    Dim cApp1 As New cAppendString
    Dim cApp2 As New cAppendString
    Dim cApp3 As New cAppendString
    Dim cApp4 As New cAppendString
    Dim cApp5 As New cAppendString
    Dim cApp6 As New cAppendString
    Dim xR    As Boolean
    Dim xR2   As Boolean
    Dim I    As Long
    Call cApp1.Append("Option Explicit" & vbCrLf & vbCrLf)
    Call cApp3.Append("Private Sub UserControl_WriteProperties(ByVal PropBag As PropertyBag)" & vbCrLf & vbTab & "With PropBag" & vbCrLf)
    Call cApp4.Append("Private Sub UserControl_ReadProperties(ByVal PropBag As PropertyBag)" & vbCrLf & vbTab & "With PropBag" & vbCrLf)
    Call cApp5.Append("Private Sub pTerminate" & vbCrLf)
    Call cApp6.Append("Private Sub UserControl_InitProperties()" & vbCrLf)
    For I = 1 To m_Count
        xR = False
        With m_Items(I)
            Call cApp1.Append("Private m_" & .lName & " As " & .lData & vbCrLf)
            If .lPara = vbNullString Then
                Select Case .lData
                Case "Boolean", "Byte", "Single", "Double", "Long", "Currency", "OLE_COLOR", "String", "IFontDisp", "IPictureDisp", "StdFont", "StdPicture"
                    Call cApp3.Append(vbTab & vbTab & "Call .WriteProperty(""" & .lName & """ , " & .lName & ")" & vbCrLf)
                End Select
            End If
            Select Case .lGet
            Case "P"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Public Property Get " & .lName & "() As " & .lData)
                Else
                    Call cApp2.Append("Public Property Get " & .lName & "(" & .lPara & ") As " & .lData)
                End If
                Select Case .lData
                Case "Boolean", "Byte", "Single", "Double", "Long", "Currency", "OLE_COLOR", "String"
                    Call cApp2.Append(vbCrLf & vbTab & .lName & " = " & "m_" & .lName & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                Case Else
                    Call cApp2.Append(vbCrLf & vbTab & "Set " & .lName & " = " & "m_" & .lName & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                End Select
            Case "V"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Private Property Get " & .lName & "() As " & .lData)
                Else
                    Call cApp2.Append("Private Property Get " & .lName & "(" & .lPara & ") As " & .lData)
                End If
                Select Case .lData
                Case "Boolean", "Byte", "Single", "Double", "Long", "Currency", "OLE_COLOR", "String"
                    Call cApp2.Append(vbCrLf & vbTab & .lName & " = " & "m_" & .lName & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                Case Else
                    Call cApp2.Append(vbCrLf & vbTab & "Set " & .lName & " = " & "m_" & .lName & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                End Select
            Case "F"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Friend Property Get " & .lName & "() As " & .lData)
                Else
                    Call cApp2.Append("Friend Property Get " & .lName & "(" & .lPara & ") As " & .lData)
                End If
                Select Case .lData
                Case "Boolean", "Byte", "Single", "Double", "Long", "Currency", "OLE_COLOR", "String"
                    Call cApp2.Append(vbCrLf & vbTab & .lName & " = " & "m_" & .lName & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                Case Else
                    Call cApp2.Append(vbCrLf & vbTab & "Set " & .lName & " = " & "m_" & .lName & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                End Select
            End Select
        
            Select Case .lLet
            Case "P"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Public Property Let " & .lName & "(ByVal vNewValue As " & .lData & ")")
                Else
                    Call cApp2.Append("Public Property Let " & .lName & "(" & .lPara & ", vNewValue As " & .lData & ")")
                End If
                Call cApp2.Append(vbCrLf & vbTab & "m_" & .lName & " = vNewValue" & vbCrLf & vbTab & "Call PropertyChanged(""" & .lName & """)" & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                xR = True
            Case "V"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Private Property Let " & .lName & "(ByVal vNewValue As " & .lData & ")")
                Else
                    Call cApp2.Append("Private Property Let " & .lName & "(" & .lPara & ", vNewValue As " & .lData & ")")
                End If
                Call cApp2.Append(vbCrLf & vbTab & "m_" & .lName & " = vNewValue" & vbCrLf & vbTab & "Call PropertyChanged(""" & .lName & """)" & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                xR = True
            Case "F"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Friend Property Let " & .lName & "(ByVal vNewValue As " & .lData & ")")
                Else
                    Call cApp2.Append("Friend Property Let " & .lName & "(" & .lPara & ", vNewValue As " & .lData & ")")
                End If
                Call cApp2.Append(vbCrLf & vbTab & "m_" & .lName & " = vNewValue" & vbCrLf & vbTab & "Call PropertyChanged(""" & .lName & """)" & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                xR = True
            End Select
            If .lPara = vbNullString Then
                Select Case .lData
                Case "Boolean", "Byte", "Single", "Double", "Long", "Currency", "OLE_COLOR", "String", "IFontDisp", "IPictureDisp", "StdFont", "StdPicture"
                    If xR Then
                        If Trim(.lDefValue) = vbNullString Then
                            Call cApp4.Append(vbTab & vbTab & .lName & " = .ReadProperty(""" & .lName & """)" & vbCrLf)
                        Else
                            Call cApp4.Append(vbTab & vbTab & .lName & " = .ReadProperty(""" & .lName & """ , " & .lDefValue & ")" & vbCrLf)
                            Call cApp6.Append(vbTab & .lName & " = " & .lDefValue & vbCrLf)
                        End If
                        xR2 = True
                    End If
                    xR = False
                End Select
            End If
            
            Select Case .lSet
            Case "P"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Public Property Set " & .lName & "(ByVal vNewValue As " & .lData & ")")
                Else
                    Call cApp2.Append("Public Property Set " & .lName & "(" & .lPara & ", vNewValue As " & .lData & ")")
                End If
                Call cApp2.Append(vbCrLf & vbTab & "Set m_" & .lName & " = vNewValue" & vbCrLf & vbTab & "Call PropertyChanged(""" & .lName & """)" & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                xR = True
            Case "V"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Private Property Set " & .lName & "(ByVal vNewValue As " & .lData & ")")
                Else
                    Call cApp2.Append("Private Property Set " & .lName & "(" & .lPara & ", vNewValue As " & .lData & ")")
                End If
                Call cApp2.Append(vbCrLf & vbTab & "Set m_" & .lName & " = vNewValue" & vbCrLf & vbTab & "Call PropertyChanged(""" & .lName & """)" & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                xR = True
            Case "F"
                If Trim(.lPara) = vbNullString Then
                    Call cApp2.Append("Friend Property Set " & .lName & "(ByVal vNewValue As " & .lData & ")")
                Else
                    Call cApp2.Append("Friend Property Set " & .lName & "(" & .lPara & ", vNewValue As " & .lData & ")")
                End If
                Call cApp2.Append(vbCrLf & vbTab & "Set m_" & .lName & " = vNewValue" & vbCrLf & vbTab & "Call PropertyChanged(""" & .lName & """)" & vbCrLf & "End Property" & vbCrLf & vbCrLf)
                xR = True
            End Select
            If .lPara = vbNullString Then
                Select Case .lData
                Case "Boolean", "Byte", "Single", "Double", "Long", "Currency", "OLE_COLOR", "String", "IFontDisp", "IPictureDisp", "StdFont", "StdPicture"
                    If xR And Not xR2 Then
                        If Trim(.lDefValue) = vbNullString Then
                            Call cApp4.Append(vbTab & vbTab & "Set " & .lName & " = .ReadProperty(""" & .lName & """)" & vbCrLf)
                        Else
                            Call cApp4.Append(vbTab & vbTab & "Set " & .lName & " = .ReadProperty(""" & .lName & """ , " & .lDefValue & ")" & vbCrLf)
                            Call cApp6.Append(vbTab & "Set " & .lName & " = " & .lDefValue & vbCrLf)
                        End If
                    End If
                End Select
            End If
            Select Case .lData
            Case "Boolean", "Byte", "Single", "Double", "Long", "Currency", "OLE_COLOR"
            Case "String"
                Call cApp5.Append(vbTab & .lName & " = vbNullString" & vbCrLf)
            Case Else
                Call cApp5.Append(vbTab & "Set " & .lName & " = Nothing" & vbCrLf)
            End Select
        End With
    Next
    Call cApp5.Append("End Sub")
    Call cApp6.Append("End Sub")
    Call cApp3.Append(vbTab & "End With" & vbCrLf & "End Sub")
    Call cApp4.Append(vbTab & "End With" & vbCrLf & "End Sub")
    Call cApp1.Append(vbCrLf & cApp2.Value)
    Call cApp1.Append(vbCrLf & cApp3.Value)
    Call cApp1.Append(vbCrLf & vbCrLf & cApp4.Value)
    Call cApp1.Append(vbCrLf & vbCrLf & cApp5.Value)
    Call cApp1.Append(vbCrLf & vbCrLf & cApp6.Value)
    Call cApp2.Clear
    Set cApp2 = Nothing
    Dim New_Form As New frmGenerate
    New_Form.Text.Text = cApp1.Value
    Call cApp1.Clear
    Set cApp1 = Nothing
    Call cApp3.Clear
    Call cApp4.Clear
    Call cApp5.Clear
    Call cApp6.Clear
    Set cApp3 = Nothing
    Set cApp4 = Nothing
    Set cApp5 = Nothing
    Set cApp6 = Nothing
    Call New_Form.Show
End Sub

Private Sub cmdRemove_Click()
    Call Remove(cboProps.ListIndex + 1)
End Sub

Private Sub Form_Load()
'    cboDataType.ListIndex = 0
End Sub

Private Sub Add(ByVal Name As String)
    m_Count = m_Count + 1
    ReDim Preserve m_Items(1 To m_Count) As typProperty
    With m_Items(m_Count)
        .lName = Name
        .lGet = "P"
        .lLet = "P"
        .lData = "String"
        Call cboProps.AddItem(.lName)
        cboProps.ListIndex = m_Count - 1
    End With
End Sub

Private Sub Remove(ByVal Index As Long)
    If Index <> m_Count Then
        Dim I As Long
        For I = Index To m_Count - 1
            LSet m_Items(I) = m_Items(I + 1)
        Next
    End If
    m_Count = m_Count - 1
    If m_Count = 0 Then
        Erase m_Items
    Else
        ReDim Preserve m_Items(1 To m_Count) As typProperty
    End If
    Call cboProps.RemoveItem(Index - 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase m_Items
End Sub

Private Sub lstDType_Click()
    If cboProps.ListIndex = -1 Or m_CanEdit Then Exit Sub
    txtDType.Text = lstDType.Text
End Sub

Private Sub optProp1_Click(Index As Integer)
    If cboProps.ListIndex = -1 Or m_CanEdit Then Exit Sub
    If m_Items(cboProps.ListIndex + 1).lGet = vbNullString Then
        m_CanEdit = True
        chkProps(0).Value = 1
        m_CanEdit = False
    End If
    Select Case True
    Case optProp1(0).Value
        m_Items(cboProps.ListIndex + 1).lGet = "P"
    Case optProp1(1).Value
        m_Items(cboProps.ListIndex + 1).lGet = "V"
    Case optProp1(2).Value
        m_Items(cboProps.ListIndex + 1).lGet = "F"
    End Select
End Sub

Private Sub optProp2_Click(Index As Integer)
    If cboProps.ListIndex = -1 Or m_CanEdit Then Exit Sub
    If m_Items(cboProps.ListIndex + 1).lLet = vbNullString Then
        m_CanEdit = True
        chkProps(1).Value = 1
        m_CanEdit = False
    End If
    Select Case True
    Case optProp2(0).Value
        m_Items(cboProps.ListIndex + 1).lLet = "P"
    Case optProp2(1).Value
        m_Items(cboProps.ListIndex + 1).lLet = "V"
    Case optProp2(2).Value
        m_Items(cboProps.ListIndex + 1).lLet = "F"
    End Select
End Sub

Private Sub optProp3_Click(Index As Integer)
    If cboProps.ListIndex = -1 Or m_CanEdit Then Exit Sub
    If m_Items(cboProps.ListIndex + 1).lSet = vbNullString Then
        m_CanEdit = True
        chkProps(2).Value = 1
        m_CanEdit = False
    End If
    Select Case True
    Case optProp3(0).Value
        m_Items(cboProps.ListIndex + 1).lSet = "P"
    Case optProp3(1).Value
        m_Items(cboProps.ListIndex + 1).lSet = "V"
    Case optProp3(2).Value
        m_Items(cboProps.ListIndex + 1).lSet = "F"
    End Select
End Sub

Private Sub txtDefValue_Change()
    If cboProps.ListIndex = -1 Or m_CanEdit Then Exit Sub
    m_Items(cboProps.ListIndex + 1).lDefValue = txtDefValue.Text
End Sub

Private Sub txtDType_Change()
    If cboProps.ListIndex = -1 Or m_CanEdit Then Exit Sub
    m_Items(cboProps.ListIndex + 1).lData = txtDType.Text
End Sub

Private Sub txtPara_Change()
    If cboProps.ListIndex = -1 Or m_CanEdit Then Exit Sub
    m_Items(cboProps.ListIndex + 1).lPara = txtPara.Text
End Sub
