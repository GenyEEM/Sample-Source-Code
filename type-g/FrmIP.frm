VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FrmIP 
   Caption         =   "TCP Setting of MUTs"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   5190
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   780
   End
   Begin MSFlexGridLib.MSFlexGrid mfgIP 
      Height          =   4335
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7646
      _Version        =   393216
      BackColorBkg    =   16777215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtMeterNu 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Text            =   "6"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Meter no.:"
      Height          =   180
      Index           =   25
      Left            =   1320
      TabIndex        =   1
      Top             =   165
      Width           =   900
   End
End
Attribute VB_Name = "FrmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRow As Integer
Dim iCol As Integer

Private Sub Command1_Click()
     Dim i As Integer
    Dim strTmp As String
    
    Call SaveIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtMeterNu", txtMeterNu.Text)
    
    With mfgIP
    
        For i = 1 To .Rows - 1
           If .TextMatrix(i, 1) <> "" Then
              strTmp = .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
              Call SaveIniInfo(ProgramPath & "LanComTest.ini", "MeterIP", "R" & i, strTmp)
           End If
        Next i
    End With
    
    MsgBox "保存成功！", vbInformation, "系统提示..."
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   txtMeterNu.Text = GetIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtMeterNu")
   
   mfgIP.Rows = Val(txtMeterNu.Text) + 1
   Call inimfgIP
End Sub

Private Sub inimfgIP()
    Dim i      As Integer
    Dim strTmp As String

    With mfgIP
        .Clear
        .Cols = 4
        If .Rows = 0 Then .Rows = 1
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "常数"
        .ColWidth(1) = 1000
        .TextMatrix(0, 2) = "IP"
        .ColWidth(2) = 1800
        .TextMatrix(0, 3) = "Port"
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
            strTmp = GetIniInfo(ProgramPath & "LanComTest.ini", "MeterIP", "R" & i)

            If strTmp <> "" Then
                .TextMatrix(i, 1) = GetItem(strTmp, ",", 0)
                .TextMatrix(i, 2) = GetItem(strTmp, ",", 1)
                .TextMatrix(i, 3) = GetItem(strTmp, ",", 2)
            End If

        Next i

    End With
End Sub

Private Sub mfgIP_DblClick()
    With mfgIP
        iCol = .Col
        iRow = .Row
        If .Row > 0 And .Col > 0 Then
            Text1.Top = .CellTop + .Top
            Text1.Left = .CellLeft + .Left
            Text1.Width = .CellWidth
            Text1.Height = .CellHeight
            Text1.Text = .Text
            Text1.Visible = True
            Text1.SetFocus
        End If
    End With
End Sub

Private Sub Text1_LostFocus()
    With mfgIP
        .TextMatrix(iRow, iCol) = Text1.Text
        Text1.Text = ""
        Text1.Visible = False
    End With
End Sub

Private Sub txtMeterNu_Change()
   mfgIP.Rows = Val(txtMeterNu.Text) + 1
   Call inimfgIP
End Sub
