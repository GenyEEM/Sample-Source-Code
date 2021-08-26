VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FrmMain 
   Caption         =   "YC189G Test bench toolkit"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   10365
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Debug test"
      TabPicture(0)   =   "FrmMain.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Check1"
      Tab(0).Control(2)=   "cmdStop"
      Tab(0).Control(3)=   "cmdSend"
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(6)=   "Frame1(0)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Power Supply"
      TabPicture(1)   =   "FrmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(61)"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame1(1)"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "Frame10"
      Tab(1).Control(5)=   "txtData"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Harmonic"
      TabPicture(2)   =   "FrmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(3)"
      Tab(2).Control(1)=   "Frame1(4)"
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(3)=   "Command43"
      Tab(2).Control(4)=   "Command44"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Stability test"
      TabPicture(3)   =   "FrmMain.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame11"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.CommandButton Command44 
         Caption         =   "Read"
         Height          =   360
         Left            =   -67560
         TabIndex        =   260
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton Command43 
         Caption         =   "Output"
         Height          =   360
         Left            =   -66240
         TabIndex        =   259
         Top             =   7080
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   238
         Top             =   6360
         Width           =   9855
      End
      Begin VB.Frame Frame11 
         Height          =   7095
         Left            =   240
         TabIndex        =   143
         Top             =   460
         Width           =   9855
         Begin VB.CommandButton Command45 
            Caption         =   "Multi-ErrCounter"
            Height          =   360
            Left            =   480
            TabIndex        =   280
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Load Switch"
            Height          =   315
            Index           =   6
            Left            =   2280
            TabIndex        =   275
            Top             =   2520
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox Check6 
            Caption         =   "485"
            Height          =   195
            Index           =   5
            Left            =   2280
            TabIndex        =   269
            Top             =   2300
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Err%"
            Height          =   195
            Index           =   4
            Left            =   2280
            TabIndex        =   266
            Top             =   2040
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1680
            TabIndex        =   264
            Top             =   6240
            Width           =   7815
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1680
            TabIndex        =   262
            Text            =   "FE FE FE FE 68 AA AA AA AA AA AA 68 11 04 33 33 34 33 AE 16"
            Top             =   5925
            Width           =   7815
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Error Counter"
            Height          =   195
            Left            =   4680
            TabIndex        =   261
            Top             =   840
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   9000
            TabIndex        =   258
            Text            =   "5"
            Top             =   1260
            Width           =   615
         End
         Begin VB.CommandButton Command42 
            Caption         =   "Move down"
            Height          =   360
            Left            =   2400
            TabIndex        =   255
            Top             =   6600
            Width           =   975
         End
         Begin VB.CommandButton Command41 
            Caption         =   "Move up"
            Height          =   360
            Left            =   1440
            TabIndex        =   254
            Top             =   6600
            Width           =   855
         End
         Begin VB.CommandButton Command40 
            Caption         =   "MUTs IP"
            Height          =   360
            Left            =   8640
            TabIndex        =   253
            Top             =   2880
            Width           =   1095
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Ref.Co"
            Height          =   195
            Index           =   3
            Left            =   4680
            TabIndex        =   250
            Top             =   2520
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "FrmMain.frx":0070
            Left            =   7920
            List            =   "FrmMain.frx":007A
            TabIndex        =   249
            Text            =   "1"
            Top             =   2880
            Width           =   615
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmMain.frx":0084
            Left            =   6840
            List            =   "FrmMain.frx":0091
            TabIndex        =   248
            Text            =   "1000000000"
            Top             =   3285
            Width           =   1335
         End
         Begin MSWinsockLib.Winsock Winsock1 
            Left            =   120
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.TextBox txtReadError 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   9000
            TabIndex        =   220
            Text            =   "0.1"
            Top             =   960
            Width           =   615
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Built-in Ref."
            Height          =   195
            Left            =   5880
            TabIndex        =   219
            Top             =   600
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox Check6 
            Caption         =   "C"
            Height          =   195
            Index           =   2
            Left            =   4680
            TabIndex        =   212
            Top             =   2280
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox Check6 
            Caption         =   "B"
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   211
            Top             =   2040
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   120
            Top             =   960
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Save"
            Height          =   360
            Left            =   240
            TabIndex        =   176
            Top             =   6600
            Width           =   1095
         End
         Begin VB.CommandButton Command34 
            Caption         =   "Delete"
            Height          =   360
            Left            =   9000
            TabIndex        =   175
            Top             =   3240
            Width           =   735
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Add"
            Height          =   360
            Left            =   8280
            TabIndex        =   174
            Top             =   3240
            Width           =   735
         End
         Begin MSFlexGridLib.MSFlexGrid msfgComTest 
            Height          =   2295
            Left            =   240
            TabIndex        =   173
            Top             =   3600
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4048
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedRows       =   0
            BackColorBkg    =   16777215
         End
         Begin VB.TextBox txtWaitTIme 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   9000
            TabIndex        =   168
            Text            =   "10"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton cmdComStop 
            Caption         =   "Stop"
            Height          =   360
            Left            =   8520
            TabIndex        =   167
            Top             =   6600
            Width           =   1095
         End
         Begin VB.TextBox txtTestTime 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   9000
            TabIndex        =   165
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdComStart 
            Caption         =   "Start"
            Height          =   360
            Left            =   7320
            TabIndex        =   164
            Top             =   6600
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid msfgComStandard 
            Height          =   1575
            Left            =   360
            TabIndex        =   157
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   0
            BackColorBkg    =   -2147483633
            FocusRect       =   0
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            BorderStyle     =   0
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CheckBox chkComReadStandard 
            Caption         =   "Ref. reading"
            Height          =   195
            Left            =   4680
            TabIndex        =   156
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   26
            Left            =   2280
            TabIndex        =   149
            Text            =   "220"
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   25
            Left            =   3360
            TabIndex        =   148
            Text            =   "1"
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   24
            Left            =   4560
            TabIndex        =   147
            Text            =   "0"
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   23
            Left            =   5760
            TabIndex        =   146
            Text            =   "0"
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   22
            Left            =   6840
            TabIndex        =   145
            Text            =   "50"
            Top             =   2880
            Width           =   495
         End
         Begin VB.ComboBox cmbComModel 
            Height          =   315
            Left            =   840
            TabIndex        =   144
            Text            =   "0-Active Power"
            Top             =   2880
            Width           =   855
         End
         Begin VB.CheckBox Check6 
            Caption         =   "A"
            Height          =   195
            Index           =   0
            Left            =   4680
            TabIndex        =   204
            Top             =   1800
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ready to test..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Index           =   77
            Left            =   4680
            TabIndex        =   279
            Top             =   1200
            Width           =   2190
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "MUT:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   278
            Top             =   6000
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fails:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   76
            Left            =   3240
            TabIndex        =   277
            Top             =   2520
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   75
            Left            =   4080
            TabIndex        =   276
            Top             =   2520
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Read(second):"
            Height          =   195
            Index           =   74
            Left            =   7680
            TabIndex        =   274
            Top             =   1275
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Repeat:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   73
            Left            =   8040
            TabIndex        =   273
            Top             =   1800
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   72
            Left            =   9000
            TabIndex        =   272
            Top             =   1800
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   71
            Left            =   4080
            TabIndex        =   271
            Top             =   2300
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fails:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   70
            Left            =   3240
            TabIndex        =   270
            Top             =   2295
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   69
            Left            =   4080
            TabIndex        =   268
            Top             =   2040
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fails:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   68
            Left            =   3240
            TabIndex        =   267
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label Label8 
            Caption         =   "485 received: "
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   265
            Top             =   6315
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "485 sent:"
            Height          =   195
            Index           =   4
            Left            =   720
            TabIndex        =   263
            Top             =   6000
            Width           =   690
         End
         Begin VB.Label labStdConstant 
            BackColor       =   &H00000000&
            Caption         =   "??????"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   720
            TabIndex        =   257
            Top             =   1760
            Width           =   3240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Co:  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   67
            Left            =   360
            TabIndex        =   256
            Top             =   1760
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fails:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   65
            Left            =   5760
            TabIndex        =   251
            Top             =   2520
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   66
            Left            =   6600
            TabIndex        =   252
            Top             =   2520
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ref.Co:"
            Height          =   195
            Index           =   64
            Left            =   5880
            TabIndex        =   247
            Top             =   3330
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Loop:"
            Height          =   195
            Index           =   63
            Left            =   7440
            TabIndex        =   246
            Top             =   2925
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Interval(se.):"
            Height          =   195
            Index           =   56
            Left            =   7680
            TabIndex        =   221
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fails:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   48
            Left            =   5760
            TabIndex        =   206
            Top             =   1800
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fails:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   52
            Left            =   5760
            TabIndex        =   210
            Top             =   2280
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   51
            Left            =   6600
            TabIndex        =   209
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fails:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   50
            Left            =   5760
            TabIndex        =   208
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Index           =   49
            Left            =   6600
            TabIndex        =   207
            Top             =   2040
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   240
            Index           =   47
            Left            =   6600
            TabIndex        =   205
            Top             =   1800
            Width           =   120
         End
         Begin VB.Label Label2 
            Caption         =   "Load points test:"
            Height          =   195
            Index           =   41
            Left            =   360
            TabIndex        =   172
            Top             =   3320
            Width           =   5460
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "00:00:00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   40
            Left            =   5880
            TabIndex        =   171
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Duration:"
            Height          =   195
            Index           =   39
            Left            =   4680
            TabIndex        =   170
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Wait for(s):"
            Height          =   195
            Index           =   38
            Left            =   8040
            TabIndex        =   169
            Top             =   600
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Test long(hr):"
            Height          =   195
            Index           =   37
            Left            =   7680
            TabIndex        =   166
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Index           =   36
            Left            =   9000
            TabIndex        =   163
            Top             =   2520
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   35
            Left            =   9000
            TabIndex        =   162
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   34
            Left            =   9000
            TabIndex        =   161
            Top             =   2040
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Success rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   33
            Left            =   7545
            TabIndex        =   160
            Top             =   2520
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fail rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   32
            Left            =   7920
            TabIndex        =   159
            Top             =   2280
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Port times:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   31
            Left            =   7800
            TabIndex        =   158
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mode:"
            Height          =   195
            Index           =   30
            Left            =   360
            TabIndex        =   155
            Top             =   2925
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "U(V):"
            Height          =   195
            Index           =   29
            Left            =   1800
            TabIndex        =   154
            Top             =   2925
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "I(A):"
            Height          =   195
            Index           =   28
            Left            =   2880
            TabIndex        =   153
            Top             =   2925
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ang.UI:"
            Height          =   195
            Index           =   27
            Left            =   3960
            TabIndex        =   152
            Top             =   2925
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ang.UU:"
            Height          =   195
            Index           =   26
            Left            =   5040
            TabIndex        =   151
            Top             =   2925
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fre:"
            Height          =   195
            Index           =   25
            Left            =   6360
            TabIndex        =   150
            Top             =   2925
            Width           =   300
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Reference meter"
         Height          =   735
         Left            =   -74760
         TabIndex        =   137
         Top             =   5280
         Width           =   9975
         Begin VB.CommandButton Command31 
            Caption         =   "read"
            Height          =   360
            Left            =   7440
            TabIndex        =   142
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Output"
            Height          =   360
            Left            =   8640
            TabIndex        =   141
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optComNo 
            Caption         =   "#2 port"
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   140
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optComNo 
            Caption         =   "#1 port"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   139
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Loop of readings"
            Height          =   255
            Left            =   2880
            TabIndex        =   138
            Top             =   300
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Harmonic injection"
         Height          =   4455
         Left            =   -74880
         TabIndex        =   92
         Top             =   360
         Width           =   9975
         Begin VB.ComboBox cmbSource2 
            Height          =   315
            Left            =   8760
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chkAutoReadXB 
            Caption         =   "Reading loop"
            Height          =   255
            Left            =   4560
            TabIndex        =   111
            Top             =   2320
            Width           =   1575
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Output"
            Height          =   360
            Left            =   8640
            TabIndex        =   110
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Read"
            Height          =   360
            Left            =   6240
            TabIndex        =   106
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Clear"
            Height          =   360
            Left            =   7440
            TabIndex        =   105
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Frame Frame7 
            Caption         =   "Harmonic settings"
            Height          =   1455
            Left            =   4440
            TabIndex        =   93
            Top             =   240
            Width           =   5535
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   10
               Left            =   1080
               TabIndex        =   101
               Text            =   "10"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   11
               Left            =   2280
               TabIndex        =   100
               Text            =   "0"
               ToolTipText     =   "介于500(毫秒)到200000(毫秒)之间"
               Top             =   360
               Width           =   615
            End
            Begin VB.ComboBox cmbXB 
               Height          =   315
               Left            =   3600
               Style           =   2  'Dropdown List
               TabIndex        =   99
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton optXB 
               Caption         =   "Current"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   98
               Top             =   1080
               Width           =   1215
            End
            Begin VB.OptionButton optXB 
               Caption         =   "Voltage"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   97
               Top             =   840
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.CommandButton Command10 
               Caption         =   "<Save>"
               Height          =   360
               Left            =   4200
               TabIndex        =   96
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Delete>>"
               Height          =   360
               Left            =   3000
               TabIndex        =   95
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton Command9 
               Caption         =   "<<Add"
               Height          =   360
               Left            =   1800
               TabIndex        =   94
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amp%"
               Height          =   195
               Index           =   12
               Left            =   525
               TabIndex        =   104
               Top             =   405
               Width           =   480
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ang:"
               Height          =   195
               Index           =   11
               Left            =   1800
               TabIndex        =   103
               Top             =   405
               Width           =   345
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Order"
               Height          =   195
               Index           =   10
               Left            =   3045
               TabIndex        =   102
               Top             =   405
               Width           =   420
            End
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   4095
            Left            =   240
            TabIndex        =   107
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   7223
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Send"
            TabPicture(0)   =   "FrmMain.frx":00B6
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "msfgXB"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Read"
            TabPicture(1)   =   "FrmMain.frx":00D2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            Begin MSFlexGridLib.MSFlexGrid msfgXB 
               Height          =   3615
               Left            =   120
               TabIndex        =   108
               Top             =   360
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   6376
               _Version        =   393216
               Rows            =   13
               Cols            =   4
               BackColorBkg    =   -2147483633
               FocusRect       =   0
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid msfgReadXB 
               Height          =   1935
               Left            =   -74880
               TabIndex        =   109
               Top             =   360
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   3413
               _Version        =   393216
               Rows            =   22
               Cols            =   4
               BackColorBkg    =   -2147483633
               FocusRect       =   0
               SelectionMode   =   1
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mode: "
            Height          =   195
            Index           =   14
            Left            =   8160
            TabIndex        =   218
            Top             =   1845
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note:First order must be fundamental wave"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   55
            Left            =   4560
            TabIndex        =   216
            Top             =   1800
            Width           =   3150
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Serial port test"
         Height          =   1575
         Index           =   4
         Left            =   -74760
         TabIndex        =   60
         Top             =   4800
         Width           =   9975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   1815
            TabIndex        =   242
            Top             =   480
            Width           =   1815
            Begin VB.OptionButton Option1 
               Caption         =   "GB645-1997-2"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   245
               Top             =   600
               Width           =   1575
            End
            Begin VB.OptionButton Option1 
               Caption         =   "GB645-1997-1"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   244
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton Option1 
               Caption         =   "GB645-2007"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   243
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.TextBox txt485Door 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6240
            TabIndex        =   136
            Text            =   "1"
            Top             =   200
            Width           =   375
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Read"
            Height          =   360
            Left            =   8400
            TabIndex        =   76
            Top             =   200
            Width           =   735
         End
         Begin VB.TextBox txt485Model 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5640
            TabIndex        =   74
            Text            =   "1"
            Top             =   200
            Width           =   375
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Read Myset"
            Height          =   360
            Left            =   7440
            TabIndex        =   71
            Top             =   200
            Width           =   975
         End
         Begin VB.TextBox txtRs485Receive 
            Height          =   375
            Left            =   2640
            TabIndex        =   69
            Top             =   1080
            Width           =   7215
         End
         Begin VB.TextBox txtRs485Send 
            Height          =   405
            Left            =   2640
            TabIndex        =   67
            Text            =   "FE FE FE FE 68 AA AA AA AA AA AA 68 11 04 33 33 34 33 AE 16"
            Top             =   600
            Width           =   7215
         End
         Begin VB.TextBox txtBaueRate 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3960
            TabIndex        =   65
            Text            =   "2400,E,8,1"
            Top             =   200
            Width           =   1215
         End
         Begin VB.OptionButton optRs485 
            Caption         =   "Port2(RS485)"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   64
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optRs485 
            Caption         =   "Port 1(RS232)"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   63
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Output"
            Height          =   360
            Left            =   9120
            TabIndex        =   62
            Top             =   200
            Width           =   735
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Set"
            Height          =   360
            Left            =   6720
            TabIndex        =   61
            Top             =   200
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   6000
            X2              =   6240
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label Label8 
            Caption         =   "Mode："
            Height          =   255
            Index           =   3
            Left            =   5160
            TabIndex        =   75
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Read: "
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   70
            Top             =   1185
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Sent:"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   68
            Top             =   705
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "BaudRate："
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   66
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "流水线控制"
         Height          =   495
         Index           =   3
         Left            =   -74640
         TabIndex        =   56
         Top             =   6840
         Visible         =   0   'False
         Width           =   1095
         Begin VB.CommandButton Command28 
            Caption         =   "测试/停止"
            Height          =   360
            Left            =   8760
            TabIndex        =   135
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton Command27 
            Caption         =   "检测结果"
            Height          =   360
            Left            =   8760
            TabIndex        =   134
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton Command26 
            Caption         =   "全出表滚轮"
            Height          =   360
            Left            =   8760
            TabIndex        =   133
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton Command25 
            Caption         =   "读条码"
            Height          =   360
            Left            =   8760
            TabIndex        =   132
            Top             =   120
            Width           =   1215
         End
         Begin VB.Frame Frame9 
            Caption         =   "出表控制"
            Height          =   735
            Left            =   3720
            TabIndex        =   85
            Top             =   840
            Width           =   3735
            Begin VB.CommandButton Command24 
               Caption         =   "滚轮"
               Height          =   360
               Left            =   3000
               TabIndex        =   91
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optOutPut 
               Caption         =   "左转"
               Height          =   195
               Index           =   2
               Left            =   2280
               TabIndex        =   89
               Top             =   285
               Width           =   855
            End
            Begin VB.OptionButton optOutPut 
               Caption         =   "右转"
               Height          =   195
               Index           =   1
               Left            =   1560
               TabIndex        =   88
               Top             =   285
               Width           =   855
            End
            Begin VB.OptionButton optOutPut 
               Caption         =   "停止"
               Height          =   195
               Index           =   0
               Left            =   840
               TabIndex        =   87
               Top             =   285
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.CommandButton Command16 
               Caption         =   "出 表"
               Height          =   360
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "入表控制"
            Height          =   735
            Left            =   3720
            TabIndex        =   81
            Top             =   0
            Width           =   3735
            Begin VB.CommandButton Command23 
               Caption         =   "滚轮"
               Height          =   360
               Left            =   3000
               TabIndex        =   90
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optInPut 
               Caption         =   "启动"
               Height          =   195
               Index           =   1
               Left            =   1920
               TabIndex        =   84
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton optInPut 
               Caption         =   "停止"
               Height          =   195
               Index           =   0
               Left            =   1080
               TabIndex        =   83
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.CommandButton Command15 
               Caption         =   "入 表"
               Height          =   360
               Left            =   120
               TabIndex        =   82
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.ComboBox cmbPort 
            Height          =   315
            Left            =   2640
            TabIndex        =   79
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cmbIP 
            Height          =   315
            Left            =   600
            TabIndex        =   77
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command13 
            Caption         =   "读装置状态"
            Height          =   360
            Left            =   7560
            TabIndex        =   73
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "读版本号"
            Height          =   360
            Left            =   7560
            TabIndex        =   72
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command18 
            Caption         =   "读 取"
            Height          =   360
            Left            =   7560
            TabIndex        =   59
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton Command17 
            Caption         =   "复 位"
            Height          =   360
            Left            =   7560
            TabIndex        =   57
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port:"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   80
            Top             =   285
            Width           =   330
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   78
            Top             =   285
            Width           =   210
         End
         Begin VB.Label Label7 
            Caption         =   "读取结果："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   915
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Width           =   3615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Error Counter"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   42
         Top             =   3600
         Width           =   9975
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   48
            Left            =   9240
            TabIndex        =   241
            Text            =   "0.97"
            Top             =   765
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   47
            Left            =   6840
            TabIndex        =   236
            Text            =   "0"
            Top             =   765
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   46
            Left            =   5520
            TabIndex        =   234
            Text            =   "1"
            Top             =   765
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   45
            Left            =   3960
            TabIndex        =   232
            Text            =   "220"
            Top             =   765
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   9
            Left            =   1920
            TabIndex        =   230
            Text            =   "-1.0"
            Top             =   765
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   8
            Left            =   720
            TabIndex        =   228
            Text            =   "1.0"
            Top             =   765
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   7
            Left            =   9240
            TabIndex        =   226
            Text            =   "3"
            Top             =   315
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   7200
            TabIndex        =   224
            Text            =   "0"
            Top             =   315
            Width           =   1215
         End
         Begin VB.ComboBox cmbErrMode 
            Height          =   315
            Left            =   3120
            TabIndex        =   222
            Text            =   "0"
            Top             =   315
            Width           =   615
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Clear"
            Height          =   360
            Left            =   5040
            TabIndex        =   200
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Reset"
            Height          =   360
            Left            =   7440
            TabIndex        =   54
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkAutoRead 
            Caption         =   "Loop of Err% "
            Height          =   255
            Left            =   2880
            TabIndex        =   51
            Top             =   1140
            Width           =   1455
         End
         Begin VB.OptionButton optErrNo 
            Caption         =   "#1 Counter"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   50
            Top             =   1080
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optErrNo 
            Caption         =   "#2 Counter"
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   49
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbWay 
            Height          =   315
            Left            =   720
            TabIndex        =   48
            Text            =   "0-Normal Error"
            Top             =   315
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   4800
            TabIndex        =   45
            Text            =   "0"
            Top             =   315
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Output"
            Height          =   360
            Left            =   8640
            TabIndex        =   44
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Read"
            Height          =   360
            Left            =   6240
            TabIndex        =   43
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Percentage of test:"
            Height          =   195
            Index           =   62
            Left            =   7680
            TabIndex        =   240
            Top             =   765
            Width           =   1410
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Current:"
            Height          =   195
            Index           =   60
            Left            =   6120
            TabIndex        =   237
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Pulse:"
            Height          =   195
            Index           =   59
            Left            =   4800
            TabIndex        =   235
            Top             =   765
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Starting test U:"
            Height          =   195
            Index           =   58
            Left            =   2760
            TabIndex        =   233
            Top             =   765
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Down:"
            Height          =   195
            Index           =   54
            Left            =   1440
            TabIndex        =   231
            Top             =   765
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Up :"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   229
            Top             =   765
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Time:"
            Height          =   195
            Index           =   6
            Left            =   8640
            TabIndex        =   227
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Starting test I:"
            Height          =   195
            Index           =   5
            Left            =   5880
            TabIndex        =   225
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mode:"
            Height          =   195
            Index           =   57
            Left            =   240
            TabIndex        =   223
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Way:"
            Height          =   195
            Index           =   9
            Left            =   2640
            TabIndex        =   47
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Channel:"
            Height          =   195
            Index           =   8
            Left            =   4080
            TabIndex        =   46
            Top             =   360
            Width           =   645
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ethernet "
         Height          =   975
         Index           =   1
         Left            =   -74760
         TabIndex        =   32
         Top             =   480
         Width           =   9975
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   1
            Left            =   3960
            TabIndex        =   36
            Text            =   "1234"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtIP 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   35
            Text            =   "192.168.1.3"
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtComm 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   1
            Left            =   6840
            TabIndex        =   34
            Text            =   "0"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdLink 
            Caption         =   "Connect"
            Height          =   360
            Index           =   1
            Left            =   8400
            TabIndex        =   33
            Top             =   320
            Width           =   1215
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Port:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   3360
            TabIndex        =   39
            Top             =   405
            Width           =   360
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IP:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   38
            Top             =   405
            Width           =   210
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Port:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   5880
            TabIndex        =   37
            Top             =   405
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ethernet connection"
         Height          =   975
         Index           =   0
         Left            =   -74520
         TabIndex        =   24
         Top             =   480
         Width           =   9495
         Begin VB.CommandButton cmdLink 
            Caption         =   "Connect"
            Height          =   360
            Index           =   0
            Left            =   8040
            TabIndex        =   28
            Top             =   320
            Width           =   1215
         End
         Begin VB.TextBox txtComm 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   6480
            TabIndex        =   27
            Text            =   "0"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtIP 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   960
            TabIndex        =   26
            Text            =   "192.168.1.101"
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   3960
            TabIndex        =   25
            Text            =   "1234"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Port:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   5520
            TabIndex        =   31
            Top             =   405
            Width           =   795
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IP: "
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   405
            Width           =   255
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Port: "
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   29
            Top             =   405
            Width           =   405
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Power supply"
         Height          =   2055
         Left            =   -74760
         TabIndex        =   12
         Top             =   1520
         Width           =   9975
         Begin VB.CommandButton Command39 
            Caption         =   "Read Coeff."
            Height          =   360
            Left            =   5040
            TabIndex        =   215
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton Command38 
            Caption         =   "Correction"
            Height          =   360
            Left            =   120
            TabIndex        =   214
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton Command37 
            Caption         =   "Save coeff."
            Height          =   360
            Left            =   120
            TabIndex        =   213
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtGHAddress 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2115
            TabIndex        =   203
            Text            =   "0"
            Top             =   1605
            Width           =   375
         End
         Begin VB.CommandButton Command36 
            Caption         =   "Read"
            Height          =   360
            Left            =   2520
            TabIndex        =   201
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Output %"
            Height          =   360
            Left            =   3720
            TabIndex        =   199
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   44
            Left            =   4440
            TabIndex        =   198
            Text            =   "10"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   43
            Left            =   2880
            TabIndex        =   197
            Text            =   "10"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   42
            Left            =   4440
            TabIndex        =   196
            Text            =   "10"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   41
            Left            =   2880
            TabIndex        =   195
            Text            =   "10"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   40
            Left            =   4440
            TabIndex        =   194
            Text            =   "10"
            Top             =   440
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   39
            Left            =   2880
            TabIndex        =   193
            Text            =   "10"
            Top             =   440
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   38
            Left            =   7800
            TabIndex        =   188
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   37
            Left            =   6000
            TabIndex        =   187
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   4080
            TabIndex        =   186
            Text            =   "0"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   35
            Left            =   2520
            TabIndex        =   185
            Text            =   "0"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   34
            Left            =   7800
            TabIndex        =   184
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   33
            Left            =   6000
            TabIndex        =   183
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   32
            Left            =   4080
            TabIndex        =   182
            Text            =   "0"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   31
            Left            =   2520
            TabIndex        =   181
            Text            =   "0"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   30
            Left            =   7800
            TabIndex        =   180
            Text            =   "0"
            Top             =   440
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   29
            Left            =   6000
            TabIndex        =   179
            Text            =   "0"
            Top             =   440
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   28
            Left            =   4080
            TabIndex        =   178
            Text            =   "0"
            Top             =   440
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   27
            Left            =   2520
            TabIndex        =   177
            Text            =   "0"
            Top             =   440
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   21
            Left            =   2160
            TabIndex        =   126
            Text            =   "220"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   20
            Left            =   3720
            TabIndex        =   125
            Text            =   "1"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   19
            Left            =   5400
            TabIndex        =   124
            Text            =   "0"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   18
            Left            =   7200
            TabIndex        =   123
            Text            =   "120"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   17
            Left            =   8760
            TabIndex        =   122
            Text            =   "50"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   16
            Left            =   2160
            TabIndex        =   116
            Text            =   "220"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   15
            Left            =   3720
            TabIndex        =   115
            Text            =   "1"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   14
            Left            =   5400
            TabIndex        =   114
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   7200
            TabIndex        =   113
            Text            =   "120"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   8760
            TabIndex        =   112
            Text            =   "50"
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox cmbSource 
            Height          =   315
            Left            =   120
            TabIndex        =   52
            Text            =   "0-Active Power"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Read"
            Height          =   360
            Left            =   6240
            TabIndex        =   41
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Reset"
            Height          =   360
            Left            =   7440
            TabIndex        =   40
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Output"
            Height          =   360
            Left            =   8640
            TabIndex        =   23
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   8760
            TabIndex        =   22
            Text            =   "50"
            Top             =   440
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   7200
            TabIndex        =   20
            Text            =   "120"
            Top             =   440
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   5400
            TabIndex        =   18
            Text            =   "0"
            Top             =   440
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   3720
            TabIndex        =   16
            Text            =   "1"
            Top             =   440
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   14
            Text            =   "220"
            Top             =   440
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Power Consum"
            Height          =   675
            Index           =   46
            Left            =   1320
            TabIndex        =   202
            Top             =   1515
            Width           =   945
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Correction coefficient"
            Height          =   195
            Index           =   45
            Left            =   7680
            TabIndex        =   192
            Top             =   195
            Width           =   1545
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Correction coefficient"
            Height          =   195
            Index           =   44
            Left            =   5880
            TabIndex        =   191
            Top             =   195
            Width           =   1545
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "% correction factor"
            Height          =   195
            Index           =   43
            Left            =   3720
            TabIndex        =   190
            Top             =   195
            Width           =   1410
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "% correction factor"
            Height          =   195
            Index           =   42
            Left            =   2160
            TabIndex        =   189
            Top             =   195
            Width           =   1410
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPhase U:"
            Height          =   195
            Index           =   24
            Left            =   1320
            TabIndex        =   131
            Top             =   1245
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ic:"
            Height          =   195
            Index           =   23
            Left            =   3360
            TabIndex        =   130
            Top             =   1245
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ang.UI:"
            Height          =   195
            Index           =   22
            Left            =   4800
            TabIndex        =   129
            Top             =   1245
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ang.UU:"
            Height          =   195
            Index           =   21
            Left            =   6555
            TabIndex        =   128
            Top             =   1245
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fre:"
            Height          =   195
            Index           =   20
            Left            =   8400
            TabIndex        =   127
            Top             =   1245
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BPhase U:"
            Height          =   195
            Index           =   19
            Left            =   1320
            TabIndex        =   121
            Top             =   885
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ib:"
            Height          =   195
            Index           =   18
            Left            =   3360
            TabIndex        =   120
            Top             =   885
            Width           =   210
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ang.UI:"
            Height          =   195
            Index           =   17
            Left            =   4800
            TabIndex        =   119
            Top             =   885
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ang.UU:"
            Height          =   195
            Index           =   16
            Left            =   6555
            TabIndex        =   118
            Top             =   885
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fre:"
            Height          =   195
            Index           =   15
            Left            =   8400
            TabIndex        =   117
            Top             =   885
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mode:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fre:"
            Height          =   195
            Index           =   4
            Left            =   8400
            TabIndex        =   21
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ang.UU:"
            Height          =   195
            Index           =   3
            Left            =   6555
            TabIndex        =   19
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ang.UI:"
            Height          =   195
            Index           =   2
            Left            =   4800
            TabIndex        =   17
            Top             =   480
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ia:"
            Height          =   195
            Index           =   1
            Left            =   3360
            TabIndex        =   15
            Top             =   480
            Width           =   210
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "APhase U:"
            Height          =   195
            Index           =   0
            Left            =   1320
            TabIndex        =   13
            Top             =   480
            Width           =   750
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sent"
         Height          =   3015
         Left            =   -74520
         TabIndex        =   7
         Top             =   1680
         Width           =   9495
         Begin VB.TextBox txtSend 
            Height          =   1455
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1440
            Width           =   9135
         End
         Begin VB.CommandButton cmdClearSend 
            Caption         =   "Clear"
            Height          =   360
            Left            =   8160
            TabIndex        =   9
            Top             =   480
            Width           =   1110
         End
         Begin VB.TextBox txtInPut 
            Height          =   975
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   360
            Width           =   7695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Receive"
         Height          =   2055
         Left            =   -74520
         TabIndex        =   4
         Top             =   5400
         Width           =   9495
         Begin VB.TextBox txtGet 
            Height          =   1575
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   360
            Width           =   7935
         End
         Begin VB.CommandButton cmdClearGet 
            Caption         =   "Clear"
            Height          =   360
            Left            =   8280
            TabIndex        =   5
            Top             =   360
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   360
         Left            =   -67800
         TabIndex        =   3
         Top             =   4875
         Width           =   1215
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   360
         Left            =   -66360
         TabIndex        =   2
         Top             =   4875
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "The serial port number is automatically added when data is sent"
         Height          =   255
         Left            =   -74400
         TabIndex        =   1
         Top             =   4920
         Width           =   6495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "readings："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   61
         Left            =   -74760
         TabIndex        =   239
         Top             =   6120
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Send bytes must be separated by Spaces."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74520
         TabIndex        =   11
         Top             =   7560
         Width           =   4575
      End
   End
   Begin VB.Label labInfo 
      AutoSize        =   -1  'True
      Caption         =   "Result: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   55
      Top             =   7920
      Width           =   630
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private binShowMsg As Boolean
Private intTag     As Integer  '通讯功能标记
Const strHead      As String = "59 43 53 53 2D 31 30 31 " '头标志 YCSS-101
Private intRow     As Integer

Private GenyYCSS   As New YCSS
'Private mStdMeter As New StandardMeterInterface
'Private mStdMeterData As New StandardMeterData

Private LogFullName As String

Dim bolTimeStop   As Boolean
Dim MeterNu As Integer
Dim Constant() As String
Dim MeterIP() As String
Dim MeterPort() As Integer
    

Private Sub IniContrcols()
    Dim i      As Integer
    Dim strTmp As String
    '-------------------------------------------
    cmbSource.AddItem "0-DA"
    cmbSource.AddItem "1-DQ"
    cmbSource.AddItem "2-DS"
    
    cmbSource.AddItem "0-4WA"
    cmbSource.AddItem "1-4WQ"
    cmbSource.AddItem "2-4WS"
    cmbSource.AddItem "3-3WA"
    cmbSource.AddItem "4-3WQ"
    cmbSource.AddItem "5-3WS"
    '-------------------------------------------
    cmbSource2.AddItem "0-Fundamental"
    cmbSource2.AddItem "1-Even"
    cmbSource2.AddItem "2-Odd"
    cmbSource2.AddItem "3-Sub"
    '-------------------------------------------
    cmbComModel.AddItem "0-DA"
    cmbComModel.AddItem "1-DQ"
    cmbComModel.AddItem "2-DS"
    
    cmbComModel.AddItem "0-4WA"
    cmbComModel.AddItem "1-4WQ"
    cmbComModel.AddItem "2-4WS"
    cmbComModel.AddItem "3-3WA"
    cmbComModel.AddItem "4-3WQ"
    cmbComModel.AddItem "5-3WS"
    '--------------------------------------------
    cmbErrMode.AddItem "0"
    cmbErrMode.AddItem "1"
    
    cmbWay.AddItem "0-Normal Err%"
    cmbWay.AddItem "1-DayPrec"
    cmbWay.AddItem "2-DemandPeriod"
    cmbWay.AddItem "3-StartTest"
    cmbWay.AddItem "4-Creeping"
    cmbWay.AddItem "5-DailTest"
    cmbWay.AddItem "6-PulseSensor"

    For i = 1 To 21
        cmbXB.AddItem i & " Harmonic"
    Next
    
    cmbPort.AddItem "1234"
    
    For i = 1 To 255
       cmbIP.AddItem "192.168.1." & i
       cmbPort.AddItem 1000 + i
    Next i
    
    Call InimsfgXB
    Call InimsfgComTest
    Call InmsfgComStandard
    
    chkComReadStandard.Value = GetIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "chkComReadStandard", 0)
    txtTestTime.Text = GetIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtTestTime", 1)
    txtWaitTIme.Text = GetIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtWaitTIme", 2)
    txtReadError.Text = GetIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtReadError", "0.2")
End Sub

Private Sub InmsfgComStandard()
    Dim i      As Integer
    Dim strTmp As String

    With msfgComStandard
        .Clear
        .Cols = 4
        .Rows = 6
        .ColWidth(0) = 500
        
        .Row = 0
        .Col = 1
        .Text = "A"
        .CellForeColor = &HFFFF&
        .Col = 2
        .Text = "B"
        .CellForeColor = &HFF00&
        .Col = 3
        .Text = "C"
        .CellForeColor = &HFF&

        .Col = 0
        .Row = 1
        .Text = "U:"
        .CellForeColor = &HC0C000
        .Row = 2
        .Text = "I:"
        .CellForeColor = &HC0C000
        .Row = 3
        .Text = "φ:"
        .CellForeColor = &HC0C000
        .Row = 4
        .Text = "P:"
        .CellForeColor = &HC0C000
        .Row = 5
        .Text = "Q:"
        .CellForeColor = &HC0C000
    End With

End Sub

Private Sub InimsfgComTest()
    Dim i      As Integer
    Dim strTmp As String

    With msfgComTest
        .Clear
        .Cols = 9
        .Rows = GetIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "Rows", 1)
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = "SN"
        .TextMatrix(0, 1) = "Mode"
        .ColWidth(1) = 500
        .TextMatrix(0, 2) = "U"
        .TextMatrix(0, 3) = "I"
        .TextMatrix(0, 4) = "Ang.UI"
        .TextMatrix(0, 5) = "Ang.UU"
        .TextMatrix(0, 6) = "Fre"
        .TextMatrix(0, 7) = "Loop"
        .ColWidth(7) = 500
        .TextMatrix(0, 8) = "Ref.Co"
        .ColWidth(8) = 1200

        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
            strTmp = GetIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "R" & i)

            If strTmp <> "" Then
                .TextMatrix(i, 1) = GetItem(strTmp, ",", 0)
                .TextMatrix(i, 2) = GetItem(strTmp, ",", 1)
                .TextMatrix(i, 3) = GetItem(strTmp, ",", 2)
                .TextMatrix(i, 4) = GetItem(strTmp, ",", 3)
                .TextMatrix(i, 5) = GetItem(strTmp, ",", 4)
                .TextMatrix(i, 6) = GetItem(strTmp, ",", 5)
                .TextMatrix(i, 7) = GetItem(strTmp, ",", 6)
                .TextMatrix(i, 8) = GetItem(strTmp, ",", 7)
            End If

        Next i

    End With

End Sub

Private Sub InimsfgXB()
    Dim i      As Integer
    Dim iRow   As Integer
    Dim strTmp As String

    With msfgXB
        .Clear
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = "SN"
        .TextMatrix(0, 1) = "Amp."
        .TextMatrix(0, 2) = "Angle"
        .TextMatrix(0, 3) = "Order"

        '第一行数据固定
        .TextMatrix(1, 0) = 1
        .TextMatrix(1, 1) = 100
        .TextMatrix(1, 2) = 0
        .TextMatrix(1, 3) = 1

        '从第二行开始设置
        iRow = 2

        For i = 2 To .Rows - 1
            .TextMatrix(i, 0) = i
            strTmp = GetIniInfo(ProgramPath & "LanComTest.ini", "XieBo", i, "0,0,0")

            If Val(strTmp) <> 0 Then
                .TextMatrix(iRow, 1) = GetItem(strTmp, ",", 0)
                .TextMatrix(iRow, 2) = GetItem(strTmp, ",", 1)
                .TextMatrix(iRow, 3) = GetItem(strTmp, ",", 2)
                iRow = iRow + 1
            End If

        Next i

    End With

    With msfgReadXB
        .Clear
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = "SN"
        .TextMatrix(0, 1) = "Amp."
        .TextMatrix(0, 2) = "Angle."
        .TextMatrix(0, 3) = "Order"
    End With

End Sub

Private Function LinkIp(strIP As String, strPort As String, Optional IsTest As Boolean = False) As Boolean

    Dim lngX As Long

    '-----------------------------------------------------------------------
    With Winsock1
        If IsTest Then .Close
        If (.State <> sckConnected And Not IsStop) Or (.RemoteHostIP <> strIP And .RemotePort <> strPort) Then
            .Close      '先关闭连接
            .Connect strIP, strPort    '网卡连接
            lngX = GetTickCount

            Do
                DoEvents
                Sleep 10
            Loop Until .State = sckConnected Or IsStop Or GetTickCount - lngX > 5000 '判断连接超时

        End If

    End With

    LinkIp = (Winsock1.State = sckConnected)  '返回连接状态
End Function

Private Sub SetIpPort()
    txtIP(0).Text = cmbIP.Text
    txtIP(1).Text = cmbIP.Text
    
    txtPort(0).Text = cmbPort.Text
    txtPort(1).Text = cmbPort.Text
End Sub


Private Sub cmbIP_Change()
    Call SetIpPort
End Sub

Private Sub cmbIP_Click()
    Call SetIpPort
End Sub

Private Sub cmbPort_Change()
    Call SetIpPort
End Sub

Private Sub cmbPort_Click()
    Call SetIpPort
End Sub

Private Sub cmbSource_Change()
    If Val(cmbSource.Text) = 3 Then
       Text1(3).Text = 0
       Text1(13).Text = 120
       Text1(18).Text = 240
    Else
       Text1(3).Text = 120
       Text1(13).Text = 120
       Text1(18).Text = 120
    End If
End Sub

Private Sub cmdClearGet_Click()
    txtGet.Text = ""  '清空接收数据显示框
End Sub

Private Sub cmdClearSend_Click()
    txtSend.Text = ""  '清空发送数据显示框
End Sub

Private Sub cmdComStart_Click()
    On Error GoTo ToExit '打开错误陷阱
    '------------------------------------------------
    Dim i As Integer
    Dim strTmp As String

    Label2(34).Caption = 0
    Label2(35).Caption = 0
    Label2(36).Caption = 0

    Label2(47).Caption = 0
    Label2(49).Caption = 0
    Label2(51).Caption = 0

    cmdComStart.Enabled = False

    MeterNu = Val(GetIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtMeterNu"))
    ReDim MeterIP(1 To MeterNu) As String
    ReDim MeterPort(1 To MeterNu) As Integer
    ReDim Constant(1 To MeterNu) As String

    For i = 1 To MeterNu
        strTmp = GetIniInfo(ProgramPath & "LanComTest.ini", "MeterIP", "R" & i)

        If strTmp <> "" Then
            Constant(i) = GetItem(strTmp, ",", 0)
            MeterIP(i) = GetItem(strTmp, ",", 1)
            MeterPort(i) = GetItem(strTmp, ",", 2)
        End If
    Next i

    bolTimeStop = False
    Label2(40) = "00:00:00"
    Timer1.Enabled = True
    blnStop = False

'    mStdMeter.ComPort = 1
'    mStdMeter.Name = "SZ-03A-K6D"

    Do
        Label2(72).Caption = Val(Label2(72).Caption) + 1

        For i = 1 To msfgComTest.Rows - 1
            If msfgComTest.TextMatrix(i, 2) <> "" Then
                RunTest (i)
            End If
            If blnStop Or bolTimeStop Then GoTo EndP
        Next i
    Loop Until blnStop Or bolTimeStop

EndP:
    Label2(77).Caption = "Down..."
    Call GenyYCSS.Source.Reset(txtIP(1).Text, txtPort(1).Text)
    Call InmsfgComStandard
    Timer1.Enabled = False
    cmdComStart.Enabled = True
    Label2(77).Caption = "Ready to test..."
    '------------------------------------------------
    Exit Sub
    '----------------
ToExit:
End Sub

Private Sub RunTest(iRow As Integer)
    On Error GoTo EndP

    Dim strBack As String
    Dim i As Integer
    Dim bolTmp As Boolean

    With msfgComTest

        For i = 1 To .Cols - 1
            .Row = iRow
            .Col = i
            .CellBackColor = &HFFC0C0
        Next i
        '控制源---------------------------------------------------------------------------------------------------------------------------
        Label2(77).Caption = "Output..."
        If Val(.TextMatrix(iRow, 1)) = 3 Then
            strBack = GenyYCSS.Source.SetValue(txtIP(1).Text, txtPort(1).Text, .TextMatrix(iRow, 1), .TextMatrix(iRow, 2), .TextMatrix(iRow, 3), .TextMatrix(iRow, 4), 0, .TextMatrix(iRow, 6), .TextMatrix(iRow, 2), .TextMatrix(iRow, 3), .TextMatrix(iRow, 4), 120, .TextMatrix(iRow, 6), .TextMatrix(iRow, 2), .TextMatrix(iRow, 3), .TextMatrix(iRow, 4), 240, .TextMatrix(iRow, 6))
        Else
            strBack = GenyYCSS.Source.SetValue(txtIP(1).Text, txtPort(1).Text, .TextMatrix(iRow, 1), .TextMatrix(iRow, 2), .TextMatrix(iRow, 3), .TextMatrix(iRow, 4), .TextMatrix(iRow, 5), .TextMatrix(iRow, 6), .TextMatrix(iRow, 2), .TextMatrix(iRow, 3), .TextMatrix(iRow, 4), .TextMatrix(iRow, 5), .TextMatrix(iRow, 6), .TextMatrix(iRow, 2), .TextMatrix(iRow, 3), .TextMatrix(iRow, 4), .TextMatrix(iRow, 5), .TextMatrix(iRow, 6))
        End If

        labInfo = "Result：" & IIf(strBack = "ok", "Success", "Fail")
        Call ShowComCount(strBack = "ok", "Source Control")

        If blnStop Or bolTimeStop Then Exit Sub
        '设置误差---------------------------------------------------------------------------------------------------------------------------
        If Check6(4).Value = 1 Then
            Label2(77).Caption = "Sending..."
            Call SetError(.TextMatrix(iRow, 7), .TextMatrix(iRow, 8))
        End If

        If blnStop Or bolTimeStop Then Exit Sub

        Label2(77).Caption = "Waiting..."
        Delay Val(txtWaitTIme.Text)

        If blnStop Or bolTimeStop Then Exit Sub

        If Check6(4).Value = 1 Then
            Label2(77).Caption = "Readings..."
            ReadError (Val(Text2.Text))
        End If

        If blnStop Or bolTimeStop Then Exit Sub

        If chkComReadStandard.Value = 1 Then
            Label2(77).Caption = "Readings from Ref..."
            If Check4.Value = 1 Then
                '通过源读取标准表
                strBack = GenyYCSS.StandardMeter.ReadValue(txtIP(1).Text, txtPort(1).Text)
                labInfo = "Result: " & IIf(strBack = "ok", "Success", "Fail")
            Else
                '直接读取标准表
'                bolTmp = mStdMeter.ReadData(mStdMeterData, False)
'                labInfo = "读取标准表：" & IIf(bolTmp, "Success", "Fail")
'                strBack = "ok"
            End If
            Call ShowComCount(strBack = "ok", "Readings of Ref.")
            If blnStop Or bolTimeStop Then Exit Sub
            If strBack = "ok" Then
                Call DisplayStandardMeterData(GenyYCSS.StandardMeter.StandardMeterData)
                '比对读取到的数据
                Call CompareData(GenyYCSS.StandardMeter.StandardMeterData, .TextMatrix(iRow, 2), .TextMatrix(iRow, 3), .TextMatrix(iRow, 4), .TextMatrix(iRow, 8))
            End If

            'Delay Val(txtWaitTIme.Text)
        End If

        If blnStop Or bolTimeStop Then Exit Sub

        '485通讯---------------------------------------------------------------------------------------------------------------------------
        If Check6(5).Value = 1 Then
            Label2(77).Caption = "485 comm. test..."
            Call Com485Test
            Label8(6) = ""
        End If

        If blnStop Or bolTimeStop Then Exit Sub

        '负荷开关---------------------------------------------------------------------------------------------------------------------------
        If Check6(6).Value = 1 Then
            Label2(77).Caption = "Load switch..."
            Call LoadSwitchTest
        End If

        If blnStop Or bolTimeStop Then Exit Sub

        For i = 1 To .Cols - 1
            .Row = iRow
            .Col = i
            .CellBackColor = &HFFFFFF
        Next i
    End With

    Exit Sub

EndP:
    MsgBox Err.Description
End Sub

Private Sub LoadSwitchTest()
    Dim strBack        As String
    Dim strSendValue() As String
    Dim strBackValue() As String
    Dim i              As Integer
    Dim k As Integer
    Dim LSH As LoadSwitch

    '开-------------------------------------------------------------------------------------------------------------------------
    LSH = GenyYCSS.MultiErrorCount.IniLoadSwitch
    For i = 1 To MeterNu
        If blnStop Or bolTimeStop Then Exit Sub
        strBack = GenyYCSS.MultiErrorCount.SetLoadSwitch(MeterIP(i), CStr(MeterPort(i)), LSH)

        If UCase$(strBack) <> "OK" Then
            Call ShowComCount(False, "LoadSwithc{On}Pos" & i)
            Label2(75).Caption = Label2(75).Caption + 1
        End If
        Delay Val(txtReadError.Text)
    Next i
   ' Delay MeterNu * 0.2
    '读开-------------------------------------------------------------------------------------------------------------------------
    For i = 1 To MeterNu
        If blnStop Or bolTimeStop Then Exit Sub
        strBack = GenyYCSS.MultiErrorCount.ReadLoadSwitch(MeterIP(i), CStr(MeterPort(i)), LSH)

        Call ShowComCount((strBack <> "fail"), "Readings of LoadSwithc{on}position" & i)

        If UCase$(strBack) <> "OK" Then
            Label2(75).Caption = Label2(75).Caption + 1
        ElseIf ((i - 1) Mod 4 + 1 = 1 And LSH.MeterNo1 <> 1) Then
            Label2(75).Caption = Label2(75).Caption + 1
        ElseIf ((i - 1) Mod 4 + 1 = 2 And LSH.MeterNo2 <> 1) Then
            Label2(75).Caption = Label2(75).Caption + 1
        ElseIf ((i - 1) Mod 4 + 1 = 3 And LSH.MeterNo3 <> 1) Then
            Label2(75).Caption = Label2(75).Caption + 1
        ElseIf ((i - 1) Mod 4 + 1 = 4 And LSH.MeterNo4 <> 1) Then
            Label2(75).Caption = Label2(75).Caption + 1
        End If
        Delay Val(txtReadError.Text)
    Next i
    '关-------------------------------------------------------------------------------------------------------------------------
    For i = 1 To MeterNu
        If blnStop Or bolTimeStop Then Exit Sub
        LSH.MeterNo1 = 0
        LSH.MeterNo2 = 0
        LSH.MeterNo3 = 0
        LSH.MeterNo4 = 0
        strBack = GenyYCSS.MultiErrorCount.SetLoadSwitch(MeterIP(i), CStr(MeterPort(i)), LSH)

        If UCase$(strBack) <> "OK" Then
            Call ShowComCount(False, "LoadSwitch{Off}Position " & i)
            Label2(75).Caption = Label2(75).Caption + 1
        End If
        Delay Val(txtReadError.Text)
    Next i
   ' Delay MeterNu * 0.2
    '读关-------------------------------------------------------------------------------------------------------------------------
    For i = 1 To MeterNu
        If blnStop Or bolTimeStop Then Exit Sub

        strBack = GenyYCSS.MultiErrorCount.ReadLoadSwitch(MeterIP(i), CStr(MeterPort(i)), LSH)
        Call ShowComCount((strBack <> "fail"), "Readings of LoadSwithc{off}position " & i)

        If UCase$(strBack) <> "OK" Then
            Label2(75).Caption = Label2(75).Caption + 1
        ElseIf ((i - 1) Mod 4 + 1 = 1 And LSH.MeterNo1 <> 0) Then
            Label2(75).Caption = Label2(75).Caption + 1
        ElseIf ((i - 1) Mod 4 + 1 = 2 And LSH.MeterNo2 <> 0) Then
            Label2(75).Caption = Label2(75).Caption + 1
        ElseIf ((i - 1) Mod 4 + 1 = 3 And LSH.MeterNo3 <> 0) Then
            Label2(75).Caption = Label2(75).Caption + 1
        ElseIf ((i - 1) Mod 4 + 1 = 4 And LSH.MeterNo4 <> 0) Then
            Label2(75).Caption = Label2(75).Caption + 1
        End If
        Delay Val(txtReadError.Text)
    Next i
End Sub

Private Sub Com485Test()
    Dim strBack        As String
    Dim strSendValue() As String
    Dim strBackValue() As String
    Dim i              As Integer
    Dim k As Integer

    If Text3.Text <> "" Then
        strSendValue = Split(Trim$(Text3), " ")
        For i = 1 To MeterNu
            If blnStop Or bolTimeStop Then Exit Sub

            Label8(6) = "Position: " & i
            Text4.Text = ""

            strBack = GenyYCSS.MultiErrorCount.SetComBaudRate(MeterIP(i), CStr(MeterPort(i)), ((i - 1) Mod 4) + 1, 0, 0, "2400,E,8,1")

            Call ShowComCount((strBack <> "fail"), "Baud rate on " & i)

            If UCase$(strBack) = "OK" Then
                Delay Val(txtReadError.Text)
                strBack = GenyYCSS.MultiErrorCount.SendComData(MeterIP(i), CStr(MeterPort(i)), ((i - 1) Mod 4) + 1, strSendValue)
                Call ShowComCount((strBack <> "fail"), "485 sent on " & i)
                If blnStop Or bolTimeStop Then Exit Sub
                Delay (2)
                If blnStop Or bolTimeStop Then Exit Sub
                
                If UCase$(strBack) = "OK" Then
                    strBack = GenyYCSS.MultiErrorCount.ReadComData(MeterIP(i), CStr(MeterPort(i)), ((i - 1) Mod 4) + 1, strBackValue)
                    Call ShowComCount((strBack <> "fail"), "485 read on " & i)
                    If UCase$(strBack) = "OK" Then
                        For k = 0 To UBound(strBackValue)
                            strBack = strBack & " " & strBackValue(k)
                        Next k

                        Text4.Text = strBack
                    Else
                        Label2(71).Caption = Label2(71).Caption + 1
                    End If
                Else
                    Label2(71).Caption = Label2(71).Caption + 1
                End If

            ElseIf UCase$(strBack) = "FAIL" Then
                Label2(71).Caption = Label2(71).Caption + 1
            End If
            Delay Val(txtReadError.Text)
        Next i
    End If
End Sub

Private Sub SetError(DoubleLoop As Integer, stdMeterConstant As String)
    Dim i As Integer
    Dim strBack As String
    Dim EHP As ErrorHeadParam
    Dim EMP(1 To 8) As ErrorMeterParam

    For i = 1 To MeterNu
        If blnStop Or bolTimeStop Then Exit Sub
        EHP = GenyYCSS.MultiErrorCount.IniErrorHeadParam
        'EHP.DoubleLoop = DoubleLoop-1
        EHP.StdMeterConstand = stdMeterConstant


        EMP(((i - 1) Mod 4) * 2 + 1) = GenyYCSS.MultiErrorCount.IniErrorMeterParam
        EMP(((i - 1) Mod 4) * 2 + 1).Constant = Constant(i)

        strBack = GenyYCSS.MultiErrorCount.SetValue(MeterIP(i), CStr(MeterPort(i)), EHP, EMP())
        If UCase$(strBack) <> "OK" Then Label2(69).Caption = Label2(69).Caption + 1
        Call ShowComCount((strBack <> "fail"), "set Normal Err on " & i)
        If blnStop Or bolTimeStop Then Exit Sub
        Delay Val(txtReadError.Text)
    Next i
End Sub

Private Sub ReadError(WaitTime As Integer)
    Dim lngTime As Long
    Dim strBack As String
    Dim lngN As Long
    Dim strDate As String
    Dim i As Integer
    Dim BackData() As String
    Dim tmpErr As String

    If Check5.Value = 0 Then
        lngN = 0
        strDate = Date
        lngTime = Timer

        Do
            strBack = GenyYCSS.ErrorCount.ReadValue(txtIP(1).Text, txtPort(1).Text, 1)
            Delay Val(txtReadError.Text)
            lngN = lngN + 1
            Label2(41).Caption = "Readings of " & lngN & " : " & Left$(strBack, 10)
            Call ShowComCount(strBack <> "fail", "Reading from Error Counter")
        Loop Until (Timer - lngTime >= WaitTime) Or (Date > strDate) Or IsStop
    Else
        lngN = 0
        strDate = Date
        lngTime = Timer

        Do
            For i = 1 To MeterNu
                If blnStop Or bolTimeStop Then Exit Sub
                strBack = GenyYCSS.MultiErrorCount.ReadValue(MeterIP(i), CStr(MeterPort(i)), BackData())
                If strBack <> "fail" Then
                    tmpErr = BackData(((i - 1) Mod 4) * 2 + 1, 1)
                    Label2(69).Caption = Label2(69).Caption + 1
                End If

                Label2(41).Caption = "Read on  " & i & "Err%: " & tmpErr
                Call ShowComCount((strBack <> "fail"), "Readings of Error Counter")
                If blnStop Or bolTimeStop Then Exit Sub
                Delay Val(txtReadError.Text)
            Next i
        Loop Until (Timer - lngTime >= WaitTime) Or (Date > strDate) Or IsStop
    End If
End Sub

Private Sub CompareData(stdData As clsStandardMeterData, _
                        Voltage As String, _
                        Current As String, _
                        Phase As String, stdConstant As String)

'    If Me.Check4.Value = 1 Then
'        'A相比对
'        If Check6(0).Value = 1 Then
'            If Abs(Val(Voltage) - stdData.Voltage_A) > 0.5 Then
'                Label2(47).Caption = Label2(47).Caption + 1
'                Call LogWrite(Now & "->" & "Ua：STD=" & Voltage & vbTab & "Readings=" & stdData.Voltage_A, LogFullName, True)
'            End If
'
'            If Abs(Val(Current) - stdData.Current_A) > 0.5 Then
'                Label2(47).Caption = Label2(47).Caption + 1
'                Call LogWrite(Now & "->" & "Ia：STD=" & Current & vbTab & "Readings=" & stdData.Current_A, LogFullName, True)
'            End If
'
'            If Abs(Val(Phase) - stdData.Phase_A) < 1 Or Abs(Val(Phase) - stdData.Phase_A) > 359 Then
'            Else
'                Label2(47).Caption = Label2(47).Caption + 1
'                Call LogWrite(Now & "->" & "φa：STD=" & Phase & vbTab & "Readings=" & stdData.Phase_A, LogFullName, True)
'            End If
'        End If
'
'        'B相比对
'        If Check6(1).Value = 1 Then
'            If Abs(Val(Voltage) - stdData.Voltage_B) > 0.5 Then
'                Label2(49).Caption = Label2(49).Caption + 1
'                Call LogWrite(Now & "->" & "Ub：STD=" & Voltage & vbTab & "Readings=" & stdData.Voltage_B, LogFullName, True)
'            End If
'            If Abs(Val(Current) - stdData.Current_B) > 0.5 Then
'                Label2(49).Caption = Label2(49).Caption + 1
'                Call LogWrite(Now & "->" & "Ib：STD=" & Current & vbTab & "Readings=" & stdData.Current_B, LogFullName, True)
'            End If
'            If Abs(Val(Phase) - stdData.Phase_B) < 1 Or Abs(Val(Phase) - stdData.Phase_B) > 359 Then
'            Else
'                Label2(49).Caption = Label2(49).Caption + 1
'                Call LogWrite(Now & "->" & "φb：STD=" & Phase & vbTab & "Readings=" & stdData.Phase_B, LogFullName, True)
'            End If
'        End If
'
'        'C相比对
'        If Check6(2).Value = 1 Then
'            If Abs(Val(Voltage) - stdData.Voltage_C) > 0.5 Then
'                Label2(51).Caption = Label2(51).Caption + 1
'                Call LogWrite(Now & "->" & "Uc：STD=" & Voltage & vbTab & "Readings=" & stdData.Voltage_C, LogFullName, True)
'            End If
'            If Abs(Val(Current) - stdData.Current_C) > 0.5 Then
'                Label2(51).Caption = Label2(51).Caption + 1
'                Call LogWrite(Now & "->" & "Ic：STD=" & Current & vbTab & "Readings=" & stdData.Current_C, LogFullName, True)
'            End If
'            If Abs(Val(Phase) - stdData.Phase_C) < 1 Or Abs(Val(Phase) - stdData.Phase_C) > 359 Then
'            Else
'                Label2(51).Caption = Label2(51).Caption + 1
'                Call LogWrite(Now & "->" & "φc：STD=" & Phase & vbTab & "Readings=" & stdData.Phase_C, LogFullName, True)
'            End If
'        End If
'
'        '常数比对
'        If Check6(2).Value = 1 Then
'            If Abs(Val(stdConstant) - stdData.Constant) > 0.5 Then Label2(66).Caption = Label2(66).Caption + 1
'            Call LogWrite(Now & "->" & "Ref. Co: STD=" & stdConstant & vbTab & "Readings=" & stdData.Constant, LogFullName, True)
'        End If
'    Else
'        'A相比对
'        If Check6(0).Value = 1 Then
'            If Abs(Val(Voltage) - mStdMeterData.Ua) > 0.5 Then Label2(47).Caption = Label2(47).Caption + 1
'            If Abs(Val(Current) - mStdMeterData.Ia) > 0.5 Then Label2(47).Caption = Label2(47).Caption + 1
'            If Abs(Val(Phase) - mStdMeterData.PhaseA) < 1 Or Abs(Val(Phase) - mStdMeterData.PhaseA) > 359 Then
'            Else
'                Label2(47).Caption = Label2(47).Caption + 1
'            End If
'        End If
'
'        'B相比对
'        If Check6(1).Value = 1 Then
'            If Abs(Val(Voltage) - mStdMeterData.Ub) > 0.5 Then Label2(49).Caption = Label2(49).Caption + 1
'            If Abs(Val(Current) - mStdMeterData.Ib) > 0.5 Then Label2(49).Caption = Label2(49).Caption + 1
'            If Abs(Val(Phase) - mStdMeterData.PhaseB) < 1 Or Abs(Val(Phase) - mStdMeterData.PhaseB) > 359 Then
'            Else
'                Label2(49).Caption = Label2(49).Caption + 1
'            End If
'        End If
'
'        'C相比对
'        If Check6(2).Value = 1 Then
'            If Abs(Val(Voltage) - mStdMeterData.Uc) > 0.5 Then Label2(51).Caption = Label2(51).Caption + 1
'            If Abs(Val(Current) - mStdMeterData.Ic) > 0.5 Then Label2(51).Caption = Label2(51).Caption + 1
'            If Abs(Val(Phase) - mStdMeterData.PhaseC) < 1 Or Abs(Val(Phase) - mStdMeterData.PhaseC) > 359 Then
'            Else
'                Label2(51).Caption = Label2(51).Caption + 1
'            End If
'        End If
'
'        '常数比对
'        '        If Check6(2).Value = 1 Then
'        '            If Abs(Val(stdConstant) - mStdMeterData.Constant) > 0.5 Then Label2(66).Caption = Label2(66).Caption + 1
'        '        End If
'    End If

End Sub

Private Sub ShowComCount(blnOK As Boolean, errInfo As String)
      On Error Resume Next
      
      Label2(34).Caption = Val(Label2(34).Caption) + 1
      
      If Not blnOK Then
        Label2(35).Caption = Val(Label2(35).Caption) + 1
        Call LogWrite(Now & "->" & "Comm:" & errInfo & "Fail", LogFullName, True)
      End If
      
      Label2(36).Caption = Format((Val(Label2(34).Caption) - Val(Label2(35).Caption)) * 100 / Val(Label2(34).Caption), "0.00") & "%"
End Sub

Private Sub cmdComStop_Click()
    Timer1.Enabled = False
    blnStop = True
    
End Sub

Private Sub cmdLink_Click(Index As Integer)
    IsStop = False     '是否停止运行标志
    cmdLink(Index).Enabled = False
    '========================================================================
    '保存设置信息到注册表
    Call SaveSetting("LanTest", "SetData", "IP", txtIP(Index).Text)
    Call SaveSetting("LanTest", "SetData", "Port", txtPort(Index).Text)
    Call SaveSetting("LanTest", "SetData", "Comm", txtComm(Index).Text)

    '=========================================================================
    If LinkIp(txtIP(Index).Text, txtPort(Index).Text, True) = True Then
        MsgBox "连接Success！", vbInformation, "Sys. Notify..."
    Else
        MsgBox "连接Fail！", vbCritical, "Sys. Notify..."
    End If

    cmdLink(Index).Enabled = True
End Sub

Private Sub cmdSend_Click()

    Dim strSendData   As String

    Dim bytSendData() As Byte

    binShowMsg = True '提示显示错误对话框
    cmdSend.Enabled = False '锁定按钮不能操作
    IsStop = False    '是否停止运行标志

    If txtSend.Text <> "" Then    '发送数据不能为空
        If LinkIp(txtIP(0).Text, txtPort(0).Text) = True Then
            '=================连接Success后发送数据==========================
            strSendData = txtSend.Text

            If Check1.Value = 1 Then
                '自动增加串口号数据
                strSendData = "00 " & Format$(txtComm(0).Text, "00") & "" & strSendData
            End If

            strSendData = Trim$(strSendData)
            '调试打印出发送字节数据
            Debug.Print "发送数据为：" & strSendData
            '将字符串转换为数组
            Call getSendByte(bytSendData, strSendData)
            Winsock1.SendData bytSendData()
        End If

    Else
        MsgBox "sending should not be empty!", vbCritical, "Sys. Notify..."
    End If

    cmdSend.Enabled = True '释放按钮
End Sub

Private Sub cmdStop_Click()
    IsStop = True
End Sub

Private Sub Command10_Click()
    Dim i      As Integer
    Dim strTmp As String

    With msfgXB

        For i = 1 To .Rows - 1
            strTmp = .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)

            Call SaveIniInfo(ProgramPath & "LanComTest.ini", "XieBo", i, strTmp)
        Next

    End With

    Call InimsfgXB
End Sub

Private Sub Command11_Click()
    '读取谐波数据
    Dim strAmplitudeA() As String
    Dim strPhaseA()     As String
    Dim strTimesA()     As String
    Dim strAmplitudeB() As String
    Dim strPhaseB()     As String
    Dim strTimesB()     As String
    Dim strAmplitudeC() As String
    Dim strPhaseC()     As String
    Dim strTimesC()     As String
    Dim strBack               As String
    Dim i                     As Integer

    Command11.Enabled = False

    Do
        strBack = GenyYCSS.Harmonics.ReadValue(txtIP(1).Text, txtPort(1).Text, IIf(optXB(0).Value = True, "U", "I"), strAmplitudeA(), strPhaseA(), strTimesA(), strAmplitudeB(), strPhaseB(), strTimesB(), strAmplitudeC(), strPhaseC(), strTimesC())

        If UCase(strBack) = "OK" Then
            labInfo = "Result: Success！"

            For i = 0 To UBound(strAmplitudeA)

                If i + 1 > msfgXB.Rows - 1 Then Exit For

                msfgReadXB.TextMatrix(i + 1, 0) = i + 1
                msfgReadXB.TextMatrix(i + 1, 1) = strAmplitudeA(i)
                msfgReadXB.TextMatrix(i + 1, 2) = strPhaseA(i)
                msfgReadXB.TextMatrix(i + 1, 3) = strTimesA(i)
            Next i

        ElseIf strBack = "FAIL" Then
            labInfo = "Result: Fail！"
        End If

        Delay (Text1(7).Text / 1000)
    Loop While (chkAutoReadXB.Value = 1 And Not IsStop)

    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '谐波输出
    Dim strBack               As String
    Dim strAmplitude(0 To 20) As String
    Dim strPhase(0 To 20)     As String
    Dim strTimes(0 To 20)     As String

    Command12.Enabled = False

    strBack = GenyYCSS.Harmonics.Reset(txtIP(1).Text, txtPort(1).Text, IIf(optXB(0).Value = True, "U", "I"))

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    '读取装置状态

    Dim strBack As String

    Command13.Enabled = False

    strBack = GenyYCSS.Other.ReadState(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) <> "FAIL" Then
        labInfo = "Result: Success！"
        Label7.Caption = "Readings: " & strBack
    Else
        labInfo = "Result: Fail！"
        Label7.Caption = "Readings: "
    End If

    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    '误差输出
    Dim strBack As String

    Command14.Enabled = False

    strBack = GenyYCSS.ErrorCount.Reset(txtIP(1).Text, txtPort(1).Text, IIf(optErrNo(0).Value = True, 1, 2))

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    '进表

    Dim strBack As String
    
    If cmbIP.Text = "" And cmbPort.Text = "" Then
       MsgBox "IP address and port should not be empty!", vbExclamation, "Sys. Notify......"
       Exit Sub
    End If
    
    Command15.Enabled = False

    strBack = GenyYCSS.Pipelinling.PutIn(txtIP(1).Text, txtPort(1).Text)
    
    If InStr(UCase$(strBack), "FAIL") > 0 Then
        labInfo = "Result: Fail！"
    ElseIf InStr(UCase$(strBack), "OK") > 0 Then
        labInfo = "Result: Success！"
    ElseIf InStr(UCase$(strBack), "BUSY") > 0 Then
        labInfo = "Result: MUT is in......"
    ElseIf InStr(UCase$(strBack), "STOWAGE") > 0 Then
        labInfo = "Result: MUT have"
    ElseIf InStr(UCase$(strBack), "ERR") > 0 Then
        labInfo = "Result: Error(" & strBack & ")"
    End If

    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
    '出表
    Dim strBack As String
    
    If cmbIP.Text = "" And cmbPort.Text = "" Then
       MsgBox "IP address and port should not be empty!", vbExclamation, "Sys. Notify......"
       Exit Sub
    End If
    
    Command16.Enabled = False

    strBack = GenyYCSS.Pipelinling.Output(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    ElseIf UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "BUSY" Then
        labInfo = "Result: 正在出表......"
    ElseIf UCase$(strBack) = "NULL" Then
        labInfo = "Result: 无表可出！"
    ElseIf InStr(UCase$(strBack), "ERR") > 0 Then
        labInfo = "Result: 出错(" & strBack & ")"
    End If
    '
    Command16.Enabled = True
End Sub

Private Sub Command17_Click()
    '流水线复位

    Dim strBack As String

    Command17.Enabled = False

    strBack = GenyYCSS.Pipelinling.Reset(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    ElseIf UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "BUSY" Then
        labInfo = "Result: 正在进表......"
    ElseIf InStr(UCase$(strBack), "ERR") > 0 Then
        labInfo = "Result: 出错(" & strBack & ")"
    End If

    Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    '读取误差数据

    Dim strBack As String

    Command18.Enabled = False

    strBack = GenyYCSS.Pipelinling.ReadState(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) <> "FAIL" Then
        labInfo = "Result: Success！"
        Label7.Caption = "Readings: " & strBack
    Else
        labInfo = "Result: Fail！"
        Label7.Caption = "Readings: "
    End If

    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    '串口设置
    Dim strBack As String

    Command6.Enabled = False

    strBack = GenyYCSS.Rs485.SetValue(txtIP(1).Text, txtPort(1).Text, IIf(optRs485(0).Value = True, 1, 2), txt485Model.Text, txt485Door.Text, txtBaueRate.Text)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If
    
    strBack = GenyYCSS.Source.SetComNo(txtIP(1).Text, txtPort(1).Text, IIf(optRs485(0).Value = True, 1, 2), 0)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command6.Enabled = True
End Sub

Private Sub Command1_Click()
    Dim strBack As String

    Command1.Enabled = False

    strBack = GenyYCSS.Source.Reset(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command1.Enabled = True
End Sub

Private Sub Command20_Click()
    '串口设置 以十六进制发送
    '<EhHeader>
    On Error GoTo Command20_Click_Err
    '</EhHeader>
    Dim strBack        As String
    Dim strSendValue() As String
    Dim strBackValue() As String
    Dim i              As Integer

    Command20.Enabled = False

    txtRs485Receive.Text = ""

    If txtRs485Send <> "" Then
        strSendValue = Split(Trim$(txtRs485Send), " ")

        strBack = GenyYCSS.Rs485.SendData(txtIP(1).Text, txtPort(1).Text, IIf(optRs485(0).Value = True, 1, 2), txt485Model.Text, strSendValue, strBackValue)

        If UCase$(strBack) = "OK" Then
            labInfo = "Result: Success！"

            If Val(txt485Model.Text) = 1 Then

                For i = 0 To UBound(strBackValue)
                    strBack = strBack & " " & strBackValue(i)
                Next

                txtRs485Receive = strBack
            End If

        ElseIf UCase$(strBack) = "FAIL" Then
            labInfo = "Result: Fail！"
        End If
    End If

    Command20.Enabled = True
    '<EhFooter>
    Exit Sub

Command20_Click_Err:
    MsgBox Err.Description & vbCrLf & _
           "in LanTest.FrmMain.Command20_Click " & _
           "at line " & Erl, _
           vbExclamation + vbOKOnly, "Application Error"
    Resume Next
    '</EhFooter>
End Sub

Private Sub Command21_Click()
    '读取串口设置
    Dim strBack As String

    Command21.Enabled = False

    strBack = GenyYCSS.Rs485.ReadValue(txtIP(1).Text, txtPort(1).Text, IIf(optRs485(0).Value = True, 1, 2))

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    Else
        txtBaueRate.Text = Left(strBack, Len(strBack) - 2)
        txt485Model.Text = Right(strBack, 1)
    End If

    Command21.Enabled = True
End Sub

Private Sub Command22_Click()
    '读取串口内存数据
    Dim strBack        As String
    Dim strBackValue() As String
    Dim i              As Integer

    Command22.Enabled = False

    txtRs485Receive.Text = ""

    strBack = GenyYCSS.Rs485.ReadData(txtIP(1).Text, txtPort(1).Text, IIf(optRs485(0).Value = True, 1, 2), strBackValue)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
        
        strBack = ""

        For i = 0 To UBound(strBackValue)
            strBack = strBack & " " & strBackValue(i)
        Next

        txtRs485Receive = strBack
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command22.Enabled = True
End Sub

Private Sub Command23_Click()
    Dim strBack As String
    Dim intPara As Integer
     
    If cmbIP.Text = "" And cmbPort.Text = "" Then
       MsgBox "IP address and port should not be empty!", vbExclamation, "Sys. Notify......"
       Exit Sub
    End If
    
    Command23.Enabled = False

    '驱动转轮
    If optInPut(0).Value Then intPara = 0
    If optInPut(1).Value Then intPara = 1
    
    strBack = GenyYCSS.Pipelinling.InPutWheel(txtIP(1).Text, txtPort(1).Text, intPara)

    If UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    ElseIf UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "BUSY" Then
        labInfo = "Result: Turn-in is on......"
    ElseIf InStr(UCase$(strBack), "ERR") > 0 Then
        labInfo = "Result: Error(" & strBack & ")"
    End If
    
    Command23.Enabled = True

End Sub

Private Sub Command24_Click()
    Dim strBack As String
    Dim intPara As Integer
    
    If cmbIP.Text = "" And cmbPort.Text = "" Then
       MsgBox "IP address and port should not be empty!", vbExclamation, "Sys. Notify......"
       Exit Sub
    End If
    
    Command24.Enabled = False
    
    '驱动转轮
    If optOutPut(0).Value Then intPara = 0
    If optOutPut(1).Value Then intPara = 1
    If optOutPut(2).Value Then intPara = 2
    
    strBack = GenyYCSS.Pipelinling.OutPutWheel(txtIP(1).Text, txtPort(1).Text, intPara)

    If UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    ElseIf UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "BUSY" Then
        labInfo = "Result: Turn-out is on......"
    ElseIf InStr(UCase$(strBack), "ERR") > 0 Then
        labInfo = "Result: error(" & strBack & ")"
    End If
    
    Command24.Enabled = True
End Sub

Private Sub Command25_Click()
    '读取版本号

    Dim strBack As String

    Command25.Enabled = False

    strBack = GenyYCSS.Pipelinling.ReadScanCode(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) = "FAIL" Or UCase$(strBack) = "NULL" Or UCase$(strBack) = "" Then
        labInfo = "Result: Fail！"
        Label7.Caption = "Readings: " & strBack
    ElseIf InStr(UCase$(strBack), "TIMEOUT") <> 0 Then
        labInfo = "Result: Out of time！"
        Label7.Caption = "Readings: " & strBack
    Else
        labInfo = "Result: Success！"
        Label7.Caption = "Readings: " & strBack
    End If

    Command25.Enabled = True
End Sub

Private Sub Command26_Click()
    Dim strBack As String
    Dim intPara As Integer
    Dim strIP   As String
    Dim i As Integer
    
    Command26.Enabled = False
    
    '驱动转轮
    If optOutPut(0).Value Then intPara = 0
    If optOutPut(1).Value Then intPara = 1
    If optOutPut(2).Value Then intPara = 2
    
    labInfo = ""
    
    For i = 3 To 7
        
        strIP = "192.168.1." & i
        strBack = GenyYCSS.Pipelinling.OutPutWheel(strIP, "1234", intPara)
    
        If UCase$(strBack) = "FAIL" Then
            labInfo = labInfo & i & "号Fail..."
        ElseIf UCase$(strBack) = "OK" Then
            labInfo = labInfo & i & "号Success..."
        ElseIf UCase$(strBack) = "BUSY" Then
            labInfo = labInfo & i & " is busing..."
        End If
    
    Next i
    
    Command26.Enabled = True
End Sub

Private Sub Command27_Click()
    '读取检测结果

    Dim strBack As String

    Command25.Enabled = False

    strBack = GenyYCSS.Pipelinling.ReadCheckData(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) = "FAIL" Or UCase$(strBack) = "" Then
        labInfo = "Result: Fail！"
    Else
        labInfo = "Result: Success！"
        Label7.Caption = "Readings: " & strBack
    End If

    Command25.Enabled = True
End Sub

Private Sub Command28_Click()
    Dim i       As Integer
    Dim strIP   As String
    Dim intPara As Integer
    Dim strBack As String
    
    Do
        intPara = IIf(intPara = 1, 2, 1)

        For i = 4 To 7
            strIP = "192.168.1." & i

            Do
                strBack = GenyYCSS.Pipelinling.OutPutWheel(strIP, "1234", intPara)
            Loop Until UCase(strBack) = "OK"

            Delay 0.05
        Next i

        Delay 1
    Loop

End Sub

Private Sub Command2_Click()
    Dim strBack As String

    Command2.Enabled = False
    
'   '控制没有标准表
'   '==============================================================
'    Do
'        strBack = GenyYCSS.Source.SetComNo(txtIP(1).Text, txtPort(1).Text, IIf(optComNo(0).Value = True, 1, 2), 0)
'
'        If UCase$(strBack) <> "FAIL" Then
'            labInfo = "Result: Success！"
'            txtData.Text = "Readings: " & strBack
'        Else
'            labInfo = "Result: Fail！"
'            txtData.Text = "Readings: "
'        End If
'
'        Delay (0.5)
'    Loop While (chkAutoRead.Value = 1 And Not IsStop)
'    '==============================================================
'
'    Do
'        strBack = GenyYCSS.Source.SetIONo(txtIP(1).Text, txtPort(1).Text, 1, 0)
'
'        If UCase$(strBack) <> "FAIL" Then
'            labInfo = "Result: Success！"
'            txtData.Text = "Readings: " & strBack
'        Else
'            labInfo = "Result: Fail！"
'            txtData.Text = "Readings: "
'        End If
'
'        Delay (0.5)
'    Loop While (chkAutoRead.Value = 1 And Not IsStop)
'    '==============================================================

    
    strBack = GenyYCSS.Source.SetValue(txtIP(1).Text, txtPort(1).Text, Val(cmbSource.Text), Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text, Text1(16).Text, Text1(15).Text, Text1(14).Text, Text1(13).Text, Text1(12).Text, Text1(21).Text, Text1(20).Text, Text1(19).Text, Text1(18).Text, Text1(17).Text)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command2.Enabled = True
End Sub

Private Sub Command29_Click()
    Dim i As Integer
    Dim strTmp As String
    
    Call SaveIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "chkComReadStandard", chkComReadStandard.Value)
    Call SaveIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtTestTime", txtTestTime.Text)
    Call SaveIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtWaitTIme", txtWaitTIme.Text)
     Call SaveIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "txtReadError", txtReadError.Text)
       
    Call SaveIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "Rows", msfgComTest.Rows)
    With msfgComTest
        For i = 1 To .Rows - 1
           If .TextMatrix(i, 1) <> "" Then
              strTmp = .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3) & "," & .TextMatrix(i, 4) & "," & .TextMatrix(i, 5) & "," & .TextMatrix(i, 6) & "," & .TextMatrix(i, 7) & "," & .TextMatrix(i, 8)
              Call SaveIniInfo(ProgramPath & "LanComTest.ini", "ComTest", "R" & i, strTmp)
           End If
        Next i
    End With
    
    MsgBox "保存Success！", vbInformation, "Sys. Notify..."
End Sub

Private Sub Command3_Click()
    '读取源数据
    Dim strBack As String

    Command3.Enabled = False

    strBack = GenyYCSS.Source.ReadValue(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) <> "FAIL" Then
        labInfo = "Result: Success！"
        txtData.Text = "Readings: " & strBack
    Else
        labInfo = "Result: Fail！"
        txtData.Text = "Readings: "
    End If

    Command3.Enabled = True
End Sub

Private Sub Command30_Click()
      '读取误差数据

    Dim strBack As String
    
    Command30.Enabled = False
    
    Do
        strBack = GenyYCSS.Source.SetComNo(txtIP(1).Text, txtPort(1).Text, IIf(optComNo(0).Value = True, 1, 2), 1)

        If UCase$(strBack) <> "FAIL" Then
            labInfo = "Result: Success！"
            txtData.Text = "Readings: " & strBack
        Else
            labInfo = "Result: Fail！"
            txtData.Text = "Readings: "
        End If

        Delay (0.5)
    Loop While (chkAutoRead.Value = 1 And Not IsStop)
    '==============================================================
    
    Do
        strBack = GenyYCSS.Source.SetIONo(txtIP(1).Text, txtPort(1).Text, 1, 2)

        If UCase$(strBack) <> "FAIL" Then
            labInfo = "Result: Success！"
            txtData.Text = "Readings: " & strBack
        Else
            labInfo = "Result: Fail！"
            txtData.Text = "Readings: "
        End If

        Delay (0.5)
    Loop While (chkAutoRead.Value = 1 And Not IsStop)
    '==============================================================

    Command30.Enabled = True
End Sub

Private Sub Command31_Click()
   '读取误差数据

    Dim strBack As String
    Dim mStdData As clsStandardMeterData
    
    Command31.Enabled = False
    
    Do
        strBack = GenyYCSS.StandardMeter.ReadValue(txtIP(1).Text, txtPort(1).Text)

        If UCase$(strBack) <> "FAIL" Then
            labInfo = "Result: Success！"
            txtData.Text = "Readings: " & strBack
            Call DisplayStandardMeterData(GenyYCSS.StandardMeter.StandardMeterData)
        Else
            labInfo = "Result: Fail！"
            txtData.Text = "Readings: "
        End If

        Delay (0.5)
    Loop While (chkAutoRead.Value = 1 And Not IsStop)

    Command31.Enabled = True
End Sub

Private Sub DisplayStandardMeterData(stdData As clsStandardMeterData)
'     Dim bolTmp As Boolean
'
'     bolTmp = Check4.Value
'
'     txtData.Text = " Ua:" & IIf(bolTmp, stdData.Voltage_A, mStdMeterData.Ua)
'     txtData.Text = txtData.Text & " Ia:" & IIf(bolTmp, stdData.Current_A, mStdMeterData.Ia)
'
'
'     With msfgComStandard
'         .Col = 1
'         .Row = 1
'         .CellForeColor = &HFFFF&
'         .Text = IIf(bolTmp, stdData.Voltage_A, mStdMeterData.Ua)
'         .Row = 2
'         .CellForeColor = &HFFFF&
'         .Text = IIf(bolTmp, stdData.Current_A, mStdMeterData.Ia)
'         .Row = 3
'         .CellForeColor = &HFFFF&
'         .Text = IIf(bolTmp, stdData.Phase_A, mStdMeterData.PhaseA)
'         .Row = 4
'         .CellForeColor = &HFFFF&
'         .Text = IIf(bolTmp, stdData.Power_A, mStdMeterData.Pa)
'         .Row = 5
'         .CellForeColor = &HFFFF&
'         .Text = IIf(bolTmp, stdData.QPower_A, mStdMeterData.Qa)
'         '===================================
'         .Col = 2
'         .Row = 1
'         .CellForeColor = &HFF00&
'         .Text = IIf(bolTmp, stdData.Voltage_B, mStdMeterData.Ub)
'         .Row = 2
'         .CellForeColor = &HFF00&
'         .Text = IIf(bolTmp, stdData.Current_B, mStdMeterData.Ib)
'         .Row = 3
'         .CellForeColor = &HFF00&
'         .Text = IIf(bolTmp, stdData.Phase_B, mStdMeterData.PhaseB)
'         .Row = 4
'         .CellForeColor = &HFF00&
'         .Text = IIf(bolTmp, stdData.Power_B, mStdMeterData.Pb)
'         .Row = 5
'         .CellForeColor = &HFF00&
'         .Text = IIf(bolTmp, stdData.QPower_B, mStdMeterData.Qb)
'         '===================================
'         .Col = 3
'         .Row = 1
'         .CellForeColor = &HFF&
'         .Text = IIf(bolTmp, stdData.Voltage_C, mStdMeterData.Uc)
'         .Row = 2
'         .CellForeColor = &HFF&
'         .Text = IIf(bolTmp, stdData.Current_C, mStdMeterData.Ic)
'         .Row = 3
'         .CellForeColor = &HFF&
'         .Text = IIf(bolTmp, stdData.Phase_C, mStdMeterData.PhaseC)
'         .Row = 4
'         .CellForeColor = &HFF&
'         .Text = IIf(bolTmp, stdData.Power_C, mStdMeterData.Pc)
'         .Row = 5
'         .CellForeColor = &HFF&
'         .Text = IIf(bolTmp, stdData.QPower_C, mStdMeterData.Qc)
'
'         labStdConstant.Caption = Format(Val(stdData.Constant), "Scientific")
'     End With
End Sub

Private Sub Command32_Click()
   Dim strBack As String

    Command32.Enabled = False
    
    strBack = GenyYCSS.Source.SetPercentValue(txtIP(1).Text, txtPort(1).Text, Text1(0).Text, Text1(39).Text, Text1(1).Text, Text1(40).Text, Text1(16).Text, Text1(41).Text, Text1(15).Text, Text1(42).Text, Text1(21).Text, Text1(43).Text, Text1(20).Text, Text1(44).Text)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command32.Enabled = True
End Sub

Private Sub Command33_Click()
    With msfgComTest
      .AddItem .Rows & vbTab & Val(cmbComModel.Text) & vbTab & Val(Text1(26).Text) & vbTab & Val(Text1(25).Text) & vbTab & Val(Text1(24).Text) & vbTab & Val(Text1(23).Text) & vbTab & Val(Text1(22).Text) & vbTab & Val(Combo2.Text) & vbTab & Combo1.Text
    End With
End Sub

Private Sub Command34_Click()
    On Error Resume Next
    
    If intRow = 0 Then Exit Sub
    msfgComTest.RemoveItem intRow
End Sub

Private Sub Command35_Click()
    Dim strBack As String

    Command35.Enabled = False

    strBack = GenyYCSS.ErrorCount.ClearReadData(txtIP(1).Text, txtPort(1).Text, IIf(optErrNo(0).Value = True, 1, 2))

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command35.Enabled = True
End Sub

Private Sub Command36_Click()
    Dim strBack    As String
    Dim strU       As String
    Dim strI       As String
    Dim strSPower  As String
    Dim strRePower As String
    Dim strAcPower As String
    Dim strPeriod  As String
    Dim strFr      As String
    Dim strAngle   As String
    
    Command36.Enabled = False

    strBack = GenyYCSS.Source.ReadWastageValue(txtIP(1).Text, txtPort(1).Text, 2, Val(txtGHAddress.Text), strU, strI, strSPower, strRePower, strAcPower, strPeriod, strFr, strAngle)

    If UCase$(strBack) <> "FAIL" Then
        labInfo = "Result: Success！"
        txtData.Text = "Readings: " & "U:" & strU & " I:" & strI & " SPower:" & strSPower & " RePower:" & strRePower & " AcPower:" & strAcPower & " Period:" & strPeriod & " Fr:" & strFr & " Angler:" & strAngle
    Else
        labInfo = "Result: Fail！"
        txtData.Text = "Readings: "
    End If

    Command36.Enabled = True
End Sub

Private Sub Command37_Click()
     Dim strBack As String

    Command37.Enabled = False
   
    strBack = GenyYCSS.Other.OpenLock(txtIP(1).Text, txtPort(1).Text)
    
    If UCase$(strBack) = "OK" Then
         labInfo = "Result: Success！"
     ElseIf UCase$(strBack) = "FAIL" Then
         labInfo = "Result: Fail！"
         Exit Sub
     End If
   
    strBack = GenyYCSS.Source.SetValue(txtIP(1).Text, txtPort(1).Text, Val(cmbSource.Text), Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text, Text1(16).Text, Text1(15).Text, Text1(14).Text, Text1(13).Text, Text1(12).Text, Text1(21).Text, Text1(20).Text, Text1(19).Text, Text1(18).Text, Text1(17).Text)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command37.Enabled = True
End Sub

Private Sub Command38_Click()
    Dim strBack As String

    Command38.Enabled = False
   
    strBack = GenyYCSS.Other.OpenLock(txtIP(1).Text, txtPort(1).Text)
    
    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If
    
    strBack = GenyYCSS.Source.SetFactor(txtIP(1).Text, txtPort(1).Text, Text1(27).Text, Text1(28).Text, Text1(29).Text, Text1(30).Text, Text1(31).Text, Text1(32).Text, Text1(33).Text, Text1(34).Text, Text1(35).Text, Text1(36).Text, Text1(37).Text, Text1(38).Text)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command38.Enabled = True
End Sub

Private Sub Command39_Click()
    '读取
    Dim strBack As String

    Command39.Enabled = False

    strBack = GenyYCSS.Source.ReadFactor(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) <> "FAIL" Then
        labInfo = "Result: Success！"
        txtData.Text = "Readings: " & strBack
    Else
        labInfo = "Result: Fail！"
        txtData.Text = "Readings: "
    End If

    Command39.Enabled = True
End Sub

Private Sub Command4_Click()
    '读取误差数据

    Dim strBack As String

    Command4.Enabled = False

    Do
        strBack = GenyYCSS.ErrorCount.ReadValue(txtIP(1).Text, txtPort(1).Text, IIf(optErrNo(0).Value = True, 1, 2))

        If UCase$(strBack) <> "FAIL" Then
            labInfo = "Result: Success！"
            txtData.Text = "Readings: " & strBack
        Else
            labInfo = "Result: Fail！"
            txtData.Text = "Readings: "
        End If

        Delay (0.5)
    Loop While (chkAutoRead.Value = 1 And Not IsStop)

    Command4.Enabled = True
End Sub

Private Sub Command40_Click()
    FrmIP.Show vbModal
End Sub

Private Sub Command41_Click()
    Dim i As Integer
    Dim sTmp As String
    
    With msfgComTest
        If .Rows > 1 And .Row <> 0 Then
            If .Row > 1 Then
                For i = 1 To .Cols - 1
                    sTmp = .TextMatrix(.Row, i)
                    .TextMatrix(.Row, i) = .TextMatrix(.Row - 1, i)
                    .TextMatrix(.Row - 1, i) = sTmp
                Next i
                .Row = .Row - 1
                .Col = 0
            End If

        End If
    End With
End Sub

Private Sub Command42_Click()
    Dim i As Integer
    Dim sTmp As String
    
    With msfgComTest
        If .Rows > 0 And .Row <> 0 Then
            If .Row < .Rows - 1 Then
                For i = 1 To .Cols - 1
                    sTmp = .TextMatrix(.Row, i)
                    .TextMatrix(.Row, i) = .TextMatrix(.Row + 1, i)
                    .TextMatrix(.Row + 1, i) = sTmp
                Next i
                .Row = .Row + 1
                .Col = 0
            End If
        End If
    End With
End Sub

Private Sub Command43_Click()
    Dim strBack As String
    Dim EHP As ErrorHeadParam
    Dim EMP(1 To 8) As ErrorMeterParam
    Dim LSH As LoadSwitch
    Dim ReadLSH As LoadSwitch
    Dim Mark As Integer
     Dim strSendValue() As String
    Dim strBackValue() As String
    
    'strBack = GenyYCSS.MultiErrorCount.SetValue("192.168.1.12", 1234, EHP, EMP)

    Mark = 0
    LSH = GenyYCSS.MultiErrorCount.IniLoadSwitch
    LSH.MeterNo1 = Mark
    LSH.MeterNo2 = Mark
    LSH.MeterNo3 = Mark
    LSH.MeterNo4 = Mark
    strBack = GenyYCSS.MultiErrorCount.SetLoadSwitch("192.168.1.12", 1234, LSH)
    strBack = GenyYCSS.MultiErrorCount.ReadLoadSwitch("192.168.1.12", 1234, ReadLSH)

'     strBack = GenyYCSS.MultiErrorCount.SetComBaudRate("192.168.1.12", 1234, 1, 0, 0, txtBaueRate.Text)
'     strSendValue = Split(Trim$(txtRs485Send), " ")
'     strBack = GenyYCSS.MultiErrorCount.SendComData("192.168.1.12", 1234, 1, strSendValue)
'
'     strBack = GenyYCSS.MultiErrorCount.ReadComData("192.168.1.12", 1234, 1, strBackValue)
     
End Sub

Private Sub Command44_Click()
    Dim strBack As String
    Dim BackData() As String
    
    strBack = GenyYCSS.MultiErrorCount.ReadValue("192.168.1.12", 1234, BackData)

End Sub

Private Sub Command45_Click()
   Call SetError(1, "100000000")
End Sub

Private Sub Command5_Click()
    '读取版本号

    Dim strBack As String

    Command5.Enabled = False

    strBack = GenyYCSS.Other.ReadVersion(txtIP(1).Text, txtPort(1).Text)

    If UCase$(strBack) <> "FAIL" Then
        labInfo = "Result: Success！"
        Label7.Caption = "Readings: " & strBack
    Else
        labInfo = "Result: Fail！"
        Label7.Caption = "Readings: "
    End If

    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    '误差输出
    Dim strBack As String

    Command6.Enabled = False

    strBack = GenyYCSS.ErrorCount.SetValue(txtIP(1).Text, txtPort(1).Text, IIf(optErrNo(0).Value = True, 1, 2), Val(cmbWay.Text), Val(cmbErrMode.Text), _
                                            Text1(6).Text, Text1(5).Text, Text1(7).Text, Text1(8).Text, Text1(9).Text, Text1(45).Text, Text1(46).Text, Text1(47).Text, Text1(48).Text)

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command6.Enabled = True
End Sub

Private Sub Command7_Click()

    If intRow > 1 And intRow < msfgXB.Rows Then
        msfgXB.TextMatrix(intRow, 1) = 0
        msfgXB.TextMatrix(intRow, 2) = 0
        msfgXB.TextMatrix(intRow, 3) = 0
    End If

End Sub

Private Sub Command8_Click()
    '谐波输出
    Dim strBack               As String
    Dim strAmplitude(0 To 11) As String
    Dim strPhase(0 To 11)     As String
    Dim strTimes(0 To 11)     As String
    Dim i                     As Integer

    Command8.Enabled = False

    For i = 1 To msfgXB.Rows - 1
        strAmplitude(i - 1) = Val(msfgXB.TextMatrix(i, 1))
        strPhase(i - 1) = Val(msfgXB.TextMatrix(i, 2))
        strTimes(i - 1) = Val(msfgXB.TextMatrix(i, 3))
    Next i

    strBack = GenyYCSS.Harmonics.SetValue(txtIP(1).Text, txtPort(1).Text, Val(cmbSource2.Text), IIf(optXB(0).Value = True, "U", "I"), strAmplitude(), strPhase(), strTimes(), strAmplitude(), strPhase(), strTimes(), strAmplitude(), strPhase(), strTimes())

    If UCase$(strBack) = "OK" Then
        labInfo = "Result: Success！"
    ElseIf UCase$(strBack) = "FAIL" Then
        labInfo = "Result: Fail！"
    End If

    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Dim i As Integer

    With msfgXB

        For i = 2 To .Rows - 1

            If Val(.TextMatrix(i, 1)) = 0 Then
                .TextMatrix(i, 1) = Text1(10)
                .TextMatrix(i, 2) = Text1(11)
                .TextMatrix(i, 3) = Val(cmbXB.Text)
                Exit For
            End If

        Next

    End With

End Sub

Private Sub Form_Load()
    '=======初始化数据=============
    IsStop = False  '是否停止运行标志

    ProgramPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")
    
    LogFullName = ProgramPath & "Log\" & Format(Date, "YYYY-MM-DD") & ".txt"
    '========================================================================
    '从注册表得到设置信息
    txtIP(0).Text = GetSetting("LanTest", "SetData", "IP", "192.168.1.101")
    txtPort(0).Text = GetSetting("LanTest", "SetData", "Port", "1234")
    txtComm(0).Text = GetSetting("LanTest", "SetData", "Comm", "0")
    txtIP(1).Text = GetSetting("LanTest", "SetData", "IP", "192.168.1.101")
    txtPort(1).Text = GetSetting("LanTest", "SetData", "Port", "1234")
    txtComm(1).Text = GetSetting("LanTest", "SetData", "Comm", "0")
    '=========================================================================

    Call IniContrcols
End Sub

Private Sub Form_Resize()
    Dim sngTmp As Single
    Dim i As Integer
    
    With msfgComTest
         sngTmp = (.Width - .ColWidth(0) - 50) / (.Cols)
         
         For i = 1 To .Cols - 1
'         .ColWidth(i) = sngTmp
         .ColAlignment(0) = flexAlignCenterCenter
         Next i
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    blnStop = True
    End
End Sub

Private Sub msfgComTest_Click()
     intRow = msfgComTest.Row
End Sub

Private Sub msfgXB_Click()
    intRow = msfgXB.Row
End Sub

Private Sub Option1_Click(Index As Integer)
     If Index = 0 Then
        txtBaueRate.Text = "2400,E,8,1"
        txtRs485Send.Text = "FE FE FE FE 68 AA AA AA AA AA AA 68 11 04 33 33 34 33 AE 16"
     ElseIf Index = 1 Then
        txtBaueRate.Text = "1200,E,8,1"
        txtRs485Send.Text = "FE FE FE FE 68 99 99 99 99 99 99 68 01 02 43 C3 6F 16"
     Else
        txtBaueRate.Text = "1200,E,8,1"
        txtRs485Send.Text = "FE FE FE FE 68 AA AA AA AA AA AA 68 01 02 43 C3 D5 16"
     End If
End Sub

Private Sub Text1_Change(Index As Integer)
     Select Case Index
     Case 0
          Text1(16).Text = Text1(Index).Text
          Text1(21).Text = Text1(Index).Text
     Case 16
          Text1(21).Text = Text1(Index).Text
     Case 1
          Text1(15).Text = Text1(Index).Text
          Text1(20).Text = Text1(Index).Text
     Case 15
          Text1(20).Text = Text1(Index).Text
     Case 2
          Text1(14).Text = Text1(Index).Text
          Text1(19).Text = Text1(Index).Text
     Case 14
          Text1(19).Text = Text1(Index).Text
     Case 4
          Text1(12).Text = Text1(Index).Text
          Text1(17).Text = Text1(Index).Text
     Case 14
          Text1(17).Text = Text1(Index).Text
     End Select
End Sub

Private Sub Timer1_Timer()
    
    Label2(40).Caption = Format(DateAdd("s", 1, Format(Label2(40).Caption, "HH:MM:SS")), "HH:MM:SS")

    If DateDiff("s", "00:00:00", Format(Label2(40).Caption, "HH:MM:SS")) >= Val(txtTestTime) * 3600 Then
        bolTimeStop = True
        Timer1.Enabled = False
    End If

End Sub

Private Sub txtInPut_KeyPress(KeyAscii As Integer)
    Dim strSend As String

    If KeyAscii = 32 Then
        strSend = getHexData(txtInPut.Text)
        txtSend.Text = strSend
    End If

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim bytGetData() As Byte

    Dim strGetData   As String

    '接收网卡返回数据
    Winsock1.GetData bytGetData, vbArray + vbByte

    If UBound(bytGetData()) > 0 Then
        '转换得到数组为字符串
        strGetData = getReceiveData(bytGetData())

        '显示在接受框内
        If Len(strGetData) > 0 Then
            txtGet.Text = txtGet.Text & strGetData
        End If
    End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, _
                           Description As String, _
                           ByVal Scode As Long, _
                           ByVal Source As String, _
                           ByVal HelpFile As String, _
                           ByVal HelpContext As Long, _
                           CancelDisplay As Boolean)

    '打印错误信息
    If binShowMsg Then
        MsgBox Description, vbCritical, "Error Msg."
        Err.Clear
        binShowMsg = False
    End If

End Sub
