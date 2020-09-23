VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make Your ..."
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   2565
   FillColor       =   &H80000014&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   2565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameCut 
      Height          =   645
      Left            =   60
      TabIndex        =   67
      Top             =   7890
      Width           =   2445
      Begin VB.CheckBox chkTree 
         Caption         =   "Cut Operation"
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   25
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   10
         Left            =   1650
         MaxLength       =   5
         TabIndex        =   26
         Text            =   "20"
         Top             =   180
         Width           =   450
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   315
         Index           =   10
         Left            =   2100
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   20
         BuddyControl    =   "txtInput(10)"
         BuddyDispid     =   196611
         BuddyIndex      =   10
         OrigLeft        =   2101
         OrigTop         =   180
         OrigRight       =   2356
         OrigBottom      =   495
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   285
         Index           =   5
         Left            =   1545
         TabIndex        =   69
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame frameButtons 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   -30
      TabIndex        =   47
      Top             =   8535
      Width           =   2625
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   120
         Left            =   30
         TabIndex        =   49
         Top             =   -30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   212
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.CommandButton cmdFormSize 
         Height          =   315
         Left            =   2280
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   300
         Width           =   345
      End
      Begin VB.CommandButton cmdStop 
         Cancel          =   -1  'True
         Caption         =   "Stop !"
         Height          =   405
         Left            =   1470
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdTimerStop 
         Height          =   375
         Left            =   90
         Picture         =   "frmMain.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Stop Timer"
         Top             =   120
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Do it !"
         Default         =   -1  'True
         Height          =   405
         Left            =   720
         TabIndex        =   27
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame FrameFlowers 
      Height          =   1635
      Left            =   60
      TabIndex        =   40
      Top             =   6220
      Width           =   2445
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   9
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "20"
         Top             =   900
         Width           =   450
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   8
         Left            =   1650
         MaxLength       =   4
         TabIndex        =   20
         Text            =   "20"
         Top             =   180
         Width           =   450
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Flowers"
         Height          =   315
         Index           =   8
         Left            =   180
         TabIndex        =   19
         Top             =   180
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   8
         Left            =   1140
         TabIndex        =   21
         ToolTipText     =   "50"
         Top             =   540
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   9
         Left            =   1140
         TabIndex        =   24
         ToolTipText     =   "50"
         Top             =   1260
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Fruits"
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   22
         Top             =   900
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   315
         Index           =   8
         Left            =   2101
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtInput(8)"
         BuddyDispid     =   196611
         BuddyIndex      =   8
         OrigLeft        =   1200
         OrigRight       =   1455
         OrigBottom      =   525
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   315
         Index           =   9
         Left            =   2100
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   900
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtInput(9)"
         BuddyDispid     =   196611
         BuddyIndex      =   9
         OrigRight       =   255
         OrigBottom      =   525
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   285
         Index           =   4
         Left            =   1545
         TabIndex        =   44
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Fruit Speed :"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   43
         Top             =   1260
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Flower Size  :"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   42
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   285
         Index           =   0
         Left            =   1545
         TabIndex        =   41
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame FrameMain 
      Height          =   5575
      Left            =   60
      TabIndex        =   32
      Top             =   630
      Width           =   2445
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "0"
         Top             =   1645
         Width           =   450
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         Top             =   1330
         Width           =   450
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "1"
         Top             =   1015
         Width           =   450
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Broken Branches "
         Height          =   345
         Index           =   7
         Left            =   180
         TabIndex        =   9
         Top             =   2590
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "30"
         Top             =   210
         Width           =   630
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   4
         Left            =   1140
         TabIndex        =   15
         ToolTipText     =   "50"
         Top             =   4300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1650
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "2"
         Top             =   700
         Width           =   450
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   1650
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "30"
         Top             =   2590
         Width           =   450
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   1650
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "35"
         Top             =   2275
         Width           =   450
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "65"
         Top             =   1960
         Width           =   450
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Fixed Size"
         Height          =   315
         Index           =   5
         Left            =   180
         TabIndex        =   5
         Top             =   1960
         Width           =   1215
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   2
         Left            =   1140
         TabIndex        =   13
         ToolTipText     =   "50"
         Top             =   3700
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   7
         Left            =   1140
         TabIndex        =   18
         ToolTipText     =   "500"
         Top             =   5200
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   1000
         SelStart        =   500
         TickFrequency   =   100
         Value           =   500
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   3
         Left            =   1140
         TabIndex        =   14
         ToolTipText     =   "50"
         Top             =   4000
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   6
         Left            =   1140
         TabIndex        =   17
         ToolTipText     =   "50"
         Top             =   4900
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   285
         Index           =   7
         Left            =   2101
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2590
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtInput(7)"
         BuddyDispid     =   196611
         BuddyIndex      =   7
         OrigRight       =   255
         OrigBottom      =   525
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   285
         Index           =   6
         Left            =   2101
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2275
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtInput(6)"
         BuddyDispid     =   196611
         BuddyIndex      =   6
         OrigRight       =   255
         OrigBottom      =   525
         Max             =   9999
         Min             =   -9999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   285
         Index           =   5
         Left            =   2101
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txtInput(5)"
         BuddyDispid     =   196611
         BuddyIndex      =   5
         OrigRight       =   255
         OrigBottom      =   525
         Max             =   99
         Min             =   5
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   285
         Index           =   1
         Left            =   2101
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   700
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txtInput(1)"
         BuddyDispid     =   196611
         BuddyIndex      =   1
         OrigRight       =   255
         OrigBottom      =   525
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   420
         Index           =   0
         Left            =   2101
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   210
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   741
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "txtInput(0)"
         BuddyDispid     =   196611
         BuddyIndex      =   0
         OrigRight       =   255
         OrigBottom      =   525
         Max             =   99
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Fixed Angel"
         Height          =   315
         Index           =   6
         Left            =   180
         TabIndex        =   7
         Top             =   2275
         Width           =   1365
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   5
         Left            =   1140
         TabIndex        =   16
         ToolTipText     =   "50"
         Top             =   4600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   0
         Left            =   1140
         TabIndex        =   11
         ToolTipText     =   "50"
         Top             =   3100
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   285
         Index           =   3
         Left            =   2101
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1330
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtInput(3)"
         BuddyDispid     =   196611
         BuddyIndex      =   3
         OrigRight       =   255
         OrigBottom      =   525
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.Slider sldTree 
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   12
         ToolTipText     =   "50"
         Top             =   3400
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   285
         Index           =   2
         Left            =   2101
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1015
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txtInput(2)"
         BuddyDispid     =   196611
         BuddyIndex      =   2
         OrigRight       =   255
         OrigBottom      =   525
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   285
         Index           =   4
         Left            =   2101
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1645
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtInput(4)"
         BuddyDispid     =   196611
         BuddyIndex      =   4
         OrigLeft        =   480
         OrigTop         =   480
         OrigRight       =   735
         OrigBottom      =   1005
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Recision Level Low:"
         Height          =   225
         Index           =   14
         Left            =   180
         TabIndex        =   65
         Top             =   1645
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "B.P.S Scale :"
         Height          =   225
         Index           =   13
         Left            =   180
         TabIndex        =   63
         Top             =   3400
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Recision Level Hi :"
         Height          =   225
         Index           =   12
         Left            =   180
         TabIndex        =   62
         Top             =   1330
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Starting Branch :"
         Height          =   225
         Index           =   11
         Left            =   180
         TabIndex        =   60
         Top             =   1015
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "B.B Scale :"
         Height          =   225
         Index           =   10
         Left            =   180
         TabIndex        =   59
         Top             =   3100
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Width Scale :"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   58
         Top             =   4600
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Wind :"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   50
         Top             =   4900
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Trunk Height:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   48
         Top             =   4000
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Widening  :"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   46
         Top             =   5200
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Width :"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   39
         Top             =   4300
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Leaf Level :"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   38
         Top             =   3700
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Branches per Step :"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   37
         Top             =   700
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   285
         Index           =   3
         Left            =   1545
         TabIndex        =   36
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1545
         TabIndex        =   35
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   285
         Index           =   1
         Left            =   1545
         TabIndex        =   34
         Top             =   2010
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Total Steps :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   33
         Top             =   270
         Width           =   1545
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "V 2.0"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2070
      TabIndex        =   45
      Top             =   390
      Width           =   465
   End
   Begin VB.Label lblTree 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tree"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   60
      TabIndex        =   31
      Top             =   30
      Width           =   2445
   End
   Begin VB.Menu mnuTitleFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSavePicture 
         Caption         =   "Save Picture"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSaveProfile 
         Caption         =   "Save Current Profile"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTitleTypes 
      Caption         =   "&Types"
      Begin VB.Menu mnuSubTypes1 
         Caption         =   "Fractals"
         Begin VB.Menu mnuTypesFractals 
            Caption         =   "Fractal"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSubTypes2 
         Caption         =   "Normal Plants"
         Begin VB.Menu mnuTypesPlants 
            Caption         =   "Plant"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSubTypes3 
         Caption         =   "User Favorites"
         Begin VB.Menu mnuTypesUsers 
            Caption         =   "User"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuTitleOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "On Top"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Fast Paint"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Sounds"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Random Background"
         Checked         =   -1  'True
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'       Real Tree 2 (Flowers & Fruits)
'       Copyright (c) 2006 Mohammad Reza Khosravi ( Khosravi2500@yahoo.com )
'
'
'1.     Special thanks for all comments, suggestions and votes in previous version.
'
'2.     This program uses only simple methods for working with graphics.
'
'3.     It's recommended that beginners initially see the previous version at the following address:
'       http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=66280&lngWid=1
'
'4.     See tree.ini file.
'
'5.     Click on flowers and Fruits in run time, and see results.
'
'6.     "Widening" effect is very sensitive and I used values (0-1000) for it, it needs some changes.
'       also I used right click on sliders for small changes.
'
'7.     "B.B Scale" = "Broken Branches Scale" and means broken effect acts more on lower branches or higher branches.
'       "B.P.S Scale" = "Branches Per Step Scale", when you lead to right, you have more BPS in higher branches than base BPS,
'       and when you lead to left, you have less BPS in higher branches than base BPS.
'
'8.     For saving pictures, you can use Menu or "F2" key in anytime (even when rendering),
'       I used Jpeg Encoder Class written by Mr. John Korejwa, it's not a beginner code,
'       so if you are a beginner, you can ignore "Save Picture" option and remove cJpeg.cls from project.
'
'9.     Updates:
'           October 6... added "Save Current Profile".
'           October 10.. added "Save Picture".
'           October 15.. added "Cut Operation".
'
'
Dim titleHeight As Integer

Private Sub Form_Activate()
    If titleHeight = 0 Then titleHeight = Me.Height - 9175 ' in different windows (98-2000-xp), title bar height is different, this is simple way for fixing this problem when I want to resize this Form.
End Sub

Private Sub Form_Load()
    Randomize Timer
    cmdStop.Left = cmdOK.Left
    loadDefaultTrees
    LoadUserTrees
End Sub


' this is for making new Menu items and loading data in run time (fractals, normal and user Trees)
Private Sub loadDefaultTrees()
    'maybe somebody wants to use Array for holding such data, but here, I prefer to use "Tag" of Menu items
    'sequence of Numbers are based on "TabIndex" property of textBoxes, checkBoxes and Sliders on Form.
    '(because of my English, maybe some names do not match!)
    
    addMenuItem mnuTypesFractals, "Simple,15,2,,,,1,70,1,45,0,,,,,,,,,,0,,,0,,,,,", 0
    addMenuItem mnuTypesFractals, "Fantasy 1 (Circular),5,90,2,0,0,1,90,1,6000,0,30,50,10,50,85,60,90,50,912,0,,,0,,,0,"
    addMenuItem mnuTypesFractals, "Fantasy 2,7,5,3,,,1,70,1,100,0,,,,60,100,20,,,100,0,,,0,,,,,"
    addMenuItem mnuTypesFractals, "Fantasy 3 (+Flowers),7,11,,,,1,70,1,45,0,,,30,100,100,20,75,,605,1,80,60,1,45,95,,,"
    addMenuItem mnuTypesFractals, "Fantasy 4,12,3,2,,,1,70,1,180,0,,,,70,100,20,,,731,0,,,0,,,,,"
    addMenuItem mnuTypesFractals, "Fantasy 5 (Pentagon),9,5,7,,,1,70,1,180,0,,,,100,90,20,70,,,0,,,0,,,,,"
    addMenuItem mnuTypesFractals, "Fantasy 6 (Triangle),11,3,6,,,1,70,1,180,0,,,,90,75,10,70,,,0,,,0,,,,,"
    addMenuItem mnuTypesFractals, "Fantasy 7,13,8,4,,,1,70,1,1265,0,,,16,100,80,40,,,827,0,,,0,,,,,"
    addMenuItem mnuTypesFractals, "Fantasy 8 (Thorn),12,10,4,,,1,70,1,220,0,,,5,100,100,40,,,,0,,,0,,,,,"
    addMenuItem mnuTypesFractals, "Fantasy 9 (Swastika),15,8,5,,,1,70,1,230,0,,,11,100,90,25,45,,896,0,,,0,,,,,"
   
    addMenuItem mnuTypesPlants, "Default,30,,,,,,,,,,,,,,,,,,,,,,,,,,,", 0
    addMenuItem mnuTypesPlants, "Simple,30,,,,,,,,,,25,,,,,,,,,0,,,0,,,,,"
    addMenuItem mnuTypesPlants, "Garden,40,2,,,,1,60,1,30,1,33,95,,60,,40,,,,0,,,0,,,,,"
    addMenuItem mnuTypesPlants, "Heavy 1,32,3,1,,,0,65,0,30,1,54,65,,,55,30,45,55,,1,20,30,1,,85,,,"
    addMenuItem mnuTypesPlants, "Heavy 2,40,4,,,2,0,65,0,38,1,68.7,60,,55,60,30,44,,,1,10,20,1,,90,,,"
    addMenuItem mnuTypesPlants, "Heavy 3,40,6,3,,,0,55,0,25,1,82,65,,60,25,100,60,,,0,,,0,,,,,"
    addMenuItem mnuTypesPlants, "Wild 1,38,5,,,,0,65,0,68,1,75,55,55,55,60,30,45,,,1,20,20,0,,,,,"
    addMenuItem mnuTypesPlants, "Wild 2,65,5,,0,30,0,70,0,35,1,80,58,,27,65,25,45,55,,0,,,0,,,,,"
    addMenuItem mnuTypesPlants, "Wild Flower 1,30,7,3,,,0,60,0,30,1,81,45,,65,10,10,35,60,,1,5,100,0,,,,,"
    addMenuItem mnuTypesPlants, "Wild Flower 2,23,9,3,,,0,20,0,22,1,88,,,5,0,,100,70,560,1,50,90,0,,,,,"
    addMenuItem mnuTypesPlants, "Teazle,25,5,3,,,0,30,0,40,1,20,25,10,20,0,40,,,,0,,,0,,,,,"
    addMenuItem mnuTypesPlants, "Spring,35,2,,,,0,60,0,30,1,30,55,55,60,45,30,45,,,0,,,0,,,,,"
    
    addMenuItem mnuTypesUsers, "David Malekan,24,4,2,,,0,20,0,30,1,67,40,,30,10,40,30,,,0,,,0,,,,,", 0 'David is my friend
    
End Sub

' Read Tree.ini file from hard disk and insert items to user menu
Private Sub LoadUserTrees()
    On Error GoTo errHandler
    Dim myFreeFile As Integer, myProfile As String
    
    myFreeFile = FreeFile
    Open App.Path & "\Tree.ini" For Input As #myFreeFile
        While Not EOF(myFreeFile)
            Line Input #myFreeFile, myProfile
            myProfile = Trim(myProfile)
            If Left(myProfile, 1) <> "'" Then
                addMenuItem mnuTypesUsers, myProfile
            End If
        Wend
    Close #myFreeFile
Exit Sub
errHandler:
    ' it is better just go out.
End Sub

' Make new Menu items in run time.
Private Sub addMenuItem(ByVal mySubMenu As Object, ByVal myProfile As String, Optional fixIndex As Integer = -1)
    Dim j As Integer, myIndex As Integer
    
    j = InStr(myProfile, ",")
    If j > 1 Then
        myIndex = mySubMenu.UBound + 1    ' get highest number of index and add one
        If fixIndex > -1 Then myIndex = fixIndex Else _
        Load mySubMenu(myIndex)           ' make new Menu item. for first item of menus, I don't need a new object because I made it in design time.
        With mySubMenu(myIndex)           ' set properties of new Menu item
            .Caption = Left(myProfile, j - 1)
            .Tag = Mid(myProfile, j + 1)
        End With
    End If
End Sub


' this is for small changes in values of Sliders when click right mouse button on them
Private Sub sldTree_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With sldTree(Index)
            If x > .Value * (.Width - 200) / .Max + 100 Then .Value = .Value + 1 Else .Value = .Value - 1
        End With
    End If
End Sub

Private Sub sldTree_Change(Index As Integer)
    sldTree(Index).ToolTipText = sldTree(Index).Value
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    txtInput(Index).SelStart = 0
    txtInput(Index).SelLength = Len(txtInput(Index).Text)
End Sub


Private Sub chkTree_Click(Index As Integer)
    If Index < 7 Then Exit Sub ' I made some changes in ver 2
    If chkTree(Index).Value = 1 Then
        txtInput(Index).Enabled = True
        txtInput(Index).BackColor = vbWindowBackground
    Else
        txtInput(Index).Enabled = False
        txtInput(Index).BackColor = vbButtonFace
    End If

    'when you deactivate "Flowers", it's better "Fruits" became automatically deactivated. try it!
End Sub

'this is for changing height of Form
Private Sub cmdFormSize_Click()
    Static formSize As Boolean
    
    formSize = Not formSize
    frameButtons.Top = IIf(formSize, lblTree.Top + lblTree.Height, frameCut.Top + frameCut.Height) ' because I have several Buttons, I put all of them on a Frame and when I need ,I move Frame.
    Me.Height = frameButtons.Top + frameButtons.Height + titleHeight  ' titleheight is for Title bar and menu bar heights.
    If cmdOK.Visible Then cmdOK.SetFocus Else cmdStop.SetFocus   ' I don't like this button get focus so I change it.
End Sub


Private Sub cmdOK_Click()
    Static noRun As Boolean
    
    If frmTree.imgFlower.UBound > 4000 Then Unload frmTree ' As I wrote in next line, removing a high number of objects takes a long time, so I prefer reload Form.  If you have better idea (but only with simple programming)please let me know.
    cmdOK.Caption = "Wait..." 'you must know that flowers and fruits are objects (like TextBoxes),  if your last run had make a huge numbers of flowers and you want to try a new Tree, maybe removing of the last flowers takes  several secounds.
    
    If Not noRun Then ' only for the first time of running
        noRun = True
        Me.Left = 0
        Me.Top = 0
        frmTree.Left = Me.Width
        frmTree.Top = 0
        frmTree.Width = Screen.Width - Me.Width
        frmTree.Height = Screen.Height - 500
    End If
    
    ' restoring variables from textboxes, sliders and checkboxes
    myTree.totalSteps = Abs(Val(txtInput(0).Text) Mod 100)
    myTree.totalSteps = IIf(myTree.totalSteps < 3, 3, myTree.totalSteps): txtInput(0).Text = myTree.totalSteps
    
    myTree.branchPerStep = Val(txtInput(1).Text)
    myTree.branchPerStep = IIf(myTree.branchPerStep <= 0, 1, myTree.branchPerStep): txtInput(1).Text = myTree.branchPerStep
    
    myTree.startingBranch = Abs(Val(txtInput(2).Text) Mod 100)
    myTree.recisionLevelHi = Abs(Val(txtInput(3).Text) Mod 100)
    myTree.recisionLevelLow = Abs(Val(txtInput(4).Text) Mod 100)
    
    myTree.fixSize = IIf(chkTree(5).Value = 1, True, False)
    myTree.maxSize = Abs(Val(txtInput(5).Text) Mod 100)
    
    myTree.fixAngel = IIf(chkTree(6).Value = 1, True, False)
    myTree.sizeOfAngel = Val(txtInput(6).Text)
    
    myTree.brokenBranches = IIf(chkTree(7).Value = 1, Abs(Val(txtInput(7).Text) * 10 Mod 1000) / 10, 0)
    myTree.brokenBranchesScale = sldTree(0).Value
    myTree.branchPerStepScale = sldTree(1).Value
    myTree.leafLevel = sldTree(2).Value
    myTree.trunkHeight = sldTree(3).Value
    myTree.widthSize = sldTree(4).Value
    myTree.widthScale = sldTree(5).Value
    myTree.windDirection = sldTree(6).Value
    myTree.wideLevel = sldTree(7).Value

    myTree.flowerNeed = IIf(chkTree(8).Value = 1, Abs(Val(txtInput(8).Text) * 10 Mod 1000) / 10, 0)
    myTree.flowerSize = sldTree(8).Value
    
    myTree.fruitMaxChange = IIf(chkTree(9).Value = 1, Abs(Val(txtInput(9).Text) Mod 100), 0)
    myTree.fruitSpeed = sldTree(9).Value
    
    myTree.cutOperation = IIf(chkTree(10).Value = 1, True, False)
    myTree.cutOperationValue = Val(txtInput(10).Text)
    
    onTop = mnuOptions(0).Checked
    fastPaint = mnuOptions(1).Checked
    playSound = mnuOptions(2).Checked
    randomBackground = mnuOptions(3).Checked
    
    ' some calculations are different from last version
    myTree.sizeOfAngel = myTree.sizeOfAngel * 2 / myTree.branchPerStep
    myTree.maxSize = myTree.maxSize - (myTree.trunkHeight / 5 - 8) * 1.6
    If myTree.maxSize < 1 Then myTree.maxSize = 1
    
    'it's ready.
    goRender = True
    If playSound Then sndPlaySound App.Path & "\start.wav", 1
    frmTree.Show
    Me.ZOrder IIf(onTop, 0, 1) ' you can choose that this Form be on the top of the Tree Form or not.
    cmdStop.Visible = True
    cmdOK.Visible = False
    cmdOK.Caption = "Do again !"
End Sub

' stop current Tree rendering
Private Sub cmdStop_Click()
    If playSound Then sndPlaySound App.Path & "\cancel.wav", 1
    goOut = True
End Sub

' finishing timer action in frmTree ( = stop changing flowers to fruit)
Private Sub cmdTimerStop_Click()
    If playSound Then sndPlaySound App.Path & "\cancel.wav", 1
    myTree.fruitCounter = 32760
End Sub

'*************************************
'***** All Menus added in ver 2  *****

' I removed general options from main Form and put them in Menu
Private Sub mnuOptions_Click(Index As Integer)
    mnuOptions(Index).Checked = Not mnuOptions(Index).Checked
End Sub

Private Sub mnuTypesFractals_Click(Index As Integer)
    TreeSelect mnuTypesFractals(Index).Tag
End Sub

Private Sub mnuTypesPlants_Click(Index As Integer)
    TreeSelect mnuTypesPlants(Index).Tag
End Sub

Private Sub mnuTypesUsers_Click(Index As Integer)
    TreeSelect mnuTypesUsers(Index).Tag
End Sub

' restoring data from menu items Tag to Form Objects (text boxes, check boxes, slideres)
' there are several simple ways to do this, but I choose this way .
Private Sub TreeSelect(ByVal itemTag As String)
    On Error Resume Next ' This is for preventing run time errors, when user data in Tree.ini file is not correct.
    Dim myControl As Control
    Dim myDefault() As String
    Dim myItem() As String
    
    ' initially I change my string to array and then I use index of array for my mind.
    myDefault = Split("30,2,1,0,0,0,65,0,35,1,30,50,50,50,50,50,50,50,500,1,20,50,1,20,50,0,20", ",") ' my default numbers for Base Tree, also when user doesn't enter any number
    myItem = Split(itemTag, ",")

    For Each myControl In Controls
        If TypeName(myControl) = "TextBox" Then 'Textbox uses "Text" property but Slider and Chechbox use "Value" property
            myControl.Text = IIf(myItem(myControl.TabIndex) = "", myDefault(myControl.TabIndex), myItem(myControl.TabIndex))
        ElseIf (TypeName(myControl) = "Slider" Or TypeName(myControl) = "CheckBox") Then
            myControl.Value = IIf(myItem(myControl.TabIndex) = "", myDefault(myControl.TabIndex), myItem(myControl.TabIndex))
        End If
    Next
End Sub

' for saving current data profile to Tree.ini file (added in update 1)
Private Sub mnuSaveProfile_Click()
    On Error GoTo errHandler
    Dim myControl As Control
    Dim myFreeFile As Integer, myProfile As String
    Dim myItem(26) As String
    
    'gathering data from controls
    For Each myControl In Controls
        If TypeName(myControl) = "TextBox" Then 'as I wrote later, Textbox uses "Text" property but Slider and Chechbox use "Value" property
            myItem(myControl.TabIndex) = myControl.Text & ""
        ElseIf (TypeName(myControl) = "Slider" Or TypeName(myControl) = "CheckBox") Then
            myItem(myControl.TabIndex) = myControl.Value & ""
        End If
    Next
    myProfile = Join(myItem, ",") ' change array to a string with comma between numbers
    
    myProfile = "New " & Format(Now, "yyyy-mm-dd hh:mm:ss") & "," & myProfile ' make a name for your profile based on current date & time
    
    ' opening Tree.ini and save data
    myFreeFile = FreeFile
    Open App.Path & "\Tree.ini" For Append As #myFreeFile
        Print #myFreeFile, myProfile
    Close #myFreeFile
    
    addMenuItem mnuTypesUsers, myProfile 'add this profile to user menu, right Now.
    
Exit Sub
errHandler:
    MsgBox Err.Description, vbExclamation, "Save Error!"
    
End Sub

'"Save Picture" menu click. (added in update 2)
Private Sub mnuSavePicture_Click()
    Me.ZOrder 1 'send Tree control pannel to back
    DoEvents
    Call frmTree.savePicture
    Me.ZOrder 0 'return Tree control pannel on top
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

' Bye bye
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


