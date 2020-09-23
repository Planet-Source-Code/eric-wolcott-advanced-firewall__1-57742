VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFirewall 
   Caption         =   "Advanced Firewall"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   Icon            =   "frmFirewall.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Firewall Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Access Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Processes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Firewall Logs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rule List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frmMain 
      Caption         =   "Options"
      Height          =   4815
      Index           =   4
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame10 
         Caption         =   "Filter Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   5295
         Begin VB.CheckBox chkMin 
            Caption         =   "Send to Tray When Minimized"
            Height          =   255
            Left            =   2640
            TabIndex        =   104
            Top             =   840
            Width           =   2535
         End
         Begin VB.CheckBox chkDetail 
            Caption         =   "Show Detailed Block Alerts"
            Height          =   255
            Left            =   2640
            TabIndex        =   103
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkExit 
            Caption         =   "Popup Alert Upon Exiting"
            Height          =   255
            Left            =   2640
            TabIndex        =   102
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox chkStartup 
            Caption         =   "Run Adv Firewall on Startup"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox chkPrompt 
            Caption         =   "Prompt Before Blocking"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox chkBlock 
            Caption         =   "Block By Default"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Security Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   65
         Top             =   1560
         Width           =   5295
         Begin VB.CheckBox chkPw 
            Caption         =   "Enable Password Protection"
            Height          =   255
            Left            =   2760
            TabIndex        =   70
            Top             =   120
            Width           =   2415
         End
         Begin MSComctlLib.Slider sld 
            Height          =   495
            Left            =   2640
            TabIndex        =   66
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   873
            _Version        =   393216
            Enabled         =   0   'False
            LargeChange     =   1
            Min             =   1
            Max             =   3
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label lblSld 
            Caption         =   "Password Protects Firewall Enable/Disable Only"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label20 
            Caption         =   "Security Level:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   71
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "High"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   69
            ToolTipText     =   "Password Protects all firewall functions"
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "Low"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   68
            ToolTipText     =   "Password Protects Firewall Enable/Disable Only"
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Medium"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   67
            ToolTipText     =   "Password Protects Firewall Enable/Disable as well as the changing of options and deleting of rules"
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Save/Load Configuration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   57
         Top             =   2880
         Width           =   5295
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse"
            Height          =   255
            Left            =   4320
            TabIndex        =   64
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtSave 
            Height          =   285
            Left            =   1560
            TabIndex        =   63
            Text            =   "C:\firewall.fcg"
            Top             =   840
            Width           =   3615
         End
         Begin VB.CheckBox chkLoad 
            Caption         =   "Automatically Load Settings"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "Automatically Save Settings Upon Exit"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   3015
         End
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Load Settings"
            Height          =   375
            Left            =   1320
            TabIndex        =   59
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save Settings"
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   1215
         End
         Begin MSComDlg.CommonDialog cd 
            Left            =   4680
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label16 
            Caption         =   "Default Save Path:"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   1455
         End
      End
   End
   Begin VB.Frame frmMain 
      Caption         =   "Rule List"
      Height          =   4815
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame5 
         Caption         =   "Process Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   2520
         TabIndex        =   33
         Top             =   1560
         Width           =   2895
         Begin VB.ListBox lstRules 
            Height          =   2400
            Index           =   3
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   2655
         End
         Begin VB.ListBox lstRules 
            Height          =   2400
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.ListBox lstRules 
            Height          =   2400
            Index           =   1
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear List"
            Height          =   255
            Left            =   1440
            TabIndex        =   36
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Rule"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2760
            Width           =   1215
         End
         Begin VB.ListBox lstRules 
            Height          =   2400
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   2295
         Begin VB.Label lblIP 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblRport 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1800
            TabIndex        =   31
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblLport 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1800
            TabIndex        =   30
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblName 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1800
            TabIndex        =   29
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Remote IP Rules:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Remote Port Rules:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Local Port Rules:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Process Name Rules:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Categories"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton opt 
            Caption         =   "Local Port"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "Remote Port"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton opt 
            Caption         =   "Remote IP"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "Process Name"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "New Rule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2520
         TabIndex        =   15
         Top             =   240
         Width           =   2895
         Begin VB.CommandButton cmdAddRule 
            Caption         =   "Add New Rule"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtBlock 
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblBlockType 
            Caption         =   "Block If Process Name Equals:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2295
         End
      End
   End
   Begin VB.Frame frmMain 
      Caption         =   "Firewall Logs"
      Height          =   4815
      Index           =   3
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      Begin TabDlg.SSTab SSTab1 
         Height          =   4455
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7858
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Access Log"
         TabPicture(0)   =   "frmFirewall.frx":1272
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstvwAccess"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame11"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdSaveAccess"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Block Log"
         TabPicture(1)   =   "frmFirewall.frx":128E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdSaveBlock"
         Tab(1).Control(1)=   "Frame12"
         Tab(1).Control(2)=   "lstvwBlock"
         Tab(1).ControlCount=   3
         Begin VB.CommandButton cmdSaveBlock 
            Caption         =   "Save Log File"
            Height          =   375
            Left            =   -74880
            TabIndex        =   87
            Top             =   3300
            Width           =   1215
         End
         Begin VB.Frame Frame12 
            Caption         =   "Block Log Options"
            Height          =   1095
            Left            =   -73560
            TabIndex        =   82
            Top             =   3240
            Width           =   3735
            Begin VB.CheckBox chkSaveBlock 
               Caption         =   "Automatically Save Block Log"
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   240
               Width           =   2535
            End
            Begin VB.TextBox txtSaveBlock 
               Enabled         =   0   'False
               Height          =   285
               Left            =   840
               TabIndex        =   84
               Top             =   600
               Width           =   2775
            End
            Begin VB.CommandButton cmdBrowseBlock 
               Caption         =   "Browse"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2760
               TabIndex        =   83
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label22 
               Caption         =   "Save To:"
               Height          =   255
               Left            =   120
               TabIndex        =   86
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdSaveAccess 
            Caption         =   "Save Log File"
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   3300
            Width           =   1215
         End
         Begin VB.Frame Frame11 
            Caption         =   "Access Log Options"
            Height          =   1095
            Left            =   1440
            TabIndex        =   76
            Top             =   3240
            Width           =   3735
            Begin VB.CommandButton cmdBrowseAccess 
               Caption         =   "Browse"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2760
               TabIndex        =   81
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox txtSaveAccess 
               Enabled         =   0   'False
               Height          =   285
               Left            =   840
               TabIndex        =   79
               Top             =   600
               Width           =   2775
            End
            Begin VB.CheckBox chkSaveAccess 
               Caption         =   "Automatically Save Access Log"
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   240
               Width           =   2655
            End
            Begin VB.Label Label21 
               Caption         =   "Save To:"
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   600
               Width           =   735
            End
         End
         Begin MSComctlLib.ListView lstvwAccess 
            Height          =   2655
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   4683
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lstvwBlock 
            Height          =   2655
            Left            =   -74880
            TabIndex        =   11
            Top             =   480
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   4683
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
   Begin VB.Frame frmMain 
      Caption         =   "Summary"
      Height          =   4815
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      Begin VB.Timer tmrFirewall 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5040
         Top             =   240
      End
      Begin VB.Frame Frame14 
         Caption         =   "Firewall Information"
         Height          =   1935
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   5295
         Begin VB.CommandButton cmdMin 
            Caption         =   "Send to Tray"
            Height          =   375
            Left            =   1560
            TabIndex        =   105
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton cmdEnable 
            Caption         =   "Enable Firewall"
            Height          =   375
            Left            =   240
            TabIndex        =   100
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   2640
            TabIndex        =   99
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblfirewallstat 
            Caption         =   "DISABLED"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   2040
            TabIndex        =   98
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Password Protected:"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   97
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Firewall Status:"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Process Information"
         Height          =   2415
         Left            =   120
         TabIndex        =   88
         Top             =   2280
         Width           =   5295
         Begin VB.Label lblProcNum 
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   2280
            TabIndex        =   94
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Active Processes:"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   93
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblBlocked 
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   2400
            TabIndex        =   92
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Blocked Processes:"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   91
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lblLife 
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   2040
            TabIndex        =   90
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Lifetime Blocked:"
            BeginProperty Font 
               Name            =   "Haettenschweiler"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   89
            Top             =   1800
            Width           =   2055
         End
      End
   End
   Begin VB.Frame frmMain 
      Caption         =   "Processes"
      Height          =   4815
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdGetProcs 
         Caption         =   "List Internet Processes"
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1815
      End
      Begin MSComctlLib.ListView lstvwProc 
         Height          =   2655
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame frmMain 
      Caption         =   "Access Control"
      Height          =   4815
      Index           =   6
      Left            =   2280
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtList 
         Height          =   285
         Left            =   3840
         TabIndex        =   55
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddList 
         Caption         =   "Add To List"
         Height          =   295
         Left            =   4320
         TabIndex        =   53
         Top             =   4320
         Width           =   1095
      End
      Begin VB.ComboBox cmbList 
         Height          =   315
         ItemData        =   "frmFirewall.frx":12AA
         Left            =   960
         List            =   "frmFirewall.frx":12B4
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Frame Frame7 
         Caption         =   "Block List"
         Height          =   3615
         Left            =   2760
         TabIndex        =   44
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton cmdBlockDel 
            Caption         =   "Delete Item"
            Height          =   375
            Left            =   1440
            TabIndex        =   50
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CommandButton cmdBlockClr 
            Caption         =   "Clear Blocked"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   3120
            Width           =   1335
         End
         Begin VB.ListBox lstBlock 
            Height          =   2790
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Allow List"
         Height          =   3615
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   2535
         Begin VB.CommandButton cmdAllowClr 
            Caption         =   "Clear Allowed"
            Height          =   375
            Left            =   1200
            TabIndex        =   48
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdAllowDel 
            Caption         =   "Delete Item"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   3120
            Width           =   1095
         End
         Begin VB.ListBox lstAllow 
            Height          =   2790
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Label Label15 
         Caption         =   "You will have the option of adding a process to these lists when you are prompted to block them or not"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   4320
         Width           =   4095
      End
      Begin VB.Label Label5 
         Caption         =   "Process Name:"
         Height          =   255
         Left            =   2640
         TabIndex        =   54
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label label99 
         Caption         =   "Add To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3960
         Width           =   735
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuControl 
         Caption         =   "Firewall Control"
         Begin VB.Menu mnuStart 
            Caption         =   "Start"
         End
         Begin VB.Menu mnuStop 
            Caption         =   "Stop"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuAccess 
         Caption         =   "Access Control"
         Begin VB.Menu mnuAllow 
            Caption         =   "Add to Allow List"
         End
         Begin VB.Menu mnuBlock 
            Caption         =   "Add To Block List"
         End
      End
      Begin VB.Menu mnuRules 
         Caption         =   "Rules"
         Begin VB.Menu mnuProc 
            Caption         =   "Add Process Name"
         End
         Begin VB.Menu mnuRIP 
            Caption         =   "Add Remote IP"
         End
         Begin VB.Menu mnuRPort 
            Caption         =   "Add Remote Port"
         End
         Begin VB.Menu mnuLPort 
            Caption         =   "Add Local Port"
         End
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmFirewall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub changeApperence(Index As Integer)
    For i = 1 To lblMain.UBound
        If i <> Index Then
            lblMain(i).BackColor = &H8000000F
            lblMain(i).ForeColor = &H80000008
            frmMain(i).Visible = False
        ElseIf i = Index Then
            lblMain(i).BackColor = &H800000
            lblMain(i).ForeColor = &HFFFFFF
            frmMain(i).Visible = True
        End If
    Next i
End Sub

Private Sub chkBlock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tempof = chkBlock.Value
    chkTemp = 0
    If getSecLevel <> 0 Then
        If isAllowed(1) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkBlock.Value = tempof
            Exit Sub
        Else
            If chkTemp = 1 Then
                If tempof = 1 Then
                    chkBlock.Value = 0
                ElseIf tempof = 0 Then
                    chkBlock.Value = 1
                End If
            Else
                chkBlock.Value = tempof
            End If
        End If
    End If
End Sub

Private Sub chkDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
tempof = chkDetail.Value
    chkTemp = 0
    If getSecLevel <> 0 Then
        If isAllowed(1) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkDetail.Value = tempof
            Exit Sub
        Else
            If chkTemp = 1 Then
                If tempof = 1 Then
                    chkDetail.Value = 0
                ElseIf tempof = 0 Then
                    chkDetail.Value = 1
                End If
            Else
                chkDetail.Value = tempof
            End If
        End If
    End If
End Sub

Private Sub chkExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
tempof = chkExit.Value
    chkTemp = 0
    If getSecLevel <> 0 Then
        If isAllowed(1) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkExit.Value = tempof
            Exit Sub
        Else
            If chkTemp = 1 Then
                If tempof = 1 Then
                    chkExit.Value = 0
                ElseIf tempof = 0 Then
                    chkExit.Value = 1
                End If
            Else
                chkExit.Value = tempof
            End If
        End If
    End If
End Sub

Private Sub chkLoad_Click()
Dim result As Long
Dim loadOf As String
Dim loadPath As String
If noShow <> 1 Then
    If chkLoad.Value = 1 Then
        loadOf = "yes"
        cd.DialogTitle = "Load Firewall Configuration On Startup"
        cd.Filter = "Firewall Rules List (*.fcg)|*.fcg|All Files (*.*)|*.*"
        cd.DefaultExt = "fcg"
        cd.ShowOpen
        If cd.FileName <> "" Then
            loadPath = cd.FileName
        Else
            MsgBox "You must specify a filepath"
            chkLoad.Value = 0
        Exit Sub
        End If
    Else
        loadOf = "no"
    End If
    RegCreateKeyEx HKEY_LOCAL_MACHINE, "SOFTWARE\AdvFirewall", 0, "REG_SZ", 0, KEY_ALL_ACCESS, ByVal 0&, result, ret
    RegSetValueEx result, "load", 0, REG_SZ, ByVal loadOf, Len(loadOf)
    RegSetValueEx result, "path", 0, REG_SZ, ByVal loadPath, Len(loadPath)
    RegCloseKey result
End If
End Sub

Private Sub chkLoad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim temp As String
    temp = chkLoad.Value
    If getSecLevel <> 0 Then
        If isAllowed(0) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkLoad.Value = temp
            Exit Sub
        Else
            If temp = 1 Then
                chkLoad.Value = 0
            ElseIf temp = 0 Then
                chkLoad.Value = 1
            End If
        End If
    End If
End Sub

Private Sub chkMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
tempof = chkMin.Value
    chkTemp = 0
    If getSecLevel <> 0 Then
        If isAllowed(1) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkMin.Value = tempof
            Exit Sub
        Else
            If chkTemp = 1 Then
                If tempof = 1 Then
                    chkMin.Value = 0
                ElseIf tempof = 0 Then
                    chkMin.Value = 1
                End If
            Else
                chkMin.Value = tempof
            End If
        End If
    End If
End Sub

Private Sub chkPrompt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tempof = chkPrompt.Value
    chkTemp = 0
    If getSecLevel <> 0 Then
        If isAllowed(1) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkPrompt.Value = tempof
            Exit Sub
        Else
            If chkTemp = 1 Then
                If tempof = 1 Then
                    chkPrompt.Value = 0
                ElseIf tempof = 0 Then
                    chkPrompt.Value = 1
                End If
            Else
                chkPrompt.Value = tempof
            End If
        End If
    End If
End Sub

Private Sub chkPw_Click()
Dim result As Long
Dim pwd As String
Dim pwd2 As String
    cnt = 0
    If chkPw.Value = 0 Then
        sld.Enabled = False
        Label6.Caption = "NO"
        Label6.ForeColor = &HC0&
    ElseIf chkPw.Value = 1 And pwTemp <> 1 Then
        pwd = InputBox("Enter Password", "Password Protection")
        pwd2 = InputBox("Confirm Password", "Password Protection")
        If pwd = pwd2 Then
            If pwd = "" Then
                chkPw.Value = 0
                Exit Sub
            End If
            RegCreateKeyEx HKEY_LOCAL_MACHINE, "SOFTWARE\AdvFirewall", 0, "REG_SZ", 0, KEY_ALL_ACCESS, ByVal 0&, result, ret
            RegSetValueEx result, "pwd", 0, REG_SZ, ByVal pwd, Len(pwd)
            sld.Enabled = True
            Label6.Caption = "YES"
            Label6.ForeColor = &HC000&
            Call loadPWD
            RegCloseKey result
        Else
            MsgBox "Passwords do not match", vbOKOnly, "Password Error"
            chkPw.Value = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub chkPw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim temp As String
    temp = chkPw.Value
    If temp = 0 Then
        Exit Sub
    End If
    If getSecLevel <> 0 Then
        If isAllowed(0) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkPw.Value = temp
            Exit Sub
        Else
            If temp = 1 Then
                chkPw.Value = 0
            ElseIf temp = 0 Then
                chkPw.Value = 1
            End If
        End If
    End If
End Sub

Private Sub chkSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim temp As String
    temp = chkSave.Value
    If getSecLevel <> 0 Then
        If isAllowed(0) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkSave.Value = temp
            Exit Sub
        Else
            If temp = 1 Then
                chkSave.Value = 0
            ElseIf temp = 0 Then
                chkSave.Value = 1
            End If
        End If
    End If
End Sub

Private Sub chkSaveAccess_Click()
    If chkSaveAccess.Value = 1 Then
        cmdBrowseAccess.Enabled = True
        txtSaveAccess.Enabled = True
    Else
        cmdBrowseAccess.Enabled = False
        txtSaveAccess.Enabled = False
    End If
End Sub

Private Sub chkSaveAccess_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tempof = chkSaveAccess.Value
    chkTemp = 0
    If getSecLevel <> 0 Then
        If isAllowed(1) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkSaveAccess.Value = tempof
            Exit Sub
        Else
            If chkTemp = 1 Then
                If tempof = 1 Then
                    chkSaveAccess.Value = 0
                ElseIf tempof = 0 Then
                    chkSaveAccess.Value = 1
                End If
            Else
                chkSaveAccess.Value = tempof
            End If
        End If
    End If
End Sub

Private Sub chkSaveBlock_Click()
    If chkSaveBlock.Value = 1 Then
        cmdBrowseBlock.Enabled = True
        txtSaveBlock.Enabled = True
    Else
        cmdBrowseBlock.Enabled = False
        txtSaveBlock.Enabled = False
    End If
End Sub

Private Sub chkSaveBlock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tempof = chkSaveBlock.Value
    chkTemp = 0
    If getSecLevel <> 0 Then
        If isAllowed(1) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkSaveBlock.Value = tempof
            Exit Sub
        Else
            If chkTemp = 1 Then
                If tempof = 1 Then
                    chkSaveBlock.Value = 0
                ElseIf tempof = 0 Then
                    chkSaveBlock.Value = 1
                End If
            Else
                chkSaveBlock.Value = tempof
            End If
        End If
    End If
End Sub

Private Sub chkStartup_Click()
Dim result As Long
Dim runFile As String

    If chkStartup.Value = 1 Then
        RegCreateKeyEx HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", 0, "REG_SZ", 0, KEY_ALL_ACCESS, ByVal 0&, result, ret
        runFile = App.Path & "\" & App.EXEName & ".exe"
        RegSetValueEx result, "AdvFirewall", 0, REG_SZ, ByVal runFile, Len(runFile)
        RegCloseKey result
    Else
        RegCreateKeyEx HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", 0, "REG_SZ", 0, KEY_ALL_ACCESS, ByVal 0&, result, ret
        RegDeleteValue result, "AdvFirewall"
        RegCloseKey result
    End If
End Sub

Private Sub chkStartup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tempof = chkStartup.Value
    chkTemp = 0
    If getSecLevel <> 0 Then
        If isAllowed(1) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            chkStartup.Value = tempof
            Exit Sub
        Else
            If chkTemp = 1 Then
                If tempof = 1 Then
                    chkStartup.Value = 0
                ElseIf tempof = 0 Then
                    chkStartup.Value = 1
                End If
            Else
                chkStartup.Value = tempof
            End If
        End If
    End If
End Sub

Private Sub cmdAddList_Click()
    If isAllowed(1) = False Then
        MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
        Exit Sub
    End If
    If txtList.Text = "" Then
        MsgBox "You must enter a process name", vbOKOnly, "No Process Name"
        Exit Sub
    End If
    
    For i = 0 To lstAllow.ListCount - 1
       If UCase(txtList.Text) = UCase(lstAllow.List(i)) Then
          test = MsgBox("Process Name Already Used. Remove from list?", vbYesNo, "Already In use")
          If test = vbYes Then
            lstAllow.RemoveItem i
            Exit For
          Else
            Exit Sub
          End If
        End If
    Next i
    
    For i = 0 To lstBlock.ListCount - 1
       If UCase(txtList.Text) = UCase(lstBlock.List(i)) Then
          test = MsgBox("Process Name Already Used. Remove from list?", vbYesNo, "Already In use")
          If test = vbYes Then
            lstBlock.RemoveItem i
            Exit For
          Else
            Exit Sub
          End If
        End If
    Next i
    
    Select Case cmbList.Text
    
        Case "Allowed List"
            lstAllow.AddItem txtList.Text
            txtList.Text = ""
        Case "Blocked List"
            lstBlock.AddItem txtList.Text
            txtList.Text = ""
        Case Else
            MsgBox "You must choose a list to add the process to", vbOKOnly, "No List Selected"
            txtList.Text = ""
            Exit Sub
    End Select
End Sub

Private Sub cmdAddRule_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    If txtBlock.Text = "" Then
        MsgBox "You must enter a value to add a rule", vbOKOnly, "Error"
        Exit Sub
    End If
    For i = 0 To opt.UBound
            If opt(i).Value = True Then
                lstRules(i).AddItem txtBlock.Text
                txtBlock.Text = ""
                Select Case i
                
                Case 0
                    nameRules = nameRules + 1
                Case 1
                    ipRules = ipRules + 1
                Case 2
                    rportRules = rportRules + 1
                Case 3
                    lportRules = lportRules + 1
                End Select
            End If
    Next i
    Call updateStats
End Sub

Private Sub cmdAllowClr_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    lstAllow.Clear
End Sub

Private Sub cmdAllowDel_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    lstAllow.RemoveItem lstAllow.ListIndex
End Sub

Private Sub cmdBlockClr_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    lstBlock.Clear
End Sub

Private Sub cmdBlockDel_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    lstBlock.RemoveItem lstBlock.ListIndex
End Sub

Private Sub cmdBrowse_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    cd.DialogTitle = "Default Save Path..."
    cd.Filter = "Firewall Rules List (*.fcg)|*.fcg|All Files (*.*)|*.*"
    cd.DefaultExt = "fcg"
    cd.ShowSave
    If cd.FileName <> "" Then
        txtSave.Text = cd.FileName
    End If
End Sub

Private Sub cmdBrowseAccess_Click()
If isAllowed(2) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    cd.DialogTitle = "Default Save Path..."
    cd.Filter = "Firewall Access Log (*.fal)|*.fcg|All Files (*.*)|*.*"
    cd.DefaultExt = "fal"
    cd.ShowSave
    txtSaveAccess.Text = cd.FileName
End Sub

Private Sub cmdBrowseBlock_Click()
If isAllowed(2) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    cd.DialogTitle = "Default Save Path..."
    cd.Filter = "Firewall Block Log (*.fbl)|*.fcg|All Files (*.*)|*.*"
    cd.DefaultExt = "fbl"
    cd.ShowSave
    txtSaveBlock.Text = cd.FileName
End Sub

Private Sub cmdClear_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    For i = 0 To opt.UBound
        If opt(i).Value = True Then
            lstRules(i).Clear
            
            Select Case i
            
                Case 0
                    nameRules = 0
                Case 1
                    ipRules = 0
                Case 2
                    rportRules = 0
                Case 3
                    lportRules = 0
            End Select
        End If
    Next i
    Call updateStats
End Sub

Private Sub cmdDelete_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    For i = 0 To opt.UBound
        If opt(i).Value = True Then
            lstRules(i).RemoveItem lstRules(i).ListIndex
            Select Case i
            
            Case 0
                nameRules = nameRules - 1
            Case 1
                ipRules = ipRules - 1
            Case 2
                rportRules = rportRules - 1
            Case 3
                lportRules = lportRules - 1
            End Select
        End If
    Next i
    Call updateStats
End Sub

Private Sub cmdEnable_Click()
If isAllowed(0) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
If tmrFirewall.Enabled = False Then
    If chkPrompt.Value = 1 Then
        promptBlock = 1
    Else
        promptBlock = 0
    End If
    chkPrompt.Enabled = False
    tmrFirewall.Enabled = True
    cmdEnable.Caption = "Disable Firewall"
    lblfirewallstat.Caption = "Enabled"
    lblfirewallstat.ForeColor = &HC000&
    mnuStart.Enabled = False
    mnuStop.Enabled = True
Else
    chkPrompt.Enabled = True
    tmrFirewall.Enabled = False
    cmdEnable.Caption = "Enable Firewall"
    lblfirewallstat.ForeColor = &HC0&
    lblfirewallstat.Caption = "Disabled"
    mnuStart.Enabled = True
    mnuStop.Enabled = False
End If
End Sub

Private Sub cmdGetProcs_Click()
If isAllowed(2) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    notFirewall = 1
    Call RefreshStack
    Call EnumEntries
    notFirewall = 0
End Sub

Private Sub cmdLoad_Click()
If isAllowed(0) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim fcgName As String
Dim linef As String
Dim result As Long
    If noShow = 0 Then
        cd.DialogTitle = "Load Firewall Configuration"
        cd.Filter = "Firewall Rules List (*.fcg)|*.fcg|All Files (*.*)|*.*"
        cd.DefaultExt = "fcg"
        cd.ShowOpen
        fcgName = cd.FileName
    Else
        fcgName = loadPath
    End If
    noShow = 0
    If fcgName <> "" Then
        Open fcgName For Input As #1
        Do While Not EOF(1)
            Line Input #1, linef
             Select Case Left(linef, InStr(1, linef, ":") - 1)
             
                Case "opt"
                    If InStr(1, linef, ";") > 0 Then
                        If InStr(1, linef, ";") = 10 Then
                            txtSave.Text = Right(linef, Len(linef) - InStr(1, linef, ";"))
                        Else
                            sld.Value = Right(linef, 1)
                        End If
                    Else
                        Select Case Right(linef, Len(linef) - InStr(1, linef, ":"))
                            
                            Case "chkBlock"
                                chkBlock.Value = 1
                            Case "chkSave"
                                chkSave.Value = 1
                            Case "chkLoad"
                                chkLoad.Value = 1
                            Case "chkPrompt"
                                chkPrompt.Value = 1
                            Case "chkPw"
                                chkPw.Value = 1
                                Call loadPWD
                            Case "chkSaveBlock"
                                chkSaveBlock.Value = 1
                            Case "chkSaveAccess" = 1
                                chkSaveAccess.Value = 1
                            Case "chkStartup" = 1
                                chkStartup.Value = 1
                            Case "chkExit" = 1
                                chkExit.Value = 1
                            Case "chkDetail" = 1
                                chkDetail.Value = 1
                            
                        End Select
                    End If
                Case "allow"
                    lstAllow.AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case "block"
                    lstBlock.AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case "rule(0)"
                    lstRules(0).AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case "rule(1)"
                    lstRules(1).AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case "rule(2)"
                    lstRules(2).AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case "rule(3)"
                    lstRules(3).AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case "acc"
                    txtSaveAccess.Text = Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case "blk"
                    txtSaveBlock.Text = Right(linef, Len(linef) - InStr(1, linef, ":"))
             End Select
             
        Loop
    End If
    Close #1
    sldLev = sld.Value
End Sub

Private Sub cmdMin_Click()
If isAllowed(2) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = frmFirewall.hwnd
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = frmFirewall.Icon
    TrayI.szTip = trayMsg & Chr$(0)
    Shell_NotifyIcon NIM_ADD, TrayI
    Me.Hide
End Sub

Private Sub cmdSave_Click()
If isAllowed(0) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim fcgName As String
    cd.DialogTitle = "Default Save Path..."
    cd.Filter = "Firewall Rules List (*.fcg)|*.fcg|All Files (*.*)|*.*"
    cd.DefaultExt = "fcg"
    cd.ShowSave
    fcgName = cd.FileName
    If fcgName <> "" Then
        Open fcgName For Append As #1
        Close #1
        Open fcgName For Output As #1
        If chkBlock.Value = 1 Then
            Print #1, "opt:chkBlock"
        End If
        If chkPw.Value = 1 Then
            Print #1, "opt:chkPw"
        End If
        If chkPrompt.Value = 1 Then
            Print #1, "opt:chkPrompt"
        End If
        If chkSave.Value = 1 Then
            Print #1, "opt:chkSave"
        End If
        If chkLoad.Value = 1 Then
            Print #1, "opt:chkLoad"
        End If
        If chkSaveBlock.Value = 1 Then
            Print #1, "opt:chkSaveBlock"
        End If
        If chkSaveAccess.Value = 1 Then
            Print #1, "opt:chkSaveAccess"
        End If
        If chkStartup.Value = 1 Then
            Print #1, "opt:chkStartup"
        End If
        If chkExit.Value = 1 Then
            Print #1, "opt:chkExit"
        End If
        If chkDetail.Value = 1 Then
            Print #1, "opt:chkDetail"
        End If
        Print #1, "opt:seclevel;" & sld.Value
        Print #1, "opt:sPath;" & txtSave.Text
        Print #1, "acc:" & txtSaveAccess.Text
        Print #1, "blk:" & txtSaveBlock.Text
        For i = 0 To 3
            For s = 0 To lstRules(i).ListCount - 1
                Print #1, "rule(" & i & "):" & lstRules(i).List(s)
            Next s
        Next i
        
        For i = 0 To lstAllow.ListCount - 1
            Print #1, "allow:" & lstAllow.List(i)
        Next i
        
        For i = 0 To lstBlock.ListCount - 1
            Print #1, "block:" & lstBlock.List(i)
        Next i
    End If
    Close #1
End Sub

Private Sub cmdSaveAccess_Click()
If isAllowed(2) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim fcgName As String
    cd.DialogTitle = "Save Access Log..."
    cd.Filter = "Firewall Access Logs (*.fal)|*.fal|All Files (*.*)|*.*"
    cd.DefaultExt = "fal"
    cd.ShowSave
    fcgName = cd.FileName
    If fcgName <> "" Then
        Open fcgName For Append As #1
        Close #1
        Open fcgName For Output As #1
        Print #1, "[Access Log. For Use With Advanced Firewall]"
        Print #1, "[ProcessName][ProcessID][LocalAddress][LocalPort][RemoteAddress][RemotePort][Attempts][Time]"
        For i = 1 To lstvwAccess.ListItems.Count - 1
            Print #1, lstvwAccess.ListItems(i).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(1).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(2).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(3).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(4).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(5).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(6).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(7).Text & ";" & Date & " " & Time
        Next i
        Close #1
    End If
End Sub

Private Sub cmdSaveBlock_Click()
If isAllowed(2) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim fcgName As String
    cd.DialogTitle = "Save Block Log..."
    cd.Filter = "Firewall Block Logs (*.fbl)|*.fbl|All Files (*.*)|*.*"
    cd.DefaultExt = "fbl"
    cd.ShowSave
    fcgName = cd.FileName
    If fcgName <> "" Then
        Open fcgName For Append As #1
        Close #1
        Open fcgName For Output As #1
        Print #1, "[Block Log. For Use With Advanced Firewall]"
        Print #1, "[ProcessName][ProcessID][LocalAddress][LocalPort][RemoteAddress][RemotePort][Attempts][Time]"
        For i = 1 To lstvwBlock.ListItems.Count
            Print #1, lstvwBlock.ListItems(i).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(1).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(2).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(3).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(4).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(5).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(6).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(7).Text & ";" & Date & " " & Time
        Next i
        Close #1
    End If
End Sub

Private Sub Form_Load()
    noShow = 0
    pwTemp = 0
    sldLev = 0
    nameRules = 0
    procNum = 0
    block = 1
    blocked = 0
    ipRules = 0
    rportRules = 0
    lportRules = 0
    Call readyLogView
    Call loadLife
    If loadCheck = True Then
        noShow = 1
        chkLoad.Value = 1
        cmdLoad_Click
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mg = X / Screen.TwipsPerPixelX
    If mg = WM_LBUTTONDBLCLK Then
        Me.Show
        TrayI.cbSize = Len(TrayI)
        TrayI.hwnd = frmFirewall.hwnd
        TrayI.uId = 1&
        Shell_NotifyIcon NIM_DELETE, TrayI
    ElseIf mg = WM_RBUTTONUP Then
        Me.PopupMenu mnuMain
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tempRep As String
    If chkExit.Value = 1 Then
        tempRep = MsgBox("Are you sure you want to exit AdvFirewall?", vbYesNo, "Alert")
        If tempRep = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    If chkSaveBlock.Value = 1 Then
        Call writeBlock
    End If
    If chkSaveAccess.Value = 1 Then
        Call writeAccess
    End If
    tmrFirewall.Enabled = False
    Call updateLifeBlock
    If chkSave.Value = 1 Then
        Call cmdSave_Click
    End If
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = frmFirewall.hwnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
    End
End Sub

Private Sub lblMain_Click(Index As Integer)
    Call changeApperence(Index)
End Sub

Private Sub readyLogView()
    'access
    Set colHead = lstvwAccess.ColumnHeaders.Add(lstvwAccess.ColumnHeaders.Count + 1, "Process Name", "Process Name", TextWidth("Process Name") * 1.5)
    Set colHead = lstvwAccess.ColumnHeaders.Add(lstvwAccess.ColumnHeaders.Count + 1, "Process ID", "Process ID", TextWidth("Process ID") * 1.3)
    Set colHead = lstvwAccess.ColumnHeaders.Add(lstvwAccess.ColumnHeaders.Count + 1, "Local Address", "Local Address", TextWidth("Local Address") * 1.5)
    Set colHead = lstvwAccess.ColumnHeaders.Add(lstvwAccess.ColumnHeaders.Count + 1, "Local Port", "Local Port", TextWidth("Local Port") * 1.5)
    Set colHead = lstvwAccess.ColumnHeaders.Add(lstvwAccess.ColumnHeaders.Count + 1, "Remote Address", "Remote Address", TextWidth("Remote Address") * 1.5)
    Set colHead = lstvwAccess.ColumnHeaders.Add(lstvwAccess.ColumnHeaders.Count + 1, "Remote Port", "Remote Port", TextWidth("Remote Port") * 1.3)
    Set colHead = lstvwAccess.ColumnHeaders.Add(lstvwAccess.ColumnHeaders.Count + 1, "Attempts", "Attempts", TextWidth("Attempts") * 1.5)
    Set colHead = lstvwAccess.ColumnHeaders.Add(lstvwAccess.ColumnHeaders.Count + 1, "Time", "Time", TextWidth("Time") * 5)
    'block
    Set colHead = lstvwBlock.ColumnHeaders.Add(lstvwBlock.ColumnHeaders.Count + 1, "Process Name", "Process Name", TextWidth("Process Name") * 1.5)
    Set colHead = lstvwBlock.ColumnHeaders.Add(lstvwBlock.ColumnHeaders.Count + 1, "Process ID", "Process ID", TextWidth("Process ID") * 1.3)
    Set colHead = lstvwBlock.ColumnHeaders.Add(lstvwBlock.ColumnHeaders.Count + 1, "Local Address", "Local Address", TextWidth("Local Address") * 1.5)
    Set colHead = lstvwBlock.ColumnHeaders.Add(lstvwBlock.ColumnHeaders.Count + 1, "Local Port", "Local Port", TextWidth("Local Port") * 1.5)
    Set colHead = lstvwBlock.ColumnHeaders.Add(lstvwBlock.ColumnHeaders.Count + 1, "Remote Address", "Remote Address", TextWidth("Remote Address") * 1.5)
    Set colHead = lstvwBlock.ColumnHeaders.Add(lstvwBlock.ColumnHeaders.Count + 1, "Remote Port", "Remote Port", TextWidth("Remote Port") * 1.3)
    Set colHead = lstvwBlock.ColumnHeaders.Add(lstvwBlock.ColumnHeaders.Count + 1, "Attempts", "Attempts", TextWidth("Attempts") * 1.5)
    Set colHead = lstvwBlock.ColumnHeaders.Add(lstvwBlock.ColumnHeaders.Count + 1, "Time", "Time", TextWidth("Time") * 5)
    'Processes
    Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "Process Name", "Process Name", TextWidth("Process Name") * 1.5)
    Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "Process ID", "Process ID", TextWidth("Process ID") * 1.3)
    Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "Local Address", "Local Address", TextWidth("Local Address") * 1.5)
    Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "Local Port", "Local Port", TextWidth("Local Port") * 1.5)
    Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "Remote Address", "Remote Address", TextWidth("Remote Address") * 1.5)
    Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "Remote Port", "Remote Port", TextWidth("Remote Port") * 1.3)
    Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "State", "State", TextWidth("State") * 4)
End Sub

Private Sub mnuAllow_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim addName As String
    addName = InputBox("Enter Process Name:", "Add To Allow List")
    If addName <> "" Then
        For i = 0 To lstAllow.ListCount - 1
           If UCase(txtList.Text) = UCase(lstAllow.List(i)) Then
              test = MsgBox("Process Name Already Used. Remove from list?", vbYesNo, "Already In use")
              If test = vbYes Then
                lstAllow.RemoveItem i
                Exit For
              Else
                Exit Sub
              End If
            End If
        Next i
    Else
        MsgBox "Invalid Process Name", vbOKOnly, "Error"
        Exit Sub
    End If
    lstAllow.AddItem addName
End Sub

Private Sub mnuBlock_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim addName As String
    addName = InputBox("Enter Process Name:", "Add To Block List")
    If addName <> "" Then
        For i = 0 To lstBlock.ListCount - 1
           If UCase(addName) = UCase(lstBlock.List(i)) Then
              test = MsgBox("Process Name Already Used. Remove from list?", vbYesNo, "Already In use")
              If test = vbYes Then
                lstBlock.RemoveItem i
                Exit For
              Else
                Exit Sub
              End If
            End If
        Next i
    Else
        MsgBox "Invalid Process Name", vbOKOnly, "Error"
        Exit Sub
    End If
    lstBlock.AddItem addName
End Sub

Private Sub mnuExit_Click()
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = frmFirewall.hwnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
    End
End Sub

Private Sub mnuLPort_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim addName As String
    addName = InputBox("Enter Local Port:", "Add Rule")
    If addName = "0" Or addName = "" Then
        MsgBox "Invalid Port", vbOKOnly, "Error"
        Exit Sub
    End If
    lstRules(3).AddItem addName
    lportRules = lportRules + 1
    Call updateStats
End Sub

Private Sub mnuOpen_Click()
    Me.Show
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = frmFirewall.hwnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
    Me.WindowState = 0
End Sub

Private Sub mnuProc_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim addName As String
    addName = InputBox("Enter Process Name:", "Add Rule")
    If addName = "" Then
        MsgBox "Invalid Process Name", vbOKOnly, "Error"
        Exit Sub
    End If
    lstRules(0).AddItem addName
    nameRules = nameRules + 1
    Call updateStats
End Sub

Private Sub mnuRIP_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim addName As String
    addName = InputBox("Enter Remote IP:", "Add Rule")
    If addName = "0" Or addName = "" Then
        MsgBox "Invalid IP", vbOKOnly, "Error"
        Exit Sub
    End If
    lstRules(1).AddItem addName
    ipRules = ipRules + 1
    Call updateStats
End Sub

Private Sub mnuRPort_Click()
If isAllowed(1) = False Then
    MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
    Exit Sub
End If
Dim addName As String
    addName = InputBox("Enter Remote Port:", "Add Rule")
    If addName = "0" Or addName = "" Then
        MsgBox "Invalid Port", vbOKOnly, "Error"
        Exit Sub
    End If
    lstRules(2).AddItem addName
    rportRules = rportRules + 1
    Call updateStats
End Sub

Private Sub mnuStart_Click()
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = frmFirewall.hwnd
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = frmAlert.Icon
    TrayI.szTip = trayMsg & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, TrayI
    Call cmdEnable_Click
End Sub

Private Sub mnuStop_Click()
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = frmFirewall.hwnd
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = frmFirewall.Icon
    TrayI.szTip = trayMsg & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, TrayI
    Call cmdEnable_Click
End Sub

Private Sub opt_Click(Index As Integer)
    Select Case Index
    
        Case 0
            lblBlockType.Caption = "Block If Process Name Equals:"
            Frame5.Caption = "Process Name Rules"
        Case 1
            lblBlockType.Caption = "Block If Remote IP Equals:"
            Frame5.Caption = "Remote IP Rules"
        Case 2
            lblBlockType.Caption = "Block If Remote Port Equals:"
            Frame5.Caption = "Remote Port Rules"
        Case 3
            lblBlockType.Caption = "Block If Local Port Equals:"
            Frame5.Caption = "Local Port Rules"
            
    End Select
    
    For i = 0 To opt.UBound
        If i = Index Then
            lstRules(i).Visible = True
        ElseIf i <> Index Then
            lstRules(i).Visible = False
        End If
    Next i
End Sub

Private Sub updateStats()
    lblName.Caption = nameRules
    lblIP.Caption = ipRules
    lblRport.Caption = rportRules
    lblLport.Caption = lportRules
End Sub

Private Sub sld_Change()
    If sld.Value = 1 Then
        lblSld.Caption = "Password Protects Firewall Enable/Disable Only"
    ElseIf sld.Value = 2 Then
        lblSld.Caption = "Password Protects Firewall Enable/Disable as well as the changing of options and deleting of rules"
    ElseIf sld.Value = 3 Then
        lblSld.Caption = "Password Protects all firewall functions"
    End If
    If sldtemp <> 0 Then
        sld.Value = sldtemp
    End If
    sldLev = sld.Value
End Sub

Private Sub sld_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sldtemp = sld.Value
    If getSecLevel <> 0 Then
        If isAllowed(0) = False Then
            MsgBox "Invalid Password", vbOKOnly, "Password Invalid"
            Exit Sub
        Else
            sldtemp = 0
        End If
    Else
        sldtemp = 0
    End If
End Sub

Private Sub tmrFirewall_Timer()
    Call RefreshStack
    Call EnumEntries
End Sub

Public Sub updateLifeBlock()
Dim blockedNum As String
Dim tempNamex As String
Dim result As Long
    tempNamex = lblLife.Caption
    blockedNum = Val(tempNamex)
    RegCreateKeyEx HKEY_LOCAL_MACHINE, "SOFTWARE\AdvFirewall", 0, "REG_SZ", 0, KEY_ALL_ACCESS, ByVal 0&, result, ret
    RegSetValueEx result, "blocked", 0, REG_SZ, ByVal blockedNum, Len(blockedNum)
    RegCloseKey result
End Sub

Public Sub loadLife()
    blockedLife = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\AdvFirewall\", "blocked")
    lblLife.Caption = blockedLife
End Sub
