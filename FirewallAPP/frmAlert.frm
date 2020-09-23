VERSION 5.00
Begin VB.Form frmAlert 
   Caption         =   "Access Attempt"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optAllow 
      Caption         =   "Add to Allow List"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.OptionButton optBlock 
      Alignment       =   1  'Right Justify
      Caption         =   "Add to block list"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Allow"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "Block"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblLIP 
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLPort 
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblRPort 
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblRIP 
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Is trying to access the internet."
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Local Port:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Local IP:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Remote Port:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblFname 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNo_Click()
    If optBlock.Value = True Then
        frmFirewall.lstBlock.AddItem lblFname.Caption
        blockAlert = 1
    Else
        blockAlert = 1
    End If
    holdLoop = 1
    Unload frmAlert
End Sub

Private Sub cmdYes_Click()
    If optAllow.Value = True Then
        frmFirewall.lstAllow.AddItem lblFname.Caption
        blockAlert = 0
    Else
        blockAlert = 0
    End If
    holdLoop = 1
    Unload frmAlert
End Sub
