VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert IT"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "On Top"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   6600
      Picture         =   "Form1.frx":08E1
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Set Top-Most Window"
      Top             =   0
      Value           =   1  'Checked
      Width           =   915
   End
   Begin ConvertIT.cmdopen CD1 
      Left            =   8880
      Top             =   2160
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   8880
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8880
      Top             =   240
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4440
      TabIndex        =   3
      Top             =   3240
      Width           =   3135
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stereo"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2160
         TabIndex        =   26
         Top             =   1440
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mono"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1200
         TabIndex        =   25
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Combo6 
         Height          =   330
         Left            =   1200
         TabIndex        =   23
         Text            =   "Combo6"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         Height          =   330
         Left            =   1200
         TabIndex        =   22
         Text            =   "Combo5"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         Height          =   330
         Left            =   1200
         TabIndex        =   21
         Text            =   "Combo4"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Channels"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Samples"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Codec"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Audio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   615
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   3120
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   0
         Y1              =   2160
         Y2              =   0
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00FFC0C0&
         X1              =   3120
         X2              =   3120
         Y1              =   0
         Y2              =   2160
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   3135
      Begin VB.ComboBox Combo9 
         Height          =   330
         Left            =   1200
         TabIndex        =   40
         Text            =   "Combo9"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "DeInterlace"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1800
         TabIndex        =   37
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SameQuality"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   360
         TabIndex        =   36
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox Combo8 
         Height          =   330
         Left            =   1200
         TabIndex        =   33
         Text            =   "Combo8"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo7 
         Height          =   330
         Left            =   1200
         TabIndex        =   24
         Text            =   "Combo7"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         Height          =   330
         Left            =   1200
         TabIndex        =   20
         Text            =   "Combo3"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   330
         Left            =   1200
         TabIndex        =   19
         Text            =   "Combo2"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Extension"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Framerate"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Codec"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2400
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00FFC0C0&
         X1              =   3120
         X2              =   3120
         Y1              =   0
         Y2              =   2400
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   3120
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Video"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4215
      Begin ConvertIT.Button Command2 
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         ToolTipText     =   "Browse For Video"
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":11AB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   4200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFC0C0&
         X1              =   4200
         X2              =   4200
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   4200
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Input"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Input Video"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   4215
      Begin VB.Line Line11 
         BorderColor     =   &H00FFC0C0&
         X1              =   4200
         X2              =   4200
         Y1              =   0
         Y2              =   1440
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1440
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   4320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   4200
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Output"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Output Directory"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   4215
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   720
         Width           =   3735
      End
      Begin ConvertIT.Button Command3 
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":11C7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ConvertIT.Button Command1 
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Convert"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":11E3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ConvertIT.Button Button1 
         Height          =   255
         Left            =   240
         TabIndex        =   41
         ToolTipText     =   "Open Media Directory"
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":11FF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FFC0C0&
         X1              =   4200
         X2              =   4200
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   4200
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   4200
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Convert"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   3255
      End
   End
   Begin ConvertIT.Button Button2 
      Height          =   735
      Left            =   5400
      TabIndex        =   43
      ToolTipText     =   "View Converted Video"
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "View it"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "Form1.frx":121B
      PICN            =   "Form1.frx":1237
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":1689
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Convert IT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   7680
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5280
   End
   Begin VB.Image Image1 
      Height          =   8625
      Left            =   0
      Picture         =   "Form1.frx":1F6A
      Top             =   -360
      Width           =   10200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IntCaptn As Integer
Private WithEvents AVC1 As AVC
Attribute AVC1.VB_VarHelpID = -1
Private shlShell As Shell32.Shell
Private shlFolder As Shell32.Folder
Dim GlobalFileName As String

Private Sub Button1_Click()
      If shlShell Is Nothing Then
          Set shlShell = New Shell32.Shell
      End If
    shlShell.Explore (App.Path & "\Media")
End Sub



Private Sub Button2_Click()
    Dim ngReturnNumber As Long
    If Trim(GlobalFileName) = "" Then Exit Sub
    ngReturnNumber = ShellExecLaunchFile(GlobalFileName, "", App.Path)
End Sub

Private Sub Check3_Click()
On Error Resume Next
If Check3.Value = 1 Then
    SetTopMostWindow Me.hWnd, True
    Open App.Path & "\OnTop" For Output As #1: Close #1
Else
    SetTopMostWindow Me.hWnd, False
    Kill App.Path & "\OnTop"
End If
Me.Width = 7815
Me.Height = 5625
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Command3_Click()
On Error Resume Next
AVC1.CancelConvert
Command3.Enabled = False
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Text1.Text = "" Then Command2_Click: Exit Sub
If Text2.Text = "" Then Command2_Click: Exit Sub
Command1.Enabled = False
Command2.Enabled = False
Command3.Visible = True
Command3.Enabled = True
ConvertFLV (Combo1.Text)
End Sub

Private Sub ConvertFLV(FormatType As String)
On Error Resume Next
Dim OutFileExt As String, OutFileName As String, Bagel As String, Bagel1 As String, Extn As String, i As Integer
OutFileExt = ""
OutFileName = ""
Select Case FormatType

Case Is = "Animated GIF (64x48) GIF"
 OutFileExt = ".gif"
 AVC1.VideoSize = "64x48"
 AVC1.ForceFormat = "gif"
Case Is = "Animated GIF (160x120) GIF"
 OutFileExt = ".gif"
 AVC1.VideoSize = "160x120"
 AVC1.ForceFormat = "gif"
Case Is = "Animated GIF (320x240) GIF"
 OutFileExt = ".gif"
 AVC1.VideoSize = "320x240"
 AVC1.ForceFormat = "gif"
Case Is = "Animated GIF (Same as Input) GIF"
 OutFileExt = ".gif"
 'AVC1.VideoSize = "320x240"
 AVC1.ForceFormat = "gif"
 'AVC1.AudioCodec = "pcm_s16le"
 'AVC1.AudioChannels = "2"
 'AVC1.AudioBitrate = "128"
 'AVC1.AudioSamples = "41000"
Case Is = "WAVE Sound (WAV)"
 OutFileExt = ".wav"
 AVC1.ForceFormat = "wav"
 AVC1.AudioCodec = "pcm_s16le"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "128"
 AVC1.AudioSamples = "41000"
Case Is = "MP3 Sound (MP3)"
 OutFileExt = ".mp3"
 AVC1.ForceFormat = "mp3"
 AVC1.AudioCodec = "mp3"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "128"
 AVC1.AudioSamples = "44100"
Case Is = "3G2 (for Mobile)"
 OutFileExt = ".3g2"
 AVC1.VideoSize = "176x144"
 AVC1.VideoBitrateTolerance = "60"
 AVC1.VideoFrameRate = "11"
 AVC1.VideoBitrate = "120"
 AVC1.VideoCodec = "mpeg4"
 AVC1.AudioCodec = "amr_nb"
 AVC1.AudioChannels = "1"
 AVC1.AudioBitrate = "48"
 AVC1.AudioSamples = "8000"
Case Is = "3GP (for Mobile)"
 OutFileExt = ".3gp"
 AVC1.VideoSize = "176x144"
 AVC1.VideoBitrateTolerance = "60"
 AVC1.VideoFrameRate = "11"
 AVC1.VideoBitrate = "120"
 AVC1.VideoCodec = "h263"
 AVC1.AudioCodec = "amr_nb"
 AVC1.AudioChannels = "1"
 AVC1.AudioBitrate = "48"
 AVC1.AudioSamples = "8000"
Case Is = "Zune WMV"
 OutFileExt = ".wmv"
 AVC1.VideoSize = "320x240"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "1500"
 AVC1.VideoCodec = "wmv2"
 AVC1.AudioCodec = "mp3"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "128"
 AVC1.AudioSamples = "48000"
Case Is = "Zune MPEG-4 AVC"
 OutFileExt = ".avi"
 AVC1.VideoSize = "320x240"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "1500"
 AVC1.VideoCodec = "h264"
 AVC1.GroupOfPictureSize = "250"
 AVC1.VideoQuantiserScale = "25"
 AVC1.MaxVideoBitrate = "1500"
 AVC1.RateControlBuffer = "128"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "64"
 AVC1.AudioSamples = "48000"
Case Is = "Zune MPEG-4 Video"
 OutFileExt = ".avi"
 AVC1.VideoSize = "320x240"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "1500"
 AVC1.VideoCodec = "mpeg4"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "128"
 AVC1.AudioSamples = "48000"
Case Is = "Sony PSP (PSP AVC)"
 OutFileExt = ".avi"
 AVC1.VideoSize = "368x208"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "400"
 AVC1.VideoCodec = "h264"
 AVC1.GroupOfPictureSize = "250"
 AVC1.VideoQuantiserScale = "25"
 AVC1.MaxVideoBitrate = "1500"
 AVC1.RateControlBuffer = "128"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "128"
 AVC1.AudioSamples = "48000"
Case Is = "Sony PSP (Minimal Size)"
 OutFileExt = ".avi"
 AVC1.VideoSize = "160x112"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "216"
 AVC1.VideoCodec = "xvid"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "64"
 AVC1.AudioSamples = "24000"
Case Is = "Sony PSP (Excellent Quality)"
 OutFileExt = ".avi"
 AVC1.VideoSize = "368x208"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "768"
 AVC1.VideoCodec = "xvid"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "64"
 AVC1.AudioSamples = "24000"
Case Is = "Sony PSP (Normal)"
 OutFileExt = ".avi"
 AVC1.VideoSize = "368x208"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "512"
 AVC1.VideoCodec = "xvid"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "64"
 AVC1.AudioSamples = "24000"
Case Is = "QuickTime MOV"
 OutFileExt = ".mov"
 AVC1.VideoSize = "320x240"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "360"
 AVC1.DeInterlace = True
 AVC1.SameQuality = True
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "48"
 AVC1.AudioSamples = "44100"
Case Is = "RM (RealONE/RealPlayer)"
 OutFileExt = ".rm"
 AVC1.VideoCodec = "rv20"
 AVC1.DeInterlace = True
 AVC1.SameQuality = True
 AVC1.AudioChannels = "2"
 AVC1.AudioSamples = "44100"
Case Is = "WMV(WindowsMedia-Best Quality)"
 OutFileExt = ".wmv"
 AVC1.VideoCodec = "wmv2"
 AVC1.AudioChannels = "2"
 AVC1.SameQuality = True
 AVC1.DeInterlace = True
Case Is = "WMV(WindowsMedia-Compatible)"
 OutFileExt = ".wmv"
 AVC1.VideoCodec = "wmv2"
 AVC1.AudioChannels = "2"
 AVC1.AudioSamples = "44100"
 AVC1.DeInterlace = True
 AVC1.SameQuality = True
Case Is = "MPEG (Best Quality)"
 OutFileExt = ".mpg"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoCodec = "mpeg1video"
 AVC1.AudioChannels = "2"
 AVC1.SameQuality = True
 AVC1.DeInterlace = True
Case Is = "MPEG (Max Compatible)"
 OutFileExt = ".mpg"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoCodec = "mpeg1video"
 AVC1.AudioChannels = "2"
 AVC1.AudioSamples = "44100"
 AVC1.SameQuality = True
 AVC1.DeInterlace = True
Case Is = "iPOD Nano (MP3)"
 OutFileExt = ".mp3"
 AVC1.AudioCodec = "mp3"
Case Is = "iPOD Mini (MP3)"
 OutFileExt = ".mp3"
 AVC1.AudioCodec = "mp3"
Case Is = "iPOD Shuffle (MP3)"
 OutFileExt = ".mp3"
 AVC1.AudioCodec = "mp3"
Case Is = "iPhone Video"
 OutFileExt = ".mpg"
 AVC1.VideoSize = "320x240"
 AVC1.VideoAspectRatio = "4:3"
 AVC1.VideoBitrate = "1200"
 AVC1.VideoCodec = "mpeg4"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "80"
 AVC1.AudioSamples = "44100"
Case Is = "iPOD Video2 (640x480) MPEG4"
 OutFileExt = ".mpg"
 AVC1.VideoSize = "640x480"
 AVC1.VideoAspectRatio = "4:3"
 AVC1.VideoBitrate = "1200"
 AVC1.VideoCodec = "mpeg4"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "80"
 AVC1.AudioSamples = "44100"
Case Is = "iPOD Video (320x240) MPEG4"
 OutFileExt = ".mpg"
 AVC1.VideoSize = "320x240"
 AVC1.VideoAspectRatio = "4:3"
 AVC1.VideoBitrate = "512"
 AVC1.VideoCodec = "mpeg4"
 AVC1.AudioCodec = "aac"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "64"
 AVC1.AudioSamples = "44100"
Case Is = "AVI (MPEG4-XviD-Best Quality)"
 OutFileExt = ".avi"
 AVC1.VideoSize = "320x240"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "360"
 AVC1.VideoCodec = "xvid"
 AVC1.AudioChannels = "2"
 AVC1.SameQuality = True
 AVC1.DeInterlace = True
Case Is = "AVI (MPEG4-XviD-Compatible)"
 OutFileExt = ".avi"
 AVC1.VideoSize = "320x240"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "360"
 AVC1.VideoCodec = "xvid"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "48"
 AVC1.AudioSamples = "22050"
 AVC1.SameQuality = True
 AVC1.DeInterlace = True
Case Is = "AVI (MPEG4-DivX-Best Quality)"
 OutFileExt = ".avi"
 AVC1.VideoSize = "320x240"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "360"
 AVC1.VideoCodec = "mpeg4"
 AVC1.AudioChannels = "2"
 AVC1.SameQuality = True
 AVC1.DeInterlace = True
Case Is = "AVI (MPEG4-DivX-Compatible)"
 OutFileExt = ".avi"
 AVC1.VideoSize = "320x240"
 AVC1.VideoFrameRate = "29.97"
 AVC1.VideoBitrate = "360"
 AVC1.VideoCodec = "mpeg4"
 AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = "48"
 AVC1.AudioSamples = "22050"
 AVC1.SameQuality = True
 AVC1.DeInterlace = True
 Case Is = "FLV (Flash Live Video)"
 OutFileExt = ".flv"
 AVC1.VideoSize = "352x288"
 AVC1.VideoFrameRate = "25"
 AVC1.VideoBitrate = "360"
 AVC1.VideoCodec = "h263"
 AVC1.AudioChannels = "1"
 AVC1.AudioBitrate = "64"
 AVC1.AudioCodec = "mp3"
 AVC1.AudioSamples = "22050"
Case Is = "Custom Format"
 OutFileExt = Combo9.Text
 AVC1.VideoSize = Combo8.Text
 AVC1.VideoFrameRate = Combo7.Text
 AVC1.VideoBitrate = Combo3.Text
 AVC1.VideoCodec = Combo2.Text
 If Option1.Value = True Then AVC1.AudioChannels = "1"
 If Option2.Value = True Then AVC1.AudioChannels = "2"
 AVC1.AudioBitrate = Combo5.Text
 AVC1.AudioSamples = Combo6.Text
 If Check1 = 1 Then AVC1.SameQuality = True
 If Check2 = 1 Then AVC1.DeInterlace = True
Case Else
MsgBox "Error: " & FormatType & " is not defined.", vbExclamation, "Not Supported"
Exit Sub
End Select
OutFileName = Mid(CD1.cFileTitle(1), 1, Len(CD1.cFileTitle(1)) - 4) & OutFileExt
AVC1.SourceFile = Text1.Text
i = 0
Bagel = Replace(OutFileName, OutFileExt, "")
Bagel1 = Bagel
Schnoz:
    If Dir(Text2.Text & Trim(Bagel1) & OutFileExt) <> "" Then
        Bagel1 = Trim(Bagel) & Trim(Str(i))
        i = i + 1
        GoTo Schnoz
    Else
        OutFileName = Trim(Bagel1) & OutFileExt
        'MsgBox OutFileName
    End If
AVC1.DestFile = Text2.Text & OutFileName
GlobalFileName = Text2.Text & OutFileName
AVC1.ConvertMedia True
End Sub

Private Sub Command2_Click()
On Error Resume Next
CD1.DialogTitle = "Select File To Convert"
'CD1.Filter = "WMV (*.wmv)|*.wmv|FLV (*.flv)|*.flv|MPG (*.mpg)|*.mpg|AVI (*.avi)|*.avi|MPEG (*.mpeg)|*.mpeg|MOV (*.mov)|*.mov|WAVE (*.wav)|*.wav|MP3 (*.mp3)|*.mp3|All Files (*.*)|*.*"
CD1.Filter = "All Graphic" & Chr$(0) & "*.wmv;*.flv;*.avi;*.mpg;*.mpeg;*.mov;*.wav;*.mp3" & Chr$(0) & _
        "AVI (*.avi)" & Chr$(0) & "*.avi" & Chr$(0) & "WMV (*.wmv)" & Chr$(0) & "*.wmv" & Chr$(0) & "FLV (*.flv)" & Chr$(0) & "*.flv" & Chr$(0) & "MPG (*.mpg)" & Chr$(0) & "*.mpg" & Chr$(0) & "MPEG (*.mpeg)" & Chr$(0) & "*.mpeg" & Chr$(0) & "MOV (*.mov)" & Chr$(0) & "*.mov" & Chr$(0) & "WAVE (*.wav)" & Chr$(0) & "*.wav" & Chr$(0) & "MP3 (*.mp3)" & Chr$(0) & "*.mp3"
CD1.FilterIndex = 1

CD1.ShowOpen
If CD1.cFileTitle(1) = "" Then Exit Sub
Text1.Locked = False
Text1.Text = CD1.cFileName(1)
Text1.Locked = True
Label9.Caption = "Loaded " & CD1.cFileTitle(1)
End Sub



Private Sub Form_Initialize()
Me.Width = 7815
Me.Height = 5625
End Sub

Private Sub Form_Load()
On Error Resume Next
Set AVC1 = New AVC
Button2.BackColor = &HFFFFFF
Button2.BackOver = &HFFFFFF
IntCaptn = 0
LoadCombos
Text2.Text = App.Path & "\Media\"
If Dir(Text2.Text, vbDirectory) = "" Then MkDir (Text2.Text)
Combo1.Text = GetSetting("ConvertIT", "Options", "Op1", "AVI (MPEG4-XviD-Best Quality)")
Combo2.Text = GetSetting("ConvertIT", "Options", "Op2", "mpeg4")
Combo3.Text = GetSetting("ConvertIT", "Options", "Op3", "360")
Combo7.Text = GetSetting("ConvertIT", "Options", "Op7", "29.97")
Combo8.Text = GetSetting("ConvertIT", "Options", "Op8", "320x240")
Combo4.Text = GetSetting("ConvertIT", "Options", "Op4", "mp2")
Combo5.Text = GetSetting("ConvertIT", "Options", "Op5", "128")
Combo6.Text = GetSetting("ConvertIT", "Options", "Op6", "44100")
Check1.Value = GetSetting("ConvertIT", "Options", "Op9", False)
Check2.Value = GetSetting("ConvertIT", "Options", "Op10", False)
Option1.Value = GetSetting("ConvertIT", "Options", "Op11", False)
Option2.Value = GetSetting("ConvertIT", "Options", "Op12", True)
Combo9.Text = GetSetting("ConvertIT", "Options", "Op13", ".avi")
If Combo9.Text = "" Then Combo9.Text = ".avi"
EnableCombos
If Dir(App.Path & "\OnTop") <> "" Then
    SetTopMostWindow Me.hWnd, True
    Check3.Value = 1
Else
    SetTopMostWindow Me.hWnd, False
    Check3.Value = 0
End If
If Dir(App.Path & "\Media", vbDirectory) = "" Then
    MkDir App.Path & "\Media"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set AVC1 = Nothing
SaveSetting "ConvertIT", "Options", "Op1", Combo1.Text
SaveSetting "ConvertIT", "Options", "Op2", Combo2.Text
SaveSetting "ConvertIT", "Options", "Op3", Combo3.Text
SaveSetting "ConvertIT", "Options", "Op4", Combo4.Text
SaveSetting "ConvertIT", "Options", "Op5", Combo5.Text
SaveSetting "ConvertIT", "Options", "Op6", Combo6.Text
SaveSetting "ConvertIT", "Options", "Op7", Combo7.Text
SaveSetting "ConvertIT", "Options", "Op8", Combo8.Text
SaveSetting "ConvertIT", "Options", "Op9", Check1.Value
SaveSetting "ConvertIT", "Options", "Op10", Check2.Value
SaveSetting "ConvertIT", "Options", "Op11", Option1.Value
SaveSetting "ConvertIT", "Options", "Op12", Option2.Value
SaveSetting "ConvertIT", "Options", "Op13", Combo9.Text
End
End Sub

Private Sub AVC1_Converting()
On Error Resume Next
Select Case IntCaptn
Case Is = 0
IntCaptn = 1
Label9.Caption = "Please Wait Converting..."
Case Is = 1
IntCaptn = 2
Label9.Caption = "Please Wait Converting>.."
Case Is = 2
IntCaptn = 3
Label9.Caption = "Please Wait Converting.>."
Case Is = 3
IntCaptn = 0
Label9.Caption = "Please Wait Converting..>"
End Select
End Sub

Private Sub AVC1_Complete()
On Error Resume Next
Label9.Caption = "Done"
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command3.Visible = False
IntCaptn = 0
End Sub

Private Sub AVC1_ErrorEvent(ErrorMessage As String)
Label9.Caption = ErrorMessage
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command3.Visible = False
IntCaptn = 0
End Sub

Private Sub LoadCombos()
On Error Resume Next
Combo1.AddItem "Custom Format"
Combo1.AddItem "AVI (MPEG4-DivX-Compatible)"
Combo1.AddItem "AVI (MPEG4-DivX-Best Quality)"
Combo1.AddItem "AVI (MPEG4-XviD-Compatible)"
Combo1.AddItem "AVI (MPEG4-XviD-Best Quality)"
Combo1.AddItem "Animated GIF (64x48) GIF"
Combo1.AddItem "Animated GIF (160x120) GIF"
Combo1.AddItem "Animated GIF (320x240) GIF"
Combo1.AddItem "Animated GIF (Same as Input) GIF"
Combo1.AddItem "iPOD Video (320x240) MPEG4"
Combo1.AddItem "iPOD Video2 (640x480) MPEG4"
Combo1.AddItem "iPhone Video"
Combo1.AddItem "iPOD Nano (MP3)"
Combo1.AddItem "iPOD Mini (MP3)"
Combo1.AddItem "iPOD Shuffle (MP3)"
Combo1.AddItem "MPEG (Max Compatible)"
Combo1.AddItem "MPEG (Best Quality)"
Combo1.AddItem "WMV(WindowsMedia-Compatible)"
Combo1.AddItem "WMV(WindowsMedia-Best Quality)"
Combo1.AddItem "RM (RealONE/RealPlayer)"
Combo1.AddItem "QuickTime MOV"
Combo1.AddItem "FLV (Flash Live Video)"
Combo1.AddItem "Sony PSP (Normal)"
Combo1.AddItem "Sony PSP (Excellent Quality)"
Combo1.AddItem "Sony PSP (Minimal Size)"
Combo1.AddItem "Sony PSP (PSP AVC)"
Combo1.AddItem "Zune MPEG-4 Video"
Combo1.AddItem "Zune MPEG-4 AVC"
Combo1.AddItem "Zune WMV"
Combo1.AddItem "3GP (for Mobile)"
Combo1.AddItem "3G2 (for Mobile)"
Combo1.AddItem "WAVE Sound (WAV)"
Combo1.AddItem "MP3 Sound (MP3)"

Combo2.AddItem "4xm"
Combo2.AddItem "8bps"
Combo2.AddItem "asv1"
Combo2.AddItem "asv2"
Combo2.AddItem "camtasia"
Combo2.AddItem "cinepak"
Combo2.AddItem "cljr"
Combo2.AddItem "cyuv"
Combo2.AddItem "dvvideo"
Combo2.AddItem "ffv1"
Combo2.AddItem "ffvhuff"
Combo2.AddItem "flic"
Combo2.AddItem "flv"
Combo2.AddItem "h261"
Combo2.AddItem "h263"
Combo2.AddItem "h263i"
Combo2.AddItem "h263p"
Combo2.AddItem "h264"
Combo2.AddItem "huffyuv"
Combo2.AddItem "idcinvideo"
Combo2.AddItem "indeo3"
Combo2.AddItem "interplayvideo"
Combo2.AddItem "ljpeg"
Combo2.AddItem "loco"
Combo2.AddItem "mdec"
Combo2.AddItem "mjpeg"
Combo2.AddItem "mjpegb"
Combo2.AddItem "pgm"
Combo2.AddItem "pgmyuv"
Combo2.AddItem "png"
Combo2.AddItem "ppm"
Combo2.AddItem "qdraw"
Combo2.AddItem "qpeg"
Combo2.AddItem "qtrle"
Combo2.AddItem "rawvideo"
Combo2.AddItem "roqvideo"
Combo2.AddItem "rpza"
Combo2.AddItem "rv10"
Combo2.AddItem "rv20"
Combo2.AddItem "smc"
Combo2.AddItem "snow"
Combo2.AddItem "sp5x"
Combo2.AddItem "svq1"
Combo2.AddItem "svq3"
Combo2.AddItem "theora"
Combo2.AddItem "truemotion1"
Combo2.AddItem "ultimotion"
Combo2.AddItem "vc9"
Combo2.AddItem "vcr1"
Combo2.AddItem "vmdvideo"
Combo2.AddItem "vp3"
Combo2.AddItem "vqavideo"
Combo2.AddItem "wmv1"
Combo2.AddItem "wmv2"
Combo2.AddItem "wmv3"
Combo2.AddItem "wnv1"
Combo2.AddItem "xan_wc3"
Combo2.AddItem "xl"
Combo2.AddItem "xvid"
Combo2.AddItem "zlib"
Combo2.AddItem "mpeg1video"
Combo2.AddItem "mpeg2video"
Combo2.AddItem "mpeg4"
Combo2.AddItem "mpegvideo"
Combo2.AddItem "msmpeg4"
Combo2.AddItem "msmpeg4v1"
Combo2.AddItem "msmpeg4v2"
Combo2.AddItem "msrle"
Combo2.AddItem "msvideo1"
Combo2.AddItem "mszh"
Combo2.AddItem "pam"
Combo2.AddItem "pbm"
Combo4.AddItem "ac3"
Combo4.AddItem "aac"
Combo4.AddItem "alac"
Combo4.AddItem "amr_nb"
Combo4.AddItem "mp2"
Combo4.AddItem "mp3"
Combo4.AddItem "mp3adu"
Combo4.AddItem "mp3on4"
Combo4.AddItem "mace3"
Combo4.AddItem "mace6"
Combo4.AddItem "interplay_dpcm"
Combo4.AddItem "pcm_alaw"
Combo4.AddItem "pcm_mulaw"
Combo4.AddItem "pcm_s16be"
Combo4.AddItem "pcm_s16le"
Combo4.AddItem "pcm_s8"
Combo4.AddItem "pcm_u16be"
Combo4.AddItem "pcm_u16le"
Combo4.AddItem "pcm_u8"
Combo4.AddItem "adpcm_4xm"
Combo4.AddItem "adpcm_adx"
Combo4.AddItem "adpcm_ct"
Combo4.AddItem "adpcm_ea"
Combo4.AddItem "adpcm_ima_dk3"
Combo4.AddItem "adpcm_ima_dk4"
Combo4.AddItem "adpcm_ima_qt"
Combo4.AddItem "adpcm_ima_smjpeg"
Combo4.AddItem "adpcm_ima_wav"
Combo4.AddItem "adpcm_ima_ws"
Combo4.AddItem "adpcm_ms"
Combo4.AddItem "adpcm_swf"
Combo4.AddItem "adpcm_xa"
Combo4.AddItem "real_144"
Combo4.AddItem "real_288"
Combo4.AddItem "roq_dpcm"
Combo4.AddItem "shorten"
Combo4.AddItem "sol_dpcm"
Combo4.AddItem "sonic"
Combo4.AddItem "wmav1"
Combo4.AddItem "wmav2"
Combo4.AddItem "ws_snd1"
Combo4.AddItem "xan_dpcm"
Combo4.AddItem "sonicls"
Combo4.AddItem "vmdaudio"
Combo4.AddItem "g726"
Combo4.AddItem "flac"

Combo3.AddItem "216"
Combo3.AddItem "360"
Combo3.AddItem "512"
Combo3.AddItem "768"
Combo3.AddItem "1000"
Combo3.AddItem "1200"
Combo3.AddItem "1500"

Combo7.AddItem "30"
Combo7.AddItem "29.97"
Combo7.AddItem "25"
Combo7.AddItem "24"
Combo7.AddItem "11"

Combo8.AddItem "128x96"
Combo8.AddItem "176x144"
Combo8.AddItem "320x240"
Combo8.AddItem "352x288"
Combo8.AddItem "368x208"
Combo8.AddItem "640x480"
Combo8.AddItem "704x576"
Combo8.AddItem "800x600"
Combo8.AddItem "1024x768"
Combo8.AddItem "1408x1152"

Combo5.AddItem "48"
Combo5.AddItem "64"
Combo5.AddItem "80"
Combo5.AddItem "128"
Combo5.AddItem "192"
Combo5.AddItem "256"
Combo5.AddItem "384"
Combo5.AddItem "512"

Combo6.AddItem "8000"
Combo6.AddItem "22050"
Combo6.AddItem "44100"
Combo6.AddItem "48000"

Combo9.AddItem ".3g2"
Combo9.AddItem ".3gp"
Combo9.AddItem ".aif"
Combo9.AddItem ".aifc"
Combo9.AddItem ".aiff"
Combo9.AddItem ".au"
Combo9.AddItem ".avi"
Combo9.AddItem ".snd"
Combo9.AddItem ".asf"
Combo9.AddItem ".cda"
Combo9.AddItem ".dvr-ms"
Combo9.AddItem ".dv"
Combo9.AddItem ".flv"
Combo9.AddItem ".ivf"
Combo9.AddItem ".mpg"
Combo9.AddItem ".mpeg"
Combo9.AddItem ".m1v"
Combo9.AddItem ".mp2"
Combo9.AddItem ".mp3"
Combo9.AddItem ".mp4"
Combo9.AddItem ".mpe"
Combo9.AddItem ".mpv2"
Combo9.AddItem ".m3u"
Combo9.AddItem ".mpa"
Combo9.AddItem ".mid"
Combo9.AddItem ".midi"
Combo9.AddItem ".mov"
Combo9.AddItem ".ogg"
Combo9.AddItem ".qt"
Combo9.AddItem ".ra"
Combo9.AddItem ".rm"
Combo9.AddItem ".ram"
Combo9.AddItem ".rmi"
Combo9.AddItem ".swf"
Combo9.AddItem ".wma"
Combo9.AddItem ".wmv"
Combo9.AddItem ".wm"
Combo9.AddItem ".wma"
Combo9.AddItem ".wmv"
Combo9.AddItem ".wm"
Combo9.AddItem ".wav"
Combo9.AddItem ".wmz"
Combo9.AddItem ".wms"


End Sub

Private Sub EnableCombos()
On Error Resume Next
If Combo1.Text = "Custom Format" Then
Combo9.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
Combo5.Enabled = True
Combo6.Enabled = True
Combo7.Enabled = True
Combo8.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Else
Combo9.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False
Combo6.Enabled = False
Combo7.Enabled = False
Combo8.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
EnableCombos
End Sub
