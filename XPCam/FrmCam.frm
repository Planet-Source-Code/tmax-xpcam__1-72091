VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCam 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "     XP Cam"
   ClientHeight    =   12150
   ClientLeft      =   0
   ClientTop       =   90
   ClientWidth     =   10875
   ControlBox      =   0   'False
   Icon            =   "FrmCam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12150
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicCont1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   11000
      ScaleHeight     =   2790
      ScaleWidth      =   10890
      TabIndex        =   36
      Top             =   9330
      Width           =   10920
      Begin Project1.XPcmd CmdCloseTemplate 
         Height          =   345
         Left            =   10365
         TabIndex        =   71
         Top             =   90
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   609
         Caption         =   "^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin VB.PictureBox PicTemp 
         Height          =   1440
         Index           =   4
         Left            =   8610
         ScaleHeight     =   68.148
         ScaleMode       =   0  'User
         ScaleWidth      =   53.621
         TabIndex        =   41
         Top             =   540
         Width           =   1920
      End
      Begin VB.PictureBox PicTemp 
         Height          =   1440
         Index           =   0
         Left            =   360
         ScaleHeight     =   68.148
         ScaleMode       =   0  'User
         ScaleWidth      =   53.621
         TabIndex        =   40
         Top             =   540
         Width           =   1920
      End
      Begin VB.PictureBox PicTemp 
         Height          =   1440
         Index           =   1
         Left            =   2415
         ScaleHeight     =   68.148
         ScaleMode       =   0  'User
         ScaleWidth      =   53.621
         TabIndex        =   39
         Top             =   540
         Width           =   1920
      End
      Begin VB.PictureBox PicTemp 
         Height          =   1440
         Index           =   2
         Left            =   4485
         ScaleHeight     =   68.148
         ScaleMode       =   0  'User
         ScaleWidth      =   53.621
         TabIndex        =   38
         Top             =   540
         Width           =   1920
      End
      Begin VB.PictureBox PicTemp 
         Height          =   1440
         Index           =   3
         Left            =   6540
         ScaleHeight     =   68.148
         ScaleMode       =   0  'User
         ScaleWidth      =   53.621
         TabIndex        =   37
         Top             =   540
         Width           =   1920
      End
      Begin VB.Image Image3 
         Height          =   2910
         Left            =   0
         Picture         =   "FrmCam.frx":164A
         Stretch         =   -1  'True
         Top             =   -75
         Width           =   11430
      End
   End
   Begin VB.PictureBox PicFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   11385
      ScaleHeight     =   472
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   616
      TabIndex        =   35
      Top             =   975
      Width           =   9240
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   12840
      Top             =   720
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7230
      Left            =   150
      TabIndex        =   5
      Top             =   990
      Width           =   9300
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         DrawWidth       =   3
         FillColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7080
         Left            =   60
         ScaleHeight     =   7050
         ScaleWidth      =   9210
         TabIndex        =   6
         Top             =   60
         Width           =   9240
         Begin VB.Label LblCurrent 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   8295
            TabIndex        =   32
            Top             =   6780
            Width           =   660
         End
         Begin VB.Shape ShpMouse 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   120
            Top             =   6840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label LblRegion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            ForeColor       =   &H0000FFFF&
            Height          =   705
            Index           =   0
            Left            =   -10000
            TabIndex        =   11
            Top             =   3000
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label LBox 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   120
            Index           =   3
            Left            =   7140
            MousePointer    =   6  'Size NE SW
            TabIndex        =   10
            Top             =   1440
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label LBox 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   120
            Index           =   2
            Left            =   7380
            MousePointer    =   8  'Size NW SE
            TabIndex        =   9
            Top             =   1440
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label LBox 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   120
            Index           =   1
            Left            =   7380
            MousePointer    =   6  'Size NE SW
            TabIndex        =   8
            Top             =   1080
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label LBox 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   120
            Index           =   0
            Left            =   7140
            MousePointer    =   8  'Size NW SE
            TabIndex        =   7
            Top             =   1080
            Visible         =   0   'False
            Width           =   120
         End
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7230
      Left            =   150
      TabIndex        =   3
      Top             =   990
      Width           =   9300
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         DrawWidth       =   3
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7080
         Left            =   60
         ScaleHeight     =   7050
         ScaleWidth      =   9210
         TabIndex        =   4
         Top             =   60
         Width           =   9240
         Begin VB.Label LblCapture 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capture"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   8295
            TabIndex        =   33
            Top             =   6780
            Width           =   690
         End
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   -75
      TabIndex        =   12
      Top             =   9255
      Width           =   10995
      Begin VB.HScrollBar FSMotion 
         Height          =   240
         LargeChange     =   10
         Left            =   8580
         Max             =   100
         Min             =   1
         TabIndex        =   73
         Top             =   1110
         Value           =   1
         Width           =   1785
      End
      Begin VB.HScrollBar FSColor 
         Height          =   255
         LargeChange     =   10
         Left            =   8580
         Max             =   50
         Min             =   1
         TabIndex        =   72
         Top             =   480
         Value           =   1
         Width           =   1785
      End
      Begin Project1.XPcmd CmdSaveLV 
         Height          =   405
         Left            =   210
         TabIndex        =   61
         Top             =   165
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "Save LV"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin VB.TextBox TxtInterval 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10455
         TabIndex        =   20
         Text            =   "2"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox TxtSColor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10455
         TabIndex        =   15
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxtSMotion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10455
         TabIndex        =   14
         Text            =   "0"
         Top             =   1080
         Width           =   375
      End
      Begin VB.ComboBox cboSndFiles 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   13
         Top             =   2385
         Width           =   8490
      End
      Begin MSComctlLib.ListView lV1 
         Height          =   2115
         Left            =   1425
         TabIndex        =   16
         Top             =   210
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   3731
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Area"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Left"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Top"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "width"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Height"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Colour"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Motion"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Sound"
            Object.Width           =   7056
         EndProperty
      End
      Begin Project1.XPcmd CmdLoadLV 
         Height          =   405
         Left            =   210
         TabIndex        =   62
         Top             =   686
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "Load LV"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdClearLV 
         Height          =   405
         Left            =   210
         TabIndex        =   63
         Top             =   1207
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "Clear LV"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdCamFormat 
         Height          =   405
         Left            =   210
         TabIndex        =   64
         Top             =   1728
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "Format"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdCamOption 
         Height          =   405
         Left            =   210
         TabIndex        =   65
         Top             =   2250
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "Source"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdPlayWav 
         Height          =   360
         Left            =   9960
         TabIndex        =   66
         Top             =   2370
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdBrowse 
         Height          =   360
         Left            =   10440
         TabIndex        =   67
         Top             =   2370
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area Interval"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   9195
         TabIndex        =   21
         Top             =   1725
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motion sensitivity:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   8595
         TabIndex        =   18
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color sensitivity:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   8595
         TabIndex        =   17
         Top             =   270
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   2955
         Left            =   -465
         Picture         =   "FrmCam.frx":8B69
         Stretch         =   -1  'True
         Top             =   -75
         Width           =   11430
      End
      Begin VB.Shape Shape3 
         Height          =   1095
         Left            =   60
         Top             =   1300
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7230
      Left            =   150
      TabIndex        =   0
      Top             =   990
      Width           =   9300
      Begin Project1.XPcmd CmdClearList 
         Height          =   435
         Left            =   7875
         TabIndex        =   68
         Top             =   3615
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Caption         =   "Clear List"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.MediaCtrl MediaCtrl1 
         Height          =   3375
         Left            =   2595
         TabIndex        =   22
         Top             =   3540
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   5953
      End
      Begin VB.ListBox LstCapture 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         ItemData        =   "FrmCam.frx":10088
         Left            =   240
         List            =   "FrmCam.frx":1008A
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   120
         Width           =   9015
      End
      Begin VB.CheckBox ChkAll 
         BackColor       =   &H00000000&
         Caption         =   "Checked All"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   3525
         Width           =   1455
      End
      Begin Project1.XPcmd CmdPlayBack 
         Height          =   435
         Left            =   7890
         TabIndex        =   69
         Top             =   4260
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Caption         =   "Play Back"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdPlayAvi 
         Height          =   435
         Left            =   7890
         TabIndex        =   70
         Top             =   4920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Caption         =   "Play Avi"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin VB.Label LblFrame 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2640
         TabIndex        =   19
         Top             =   6960
         Width           =   540
      End
      Begin VB.Image ImgCapture 
         Height          =   3375
         Left            =   2625
         Stretch         =   -1  'True
         Top             =   3540
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin VB.PictureBox PicBackground 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9330
      Left            =   -15
      Picture         =   "FrmCam.frx":1008C
      ScaleHeight     =   9300
      ScaleWidth      =   10875
      TabIndex        =   23
      Top             =   -15
      Width           =   10905
      Begin Project1.XPcmd CmdVideo 
         Height          =   450
         Left            =   9585
         TabIndex        =   42
         Top             =   1155
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "Video"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin VB.PictureBox pic1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         FillColor       =   &H000000FF&
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   9570
         ScaleHeight     =   195
         ScaleWidth      =   765
         TabIndex        =   34
         Top             =   4800
         Width           =   795
      End
      Begin VB.CheckBox ChkSound 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   9585
         TabIndex        =   29
         Top             =   5550
         Width           =   195
      End
      Begin VB.CheckBox ChkDetect 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9600
         TabIndex        =   28
         Top             =   5160
         Width           =   180
      End
      Begin VB.PictureBox PicContainer 
         Height          =   165
         Left            =   450
         ScaleHeight     =   105
         ScaleWidth      =   465
         TabIndex        =   26
         Top             =   -1000
         Width           =   525
         Begin MSComctlLib.ImageList imlSmallIcons 
            Left            =   0
            Top             =   3195
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmCam.frx":2D2CB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmCam.frx":2D425
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmCam.frx":2D9BF
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmCam.frx":2DF59
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmCam.frx":2E0B3
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Image imgClick 
            Height          =   480
            Left            =   720
            Picture         =   "FrmCam.frx":2E9C5
            Top             =   3195
            Width           =   480
         End
         Begin VB.Image imgPrint 
            Height          =   240
            Left            =   1320
            Picture         =   "FrmCam.frx":2F28F
            Top             =   4035
            Width           =   240
         End
         Begin VB.Image imgBack 
            Height          =   1440
            Left            =   1365
            Picture         =   "FrmCam.frx":2F819
            Top             =   0
            Width           =   3165
         End
         Begin VB.Image imgCan 
            Height          =   6000
            Left            =   3555
            Picture         =   "FrmCam.frx":3E6DB
            Top             =   1845
            Width           =   6000
         End
         Begin VB.Image imgImages 
            Height          =   480
            Left            =   720
            Picture         =   "FrmCam.frx":43010
            Top             =   1275
            Width           =   480
         End
         Begin VB.Image imgOrder 
            Height          =   240
            Left            =   1080
            Picture         =   "FrmCam.frx":464DA
            Top             =   4755
            Width           =   240
         End
         Begin VB.Image imgSlide 
            Height          =   240
            Left            =   2040
            Picture         =   "FrmCam.frx":46A64
            Top             =   4755
            Width           =   240
         End
         Begin VB.Image imgBurn 
            Height          =   240
            Left            =   1920
            Picture         =   "FrmCam.frx":46FEE
            Top             =   4035
            Width           =   240
         End
         Begin VB.Image imgNew 
            Height          =   240
            Left            =   0
            Picture         =   "FrmCam.frx":47578
            Top             =   4395
            Width           =   240
         End
         Begin VB.Image imgUpload 
            Height          =   240
            Left            =   1680
            Picture         =   "FrmCam.frx":47B02
            Top             =   4755
            Width           =   240
         End
         Begin VB.Image imgShare 
            Height          =   240
            Left            =   480
            Picture         =   "FrmCam.frx":4808C
            Top             =   4755
            Width           =   240
         End
      End
      Begin Project1.XPcmd CmdStart 
         Height          =   450
         Left            =   9585
         TabIndex        =   43
         Top             =   2010
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "Start"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdStop 
         Height          =   450
         Left            =   9585
         TabIndex        =   44
         Top             =   2595
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "Stop"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdReset 
         Height          =   450
         Left            =   9585
         TabIndex        =   45
         Top             =   3150
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "Reset"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdRecord 
         Height          =   450
         Left            =   9585
         TabIndex        =   46
         Top             =   3855
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "Record"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdFrame3 
         Height          =   450
         Left            =   9585
         TabIndex        =   47
         Top             =   6090
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "Current"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdFrame2 
         Height          =   450
         Left            =   9585
         TabIndex        =   48
         Top             =   6735
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "Capture"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdFrame1 
         Height          =   450
         Left            =   9585
         TabIndex        =   49
         Top             =   7365
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "History"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdOption 
         Height          =   450
         Left            =   195
         TabIndex        =   50
         Top             =   8595
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "Option"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdStartCapVideo 
         Height          =   450
         Left            =   2655
         TabIndex        =   51
         Top             =   8580
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   794
         Caption         =   "Rec"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdStopCaPVideo 
         Height          =   450
         Left            =   3555
         TabIndex        =   52
         Top             =   8580
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   794
         Caption         =   "Stop"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdSnap 
         Height          =   450
         Left            =   4470
         TabIndex        =   53
         Top             =   8580
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   794
         Caption         =   "Snap"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdCloseVideo 
         Height          =   450
         Left            =   5370
         TabIndex        =   54
         Top             =   8580
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   794
         Caption         =   "Closed"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdTemplate 
         Height          =   450
         Left            =   7470
         TabIndex        =   55
         Top             =   8580
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   794
         Caption         =   "Template"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdShowBox 
         Height          =   450
         Left            =   8805
         TabIndex        =   56
         Top             =   8580
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   794
         Caption         =   "Show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdAddRegion 
         Height          =   450
         Left            =   9667
         TabIndex        =   57
         Top             =   8580
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   794
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdDelRegion 
         Height          =   450
         Left            =   10185
         TabIndex        =   58
         Top             =   8580
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   794
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
      End
      Begin Project1.XPcmd CmdEnd 
         Height          =   345
         Left            =   10290
         TabIndex        =   59
         Top             =   45
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   609
         Caption         =   "X"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   56063
      End
      Begin Project1.XPcmd CmdMinimized 
         Height          =   345
         Left            =   9765
         TabIndex        =   60
         Top             =   45
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   609
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   56063
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   9600
         X2              =   10635
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   9600
         X2              =   10635
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Label LblDetect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detect"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   9855
         TabIndex        =   31
         Top             =   5175
         Width           =   555
      End
      Begin VB.Label LblSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   9870
         TabIndex        =   30
         Top             =   5565
         Width           =   540
      End
      Begin VB.Label LblTimer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xx:xx:xx"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DAFF&
         Height          =   240
         Left            =   9675
         TabIndex        =   27
         Top             =   540
         Width           =   960
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   9630
         Top             =   -1000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   240
         Left            =   10425
         Shape           =   1  'Square
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label LblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XP Cam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DAFF&
         Height          =   240
         Left            =   165
         TabIndex        =   25
         Top             =   90
         Width           =   720
      End
      Begin VB.Label LblMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   585
         Width           =   510
      End
   End
End
Attribute VB_Name = "FrmCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' XPCam
' Created by :        Tmax (tmax_visiber@yahoo.com)
' Descriptions :
' - Custom define detection zone /region(s)
' - Alert with Wave file(s)
' - Capture detected image file
' - Capture video with capCreateCaptureWindow api function
' - Play Avi with mciSendString api function
' - Capture photo with frame template
' - video cam option source and format setting
'
' *
' * extra WebCam Mouse controller "ShpMouse"
' * when setting all the region(s) in the straitline without overlap, you can control the ShpMouse movement.
' * use WebCam to control your game play
' *
'
' - Parts of the program ideas are modify from source codes taken from the following
' - Special thanks :    Ray Mercer (raymer@shrinkwrapvb.com) http://www.shrinkwrapvb.com for video capture
'                       Haik Haiotsyan (haik_111@yahoo.com) for motion detect
'                       Teh Ming Han (teh_minghan@hotmail.com) for button.ctrl
'
'   Comments, suggestions, and bug reports can be sent to tmax_visiber@yahoo.com
'
Option Explicit
Dim DetectOn As Integer
Dim SoundOn As Integer
Dim CurrentRegion As Integer
Dim MaxRegion As Integer
Dim dX, dY As Integer
Dim Idx, Idy As Integer
Dim TodayFolder$
Dim SndFile$
Dim Square As Boolean
Dim StartCap As Boolean
Dim RecordVideo As Boolean
Dim RecordEnabled As Boolean
Private Type DetectRegion
        Left As Long
        Top As Long
        Width As Long
        Height As Long
        ColorValue As Long
        DetectValue As Long
        SoundFile As String
End Type
Dim CaptureStart As Boolean
Dim MyRegion(10) As DetectRegion
Dim FrameTemplate(0 To 4) As String
Dim ShowTemplate As Boolean
Dim FSize As Boolean
Dim VideoFile$
Dim VideoFileSize  As Long
Dim Result As String

Private Sub cboSndFiles_Click()
On Error Resume Next
SndFile$ = cboSndFiles.Text
lV1.ListItems(CurrentRegion).SubItems(7) = SndFile$
End Sub

'Press CmdReset to capture without moving object
Private Sub ChkDetect_Click()
DetectOn = ChkDetect.Value
If DetectOn Then
    LblMode.Caption = "Mode : Motion Detect"
Else
    LblMode.Caption = "Mode : Ready"
    pic1.Cls
End If
End Sub

Private Sub ChkAll_Click()
Dim i%
Dim Chk As Boolean
If ChkAll.Value = 1 Then
    Chk = True
Else
    Chk = False
End If
For i% = 0 To LstCapture.ListCount - 1
    LstCapture.Selected(i%) = Chk
Next i%
End Sub

Private Sub ChkSound_Click()
SoundOn = ChkSound.Value
Square = False
If ChkSound.Value = 1 Then Square = True
End Sub

Private Sub CmdAddRegion_Click()
On Error Resume Next
Dim Itmx As ListItem
MaxRegion = lV1.ListItems.Count + 1
Load LblRegion(MaxRegion)
LblRegion(MaxRegion).Caption = MaxRegion
LblRegion(MaxRegion).Left = 0
LblRegion(MaxRegion).Top = 0
LblRegion(MaxRegion).Width = LblRegion(MaxRegion - 1).Width
LblRegion(MaxRegion).Height = LblRegion(MaxRegion - 1).Height
LblRegion(MaxRegion).Visible = True
CurrentRegion = MaxRegion
AddRegionLV MaxRegion
End Sub

Private Sub CmdBrowse_Click()
Dim WavFile$
Dim i%
WavFile$ = BrowseForFolderDlg(App.Path, "Select Wave File ONLY", Me.hWnd, True)
If WavFile$ <> "" Then
    For i% = 0 To cboSndFiles.ListCount - 1
        If cboSndFiles.List(i%) = WavFile$ Then
            cboSndFiles.ListIndex = i%
            Exit Sub
        End If
    Next
    cboSndFiles.AddItem WavFile$
    cboSndFiles.ListIndex = cboSndFiles.ListCount - 1
End If
End Sub

Private Sub CmdCamFormat_Click()
If Not StartCap Then CmdStart_Click
SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&
End Sub

Private Sub CmdCamOption_Click()
If Not StartCap Then CmdStart_Click
SendMessage mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0
'SendMessageAsLong mCapHwnd, WM_CAP_DLG_VIDEOCOMPRESSION, 0&, 0&
End Sub

Private Sub CmdClearList_Click()
LstCapture.Clear
End Sub

Private Sub CmdClearLV_Click()
Dim i%
lV1.ListItems.Clear
If LblRegion.UBound > 0 Then
    For i% = LblRegion.UBound To 1 Step 1
        Unload LblRegion(i%)
    Next
End If
End Sub

Private Sub CmdCloseVideo_Click()
CmdVideo.Enabled = True
CmdSnap.Enabled = False
CmdStartCapVideo.Enabled = False
CmdStopCaPVideo.Enabled = False
CmdCloseVideo.Enabled = False
CloseVideo
TurnDetect True
TurnVideo False
LblMode.Caption = "Mode : No Signal"
End Sub

Private Sub CmdCloseTemplate_Click()
ShowTemplate = False
PicCont1.Left = 11000
End Sub

Private Sub CmdFrame1_Click()
Frame1.ZOrder 0
End Sub

Private Sub CmdFrame2_Click()
Frame2.ZOrder 0
End Sub

Private Sub CmdFrame3_Click()
Frame3.ZOrder 0
End Sub

Private Sub CmdDelRegion_Click()
On Error Resume Next
lV1.ListItems.Remove CurrentRegion
LoadRegion
End Sub

Private Sub CmdEnd_Click()
Unload Me
End Sub

Private Sub CmdLoadLV_Click()
lV1.ListItems.Clear
LoadLvList lV1, App.Path & "\lv.bin"
LoadRegion
End Sub

Private Sub CmdMinimized_Click()
On Error Resume Next
Me.WindowState = 1
End Sub

Private Sub CmdOption_Click()
ToggleFSize
End Sub

Private Sub CmdPlayAvi_Click()
On Error Resume Next
If CmdPlayAvi.Caption = "StopAvi" Then
    CmdPlayAvi.Caption = "PlayAvi"
    If MediaCtrl1.IsPlaying Then MediaCtrl1.mmStop
    Exit Sub
End If
Static m%
If LstCapture.Text = "" Then Exit Sub
m% = m% + 1
MediaCtrl1.Visible = True
ImgCapture.Visible = False
MediaCtrl1.mmClose
MediaCtrl1.FileName = LstCapture.Text
MediaCtrl1.AliasName = m% & LstCapture.Text
MediaCtrl1.mmPlay
CmdPlayAvi.Caption = "StopAvi"
End Sub

Private Sub CmdPlayBack_Click()
On Error Resume Next
Dim i%
If LstCapture.Text = "" Then Exit Sub
MediaCtrl1.Visible = False
ImgCapture.Visible = True
For i% = 0 To LstCapture.ListCount - 1
    If LstCapture.Selected(i%) Then
        LblFrame.Caption = "Frame # " & (i% + 1) & " of " & LstCapture.ListCount
        DoEvents
        ImgCapture.Picture = LoadPicture(TodayFolder$ & "\" & LstCapture.List(i%))
        DoEvents
    End If
Next i%
End Sub

Private Sub CmdPlayWav_Click()
PlaySnd cboSndFiles.Text
End Sub

Private Sub CmdRecord_Click()
Shape2.FillColor = vbGreen
RecordEnabled = Not RecordEnabled
CmdRecord.ForeColor = vbWhite
If RecordEnabled Then
    Shape2.FillColor = vbRed
    CmdRecord.ForeColor = vbRed
End If
End Sub

Private Sub CmdReset_Click()
Picture2.Picture = Picture1.Picture
End Sub

Private Sub CmdSaveLV_Click()
SaveLvList lV1, App.Path & "\lv.bin"
End Sub

Private Sub CmdShowBox_Click()
    If Not LblRegion(1).Visible Then
        turnOffLBox True
    Else
        turnOffLBox False
    End If
    ShpMouse.Visible = Not ShpMouse.Visible
End Sub

Private Sub CmdSnap_Click()
Snapphoto
End Sub

Private Sub CmdStart_Click()
If Not CmdVideo.Enabled Then CmdCloseVideo_Click
If CmdStart.Caption = "Start" Then
    CmdStart.Caption = "Pause"
    If Not StartCap Then
        STARTCAM
    End If
    CmdStop.Enabled = True
    CmdRecord.Enabled = True
    CaptureStart = True
    LblMode.Caption = "Mode : Ready"
Else
    CmdStart.Caption = "Start"
    CaptureStart = False
    LblMode.Caption = "Mode : Pause"
End If
TurnVideo False
TurnDetect True
End Sub

Private Sub CmdStartCapVideo_Click()
CmdStartCapVideo.Enabled = False
CmdStopCaPVideo.Enabled = True
DoEvents
LblMode.Caption = "Mode : Recoding Video"
StartCapture
End Sub

Private Sub CmdStop_Click()
    STOPCAM
    CmdStop.Enabled = False
    CmdStart.Enabled = True
    CmdRecord.Enabled = False
    Shape1.FillColor = vbBlack
    NoPic Picture2
    NoPic Picture1
    ChkDetect.Value = 0
    LblMode.Caption = "Mode : No Signal"
    CmdStart.Caption = "Start"
End Sub

Private Sub CmdStopCaPVideo_Click()
StopCapture
LblMode.Caption = "Mode : Video Capture"
CmdStartCapVideo.Enabled = True
CmdStopCaPVideo.Enabled = False
End Sub

Private Sub CmdTemplate_Click()
ShowTemplate = True
PicCont1.Left = 0
End Sub

Private Sub CmdVideo_Click()
TurnVideo True
TurnDetect False
If ChkDetect.Value = 1 Then ChkDetect.Value = 0
If CaptureStart Then CaptureStart = False
If FSize Then ToggleFSize
CmdVideo.Enabled = False
CmdStartCapVideo.Enabled = True
CmdSnap.Enabled = True
CmdCloseVideo.Enabled = True
StartPreviewVideo
LblMode.Caption = "Mode : Video Capture"
End Sub

Private Sub Form_Load()
CreateRndObj Me, 10
Square = False
UpdateAllBox
RecordVideo = False
LoadSnd
LoadPng
NoPic Picture2
NoPic Picture1
ChkDir
CmdLoadLV_Click
turnOffLBox False
RecordEnabled = False
Shape2.FillColor = vbGreen
FSize = True
LblMode.Caption = "Mode : Ready"
TurnVideo False
TurnDetect True
ToggleFSize
VideoFileSize = 1000000
CreateRndObj Picture1, 10
CreateRndObj Picture2, 10
CreateRndObj pic1, 3
CmdStart_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then DragForm Me.hWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
STOPCAM
End Sub

Private Sub FSColor_Change()
TxtSColor.Text = FSColor.Value
If CurrentRegion > 0 Then lV1.ListItems(CurrentRegion).SubItems(5) = TxtSColor.Text
End Sub

Private Sub FSColor_Scroll()
FSColor_Change
End Sub

Private Sub FSMotion_Change()
TxtSMotion.Text = FSMotion.Value
If CurrentRegion > 0 Then lV1.ListItems(CurrentRegion).SubItems(6) = TxtSMotion.Text
End Sub

Private Sub FSMotion_Scroll()
FSMotion_Change
End Sub

Private Sub LblRegion_Click(Index As Integer)
  CurrentRegion = Index
End Sub

Private Sub LblRegion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    Screen.MousePointer = 15
    CurrentRegion = Index
    Idx = x
    Idy = y
End If
End Sub

Private Sub LblRegion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    LblRegion(Index).Left = LblRegion(Index).Left - (Idx - x)
    LblRegion(Index).Top = LblRegion(Index).Top - (Idy - y)
    UpdateAllBox
End If
End Sub

Private Sub LblRegion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Idx = 0
Idy = 0
If Button = 1 Then
    UpDateLBox 0
End If
Screen.MousePointer = 0
End Sub

Private Sub LBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    dX = x
    dY = y
End If
End Sub

Private Sub LBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    LBox(Index).Left = LBox(Index).Left - (dX - x)
    LBox(Index).Top = LBox(Index).Top - (dY - y)
    UpDateLBox Index
End If
End Sub

Private Sub LBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
dX = 0
dY = 0
End Sub

Private Sub LstCapture_Click()
On Error Resume Next
LblFrame.Caption = "Frame # " & (LstCapture.ListIndex + 1) & " of " & LstCapture.ListCount
ImgCapture.Picture = LoadPicture(TodayFolder$ & "\" & Mid$(LstCapture.Text, InStr(1, LstCapture.Text, "->  ") + 5))
End Sub

Private Sub lV1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim i%
    CurrentRegion = Item.Index
    TxtSColor.Text = Item.SubItems(5)
    TxtSMotion.Text = Item.SubItems(6)
    SndFile$ = Item.SubItems(7)
    For i% = 0 To cboSndFiles.ListCount - 1
        If cboSndFiles.List(i%) = SndFile$ Then
            cboSndFiles.ListIndex = i%
            Exit For
        End If
    Next
End Sub

Private Sub PicBackground_DblClick()
Static cc As Boolean
cc = Not cc
ShowTopMost Me.hWnd, cc
End Sub

Private Sub PicBackground_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then DragForm Me.hWnd
End Sub

Private Sub PicTemp_DblClick(Index As Integer)
ShowTemplate = False
PicFrame.Cls
PasteTemplate FrameTemplate(Index)
PicFrame.Picture = PicFrame.Image
SaveJPG PicFrame, App.Path + "\SaveTemp\" + Format(Date, "ddmmyyyy") + "_" + Format(Time, "hhmmss") + ".jpg"
ShowTemplate = True
End Sub

Private Sub Timer1_Timer()
'getting picture from camera
On Error Resume Next ' prevent clipboard.getdata error
Static i%
If CaptureStart And StartCap Then
    GrapPicture
    Shape1.FillColor = vbGreen
    If DetectOn > 0 Then DetectAllZone
End If
LblTimer.Caption = Format$(Now, "hh:mm:ss")
' If show template
If i% < 5 And ShowTemplate Then
    PicFrame.Picture = Picture1.Picture
    PasteTemplate FrameTemplate(i)
    StretchBlt PicTemp(i%).hdc, 0, 0, PicTemp(i%).Width, PicTemp(i%).Height, PicFrame.hdc, 0, 0, PicFrame.Width, PicFrame.Height, vbSrcCopy
End If
i% = i% + 1
If i% = 5 Then i% = 0
End Sub

Private Sub TxtSColor_Change()
On Error Resume Next
FSColor.Value = Val(TxtSColor.Text)
End Sub

Private Sub TxtSMotion_Change()
On Error Resume Next
FSMotion.Value = Val(TxtSMotion.Text)
End Sub

Sub UpDateLBox(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0:
        If Square Then LBox(0).Top = LBox(2).Top - ((LBox(2).Left - LBox(0).Left))
        LBox(1).Top = LBox(0).Top
        LBox(3).Left = LBox(0).Left
    Case 1:
        If Square Then LBox(1).Top = LBox(3).Top - ((LBox(1).Left - LBox(3).Left))
        LBox(0).Top = LBox(1).Top
        LBox(2).Left = LBox(1).Left
    Case 2:
        If Square Then LBox(2).Top = LBox(0).Top + ((LBox(2).Left - LBox(0).Left))
        LBox(1).Left = LBox(2).Left
        LBox(3).Top = LBox(2).Top
    Case 3:
        If Square Then LBox(3).Top = LBox(1).Top + ((LBox(1).Left - LBox(3).Left))
        LBox(2).Top = LBox(3).Top
        LBox(0).Left = LBox(3).Left
    End Select
    LblRegion(CurrentRegion).Height = LBox(2).Top - LBox(1).Top
    LblRegion(CurrentRegion).Width = LBox(2).Left - LBox(3).Left
    LblRegion(CurrentRegion).Left = LBox(0).Left + LBox(0).Width / 2 ' 60
    LblRegion(CurrentRegion).Top = LBox(0).Top + LBox(0).Height / 2 '60
    SetRegion CurrentRegion
End Sub

Sub UpdateAllBox()
    LBox(0).Left = LblRegion(CurrentRegion).Left - LBox(0).Width / 2
    LBox(0).Top = LblRegion(CurrentRegion).Top - LBox(0).Width / 2
    LBox(1).Left = LblRegion(CurrentRegion).Left + LblRegion(CurrentRegion).Width - LBox(0).Width / 2
    LBox(1).Top = LblRegion(CurrentRegion).Top - LBox(0).Width / 2
    LBox(2).Left = LBox(1).Left
    LBox(2).Top = LblRegion(CurrentRegion).Top + LblRegion(CurrentRegion).Height - LBox(0).Width / 2
    LBox(3).Left = LblRegion(CurrentRegion).Left - LBox(0).Width / 2
    LBox(3).Top = LBox(2).Top
    SetRegion CurrentRegion
End Sub

Sub turnOffLBox(OnOff As Boolean)
    Dim i%
    For i% = 0 To 3
        LBox(i%).Visible = OnOff
    Next i%
    For i% = 0 To LblRegion.Count - 1
        LblRegion(i%).Visible = OnOff
    Next i%
End Sub

Sub AddRegionLV(Index As Integer)
On Error Resume Next
Dim Itmx As ListItem
Set Itmx = lV1.ListItems.Add(Index, Format$(Now, "hh:mm:ss") & Format$(Index), Format$(Index))
    Itmx.SubItems(1) = LblRegion(Index).Left
    Itmx.SubItems(2) = LblRegion(Index).Top
    Itmx.SubItems(3) = LblRegion(Index).Width
    Itmx.SubItems(4) = LblRegion(Index).Height
    Itmx.SubItems(5) = TxtSColor.Text
    Itmx.SubItems(6) = TxtSMotion.Text
    Itmx.SubItems(7) = SndFile$
End Sub

Sub GetAddRegionLV(Index As Integer)
 Dim i%
On Error Resume Next
For i% = 1 To lV1.ListItems.Count
    LblRegion(i%).Caption = lV1.ListItems(i%).Text
    LblRegion(i%).Left = Val(lV1.ListItems(i%).SubItems(1))
    LblRegion(i%).Top = Val(lV1.ListItems(i%).SubItems(2))
    LblRegion(i%).Width = Val(lV1.ListItems(i%).SubItems(3))
    LblRegion(i%).Height = Val(lV1.ListItems(i%).SubItems(4))
    LblRegion(i%).Visible = True
Next
    MyRegion(Index).Left = LblRegion(Index).Left
    MyRegion(Index).Top = LblRegion(Index).Top
    MyRegion(Index).Width = LblRegion(Index).Width
    MyRegion(Index).Height = LblRegion(Index).Height
    MyRegion(Index).ColorValue = TxtSColor.Text
    MyRegion(Index).DetectValue = TxtSMotion.Text
    MyRegion(Index).SoundFile = SndFile$
End Sub

Sub LoadRegion()
Dim i%
On Error Resume Next
UnloadRegion
For i% = 1 To lV1.ListItems.Count
    Load LblRegion(i%)
    lV1.ListItems(i).Checked = True
    LblRegion(i%).Caption = lV1.ListItems(i%).Text
    LblRegion(i%).Left = Val(lV1.ListItems(i%).SubItems(1))
    LblRegion(i%).Top = Val(lV1.ListItems(i%).SubItems(2))
    LblRegion(i%).Width = Val(lV1.ListItems(i%).SubItems(3))
    LblRegion(i%).Height = Val(lV1.ListItems(i%).SubItems(4))
    LblRegion(i%).Visible = True
    MyRegion(i%).Left = LblRegion(i%).Left
    MyRegion(i%).Top = LblRegion(i%).Top
    MyRegion(i%).Width = LblRegion(i%).Width
    MyRegion(i%).Height = LblRegion(i%).Height
    MyRegion(i%).ColorValue = Val(lV1.ListItems(i%).SubItems(5))
    MyRegion(i%).DetectValue = Val(lV1.ListItems(i%).SubItems(6))
    MyRegion(i%).SoundFile = lV1.ListItems(i%).SubItems(7)
Next
End Sub

Sub SetRegion(Index As Integer)
Dim Itmx As ListItem
If Index = 0 Then Exit Sub
lV1.ListItems(Index).Selected = True
Set Itmx = lV1.ListItems(Index)
    Itmx.SubItems(1) = LblRegion(Index).Left
    Itmx.SubItems(2) = LblRegion(Index).Top
    Itmx.SubItems(3) = LblRegion(Index).Width
    Itmx.SubItems(4) = LblRegion(Index).Height
    TxtSColor.Text = Itmx.SubItems(5)
    TxtSMotion.Text = Itmx.SubItems(6)
    SndFile$ = Itmx.SubItems(7)
    MyRegion(Index).Left = LblRegion(Index).Left
    MyRegion(Index).Top = LblRegion(Index).Top
    MyRegion(Index).Width = LblRegion(Index).Width
    MyRegion(Index).Height = LblRegion(Index).Height
End Sub

Sub UnloadRegion()
 Dim i%
If LblRegion.Count > 0 Then
    For i% = LblRegion.Count - 1 To 1 Step -1
        Unload LblRegion(i%)
    Next
End If
End Sub



Private Function Different(ByVal a As Long, ByVal b As Long) As Boolean
'Checks different of two colors
Dim ar, ag, ab, br, bg, bb As Long
Dim sense As Integer
ar = a Mod 256
ag = (a \ 256) Mod 256
ab = ((a \ 256) \ 256) Mod 256
br = b Mod 256
bg = (b \ 256) Mod 256
bb = ((b \ 256) \ 256) Mod 256
sense = 255 - FSColor.Value * 5
Different = Abs(ar - br) + Abs(ag - bg) + Abs(ab - bb) > sense  'check color different sensitivity
End Function

Sub DetectAllZone()
'detect all selected zone
Dim i%
For i% = 1 To lV1.ListItems.Count
    If lV1.ListItems(i%).Checked Then DetectMotion i%
Next i%
End Sub

Sub DetectMotion(Index As Integer)
On Error Resume Next
Dim AreaDetect As Integer
Dim A_Interval As Integer
Dim DetectCount As Long
Dim i%, j%
A_Interval = Val(TxtInterval.Text) ' 2
DetectCount = 0
Picture1.ScaleMode = 3
For i = LblRegion(Index).Left / A_Interval To (LblRegion(Index).Left + LblRegion(Index).Width) / A_Interval Step A_Interval
    For j = LblRegion(Index).Top / A_Interval To (LblRegion(Index).Top + LblRegion(Index).Height) / A_Interval Step A_Interval
        If Different(GetPixel(Picture2.hdc, i * A_Interval, j * A_Interval), GetPixel(Picture1.hdc, i * A_Interval, j * A_Interval)) Then
            SetPixel Picture1.hdc, i * A_Interval, j * A_Interval, RGB(255, 0, 0)
            DetectCount = DetectCount + 1
        End If
    Next
Next
AreaDetect = (LblRegion(Index).Width / A_Interval) * (LblRegion(Index).Height / A_Interval) / (A_Interval * A_Interval)
If DetectCount > AreaDetect Then DetectCount = AreaDetect
LblRegion(Index).Caption = Int(DetectCount * 100 / AreaDetect)
If Val(LblRegion(Index).Caption) > MyRegion(Index).DetectValue Then      ' check lv FSMotion.value & snd files
    If Not RecordEnabled Then
        Shape1.FillColor = vbYellow
        ShpMouse.Left = LblRegion(Index).Left
    Else
        Shape1.FillColor = vbRed
        SaveCapture
    End If
    If SoundOn > 0 Then
        PlaySnd MyRegion(Index).SoundFile
    End If
End If
UpDateProgress pic1, LblRegion(Index).Caption
Picture1.ScaleMode = 1
End Sub

Sub ToggleFSize()
Me.ScaleMode = 3
If FSize Then
    CmdCloseTemplate_Click
    Me.Height = 9255
Else
    Me.Height = 12150
End If
' Show Template
'me.Width=21450
' Normal
'me.width=10900
FSize = Not FSize
CreateRndObj Me, 10
Me.ScaleMode = 1
End Sub

Sub TurnDetect(OnOff As Boolean)
CmdStop.Visible = OnOff
CmdReset.Visible = OnOff
Shape1.Visible = OnOff
Shape2.Visible = OnOff
pic1.Visible = OnOff
ChkDetect.Visible = OnOff
ChkSound.Visible = OnOff
LblDetect.Visible = OnOff
LblSound.Visible = OnOff
CmdRecord.Visible = OnOff
CmdShowBox.Visible = OnOff
CmdAddRegion.Visible = OnOff
CmdDelRegion.Visible = OnOff
CmdOption.Visible = OnOff
End Sub

Sub TurnVideo(OnOff As Boolean)
CmdStartCapVideo.Visible = OnOff
CmdSnap.Visible = OnOff
CmdStopCaPVideo.Visible = OnOff
CmdCloseVideo.Visible = OnOff
End Sub

Sub STARTCAM()
Dim ret&
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, Me.hWnd, 0)
If mCapHwnd <> 0 Then StartCap = True
DoEvents
ret& = SendMessage(mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0) 'connecting to camera
If ret& = 0 Then
    StartCap = False
    CmdStart.Caption = "Start"
    MsgBox "Could not detect webcam device.", vbOKOnly, "XPress Secure"
    CmdStop_Click
End If
DoEvents
End Sub

Sub STOPCAM()
DoEvents
If StartCap Then
    SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
    DestroyWindow mCapHwnd
    StartCap = False
End If
CaptureStart = False
RecordEnabled = False
End Sub

Sub CloseVideo()
SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0&, 0&
DestroyWindow mCapHwnd
Me.MousePointer = 0
NoPic Picture2
NoPic Picture1
CmdStart.Enabled = True
CmdStop.Enabled = True
CmdReset.Enabled = True
End Sub

Sub GrapPicture()
SendMessage mCapHwnd, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0
SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
Picture1.Picture = Clipboard.GetData
End Sub

Sub SaveCapture()
    If Picture1.Picture Then
        SaveJPG Picture1, TodayFolder$ + "\" + Format(Date, "ddmmyyyy") + "_" + Format(Time, "hhmmss") + ".jpg"
        LstCapture.AddItem Format(Date, "ddmmyyyy") + "_" + Format(Time, "hhmmss") + ".jpg"
    End If
End Sub

Sub SnapPhoto1() 'Large filesize
SendMessageAsString mCapHwnd, WM_CAP_FILE_SAVEDIB, 0&, App.Path & "\TempDIB.bmp"
Picture2.Picture = LoadPicture(App.Path & "\TempDIB.bmp")
End Sub

Sub Snapphoto()
Picture1.AutoRedraw = False
BitBlt Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, 0, 0, vbSrcCopy
Picture1.AutoRedraw = True
Picture2.Picture = Picture2.Image
If Picture2.Picture Then
    SaveJPG Picture2, App.Path & "\test1.jpg"
End If
End Sub

Sub StartCapture()
VideoFile$ = App.Path & "\Video\" & Format(Date, "ddmmyyyy") + "_" + Format(Time, "hhmmss") + ".avi"
SendMessageAsString mCapHwnd, WM_CAP_FILE_SET_CAPTURE_FILE, 0&, VideoFile$
SendMessage mCapHwnd, WM_CAP_FILE_ALLOCATE, 0&, VideoFileSize
SendMessage mCapHwnd, WM_CAP_SEQUENCE, 0&, 0&
End Sub

Sub StartPreviewVideo()
Dim ret&
  If StartCap Then CmdStop_Click
    mCapHwnd = capCreateCaptureWindow("WebcamCapture", WS_VISIBLE Or WS_CHILD, 0, 0, Picture1.ScaleWidth / Screen.TwipsPerPixelX, Picture1.ScaleHeight / Screen.TwipsPerPixelY, Picture1.hWnd, 0)
    If SendMessage(mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0) Then
        '---set the preview scale---
        ret& = SendMessage(mCapHwnd, WM_CAP_SET_SCALE, True, 0)
        '---set the preview rate (ms)---
        ret& = SendMessage(mCapHwnd, WM_CAP_SET_PREVIEWRATE, 30, 0)
        '---start previewing the image---
        ret& = SendMessage(mCapHwnd, WM_CAP_SET_PREVIEW, True, 0)
    End If
End Sub

Sub StopCapture()
On Error Resume Next
DoEvents
SendMessageAsString mCapHwnd, WM_CAP_STOP, 0&, 0&
SendMessageAsString mCapHwnd, WM_CAP_FILE_SAVEAS, 0&, VideoFile$
DoEvents
LstCapture.AddItem VideoFile$
Me.MousePointer = 0
End Sub

Sub UpDateProgress(pb As Control, ByVal percent)
pb.ForeColor = vbBlack
pb.Cls
pb.ScaleWidth = 100
pb.DrawMode = 10
pb.Line (0, 0)-(percent, pb.ScaleHeight), , BF
If Shape1.FillColor = vbYellow Then
    pb.ForeColor = vbYellow
    pb.Line (percent + 1, 0)-(100, pb.ScaleHeight), , BF
End If
pb.Refresh
End Sub

Sub ChkDir()
TodayFolder$ = App.Path + "\SaveTemp"
If Dir$(TodayFolder, vbDirectory) = vbNullString Then MkDir TodayFolder$
TodayFolder$ = App.Path + "\Video"
If Dir$(TodayFolder, vbDirectory) = vbNullString Then MkDir TodayFolder$
TodayFolder$ = App.Path + "\Detected"
If Dir$(TodayFolder, vbDirectory) = vbNullString Then MkDir TodayFolder$
TodayFolder$ = App.Path + "\Detected\" + Format(Date, "ddmmyyyy")
If Dir$(TodayFolder, vbDirectory) = vbNullString Then MkDir TodayFolder$
End Sub

Sub CreateRndObj(Obj As Object, Rad As Single)
Dim hRgn2 As Long
Dim sc As Integer
    sc = Obj.ScaleMode
    Obj.ScaleMode = 3
    hRgn2 = CreateRoundRectRgn(2, 2, Obj.ScaleWidth - 2, Obj.ScaleHeight - 2, Rad, Rad)
    SetWindowRgn Obj.hWnd, hRgn2, True
    DeleteObject hRgn2
    Obj.ScaleMode = sc
End Sub

Sub LoadSnd()
SndFile$ = App.Path + "\audio\notify.wav"
cboSndFiles.AddItem App.Path + "\audio\notify.wav"
cboSndFiles.AddItem App.Path + "\audio\sound7.wav"
cboSndFiles.AddItem App.Path + "\audio\sound999.wav"
cboSndFiles.AddItem App.Path + "\audio\sound68.wav"
cboSndFiles.AddItem App.Path + "\audio\ready.wav"
cboSndFiles.ListIndex = 0
End Sub

Sub LoadPng()
Dim i%
ShowTemplate = False
For i% = 0 To 4
    FrameTemplate(i) = App.Path & "\png\pic (" & i% + 1 & ").png"
    SetStretchBltMode PicTemp(i).hdc, COLORONCOLOR  'important
Next
End Sub

Sub NoPic(Pic As PictureBox)
Dim X1, Y1, r1 As Integer
r1 = 1500
X1 = 3500
Y1 = 1200
turnOffLBox False
Pic.ScaleMode = 1
Pic.Refresh
With Pic
    .Picture = LoadPicture
    .Cls
    .DrawWidth = 20
    Pic.Line (X1 + r1 - r1 / Sqr(2), Y1 + r1 + r1 / Sqr(2))-(X1 + r1 + r1 / Sqr(2), Y1 + r1 - r1 / Sqr(2)), vbRed
    .CurrentX = .Width \ 2 - r1 * 1.5
    .CurrentY = .Width \ 2
    .Align = vbCenter
    .ForeColor = vbRed
    Pic.Print "No Signal Detected!"
    Pic.Circle ((X1 + r1), (Y1 + r1)), r1
    .DrawWidth = 3
End With
End Sub

Sub PasteTemplate(FileName$)
Dim ImgPng As cImgPng
Set ImgPng = New cImgPng
ImgPng.Load FileName$
ImgPng.StretchDC PicFrame.hdc, 0, 0, PicFrame.ScaleWidth, PicFrame.ScaleHeight
Set ImgPng = Nothing
End Sub
