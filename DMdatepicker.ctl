VERSION 5.00
Begin VB.UserControl DMdatepicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ScaleHeight     =   3750
   ScaleWidth      =   3750
   ToolboxBitmap   =   "DMdatepicker.ctx":0000
   Begin VB.PictureBox picDropDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   1455
      TabIndex        =   73
      Top             =   0
      Width           =   1455
      Begin VB.PictureBox btnCombo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DC9670&
         BorderStyle     =   0  'None
         Height          =   235
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   270
         TabIndex        =   74
         Top             =   15
         Width           =   270
      End
      Begin VB.Label txtXtext 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00DC9670&
         Height          =   195
         Left            =   90
         TabIndex        =   75
         Top             =   30
         Width           =   480
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00B99D7F&
         FillColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   0
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2760
      Top             =   360
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2880
      Top             =   840
   End
   Begin VB.PictureBox picCalendar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   2730
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   2730
      Begin VB.PictureBox picYear 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1785
         ScaleHeight     =   285
         ScaleWidth      =   855
         TabIndex        =   25
         Top             =   105
         Width           =   855
         Begin VB.PictureBox btnSpinDown 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00DC9670&
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   570
            ScaleHeight     =   135
            ScaleWidth      =   285
            TabIndex        =   27
            Top             =   155
            Width           =   285
         End
         Begin VB.PictureBox btnSpinUp 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00DC9670&
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   570
            ScaleHeight     =   135
            ScaleWidth      =   285
            TabIndex        =   26
            Top             =   0
            Width           =   285
         End
         Begin VB.Label txtSpin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2010"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Left            =   60
            TabIndex        =   29
            Top             =   45
            Width           =   360
         End
         Begin VB.Shape shpYear 
            BorderColor     =   &H00DC9670&
            Height          =   285
            Left            =   0
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picMonthList 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   90
         ScaleHeight     =   2415
         ScaleWidth      =   1335
         TabIndex        =   12
         Top             =   375
         Visible         =   0   'False
         Width           =   1335
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "December"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   11
            Left            =   45
            TabIndex        =   24
            Top             =   2145
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "November"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   10
            Left            =   45
            TabIndex        =   23
            Top             =   1945
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "October"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   9
            Left            =   45
            TabIndex        =   22
            Top             =   1745
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "September"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   8
            Left            =   45
            TabIndex        =   21
            Top             =   1545
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "August"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   7
            Left            =   45
            TabIndex        =   20
            Top             =   1345
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "July"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   6
            Left            =   45
            TabIndex        =   19
            Top             =   1145
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "June"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   5
            Left            =   45
            TabIndex        =   18
            Top             =   945
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "May"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   4
            Left            =   45
            TabIndex        =   17
            Top             =   745
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "April"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   3
            Left            =   45
            TabIndex        =   16
            Top             =   560
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "March"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   15
            Top             =   380
            Width           =   1275
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "February"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   14
            Top             =   200
            Width           =   1275
         End
         Begin VB.Shape shpMonthNames 
            BorderColor     =   &H00DC9670&
            Height          =   2415
            Left            =   0
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label lblMonth 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "January"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   13
            Top             =   0
            Width           =   1275
         End
      End
      Begin VB.PictureBox picMonth 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   90
         ScaleHeight     =   285
         ScaleWidth      =   1335
         TabIndex        =   10
         Top             =   105
         Width           =   1335
         Begin VB.PictureBox btnMonthDown 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00DC9670&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1050
            ScaleHeight     =   285
            ScaleWidth      =   285
            TabIndex        =   11
            Top             =   0
            Width           =   285
         End
         Begin VB.Label txtMonth 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Januari"
            ForeColor       =   &H00DC9670&
            Height          =   195
            Left            =   60
            TabIndex        =   28
            Top             =   45
            Width           =   510
         End
         Begin VB.Shape shpMonth 
            BorderColor     =   &H00DC9670&
            Height          =   285
            Left            =   0
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox picDayNames 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   90
         ScaleHeight     =   300
         ScaleWidth      =   2550
         TabIndex        =   2
         Top             =   480
         Width           =   2545
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H00DC9670&
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   2280
            TabIndex        =   3
            Top             =   60
            Width           =   135
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H00DC9670&
            BackStyle       =   0  'Transparent
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   60
            Width           =   165
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H00DC9670&
            BackStyle       =   0  'Transparent
            Caption         =   "T"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   8
            Top             =   60
            Width           =   135
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H00DC9670&
            BackStyle       =   0  'Transparent
            Caption         =   "W"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   7
            Top             =   60
            Width           =   195
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H00DC9670&
            BackStyle       =   0  'Transparent
            Caption         =   "T"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   1200
            TabIndex        =   6
            Top             =   60
            Width           =   135
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H00DC9670&
            BackStyle       =   0  'Transparent
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   1560
            TabIndex        =   5
            Top             =   60
            Width           =   120
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H00DC9670&
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   1920
            TabIndex        =   4
            Top             =   60
            Width           =   135
         End
         Begin VB.Shape shpDayNamesBorder 
            BackColor       =   &H00C0C0FF&
            BorderColor     =   &H00DC9670&
            FillColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   0
            Top             =   0
            Width           =   2545
         End
      End
      Begin VB.PictureBox picDays 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1545
         Left            =   90
         ScaleHeight     =   1545
         ScaleWidth      =   2550
         TabIndex        =   30
         Top             =   840
         Width           =   2545
         Begin VB.Shape shpDayNow 
            BorderWidth     =   2
            Height          =   240
            Left            =   2280
            Shape           =   2  'Oval
            Top             =   1320
            Width           =   315
         End
         Begin VB.Shape shpDays 
            BorderColor     =   &H00DC9670&
            FillColor       =   &H00DC9670&
            Height          =   1545
            Left            =   0
            Top             =   0
            Width           =   2545
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   40
            Left            =   1860
            TabIndex        =   72
            Top             =   1245
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   39
            Left            =   1500
            TabIndex        =   71
            Top             =   1245
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   38
            Left            =   1140
            TabIndex        =   70
            Top             =   1245
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   34
            Left            =   2220
            TabIndex        =   69
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   33
            Left            =   1860
            TabIndex        =   68
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   32
            Left            =   1500
            TabIndex        =   67
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   31
            Left            =   1140
            TabIndex        =   66
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   27
            Left            =   2220
            TabIndex        =   65
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   26
            Left            =   1860
            TabIndex        =   64
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   25
            Left            =   1500
            TabIndex        =   63
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   24
            Left            =   1140
            TabIndex        =   62
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   2220
            TabIndex        =   61
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   1860
            TabIndex        =   60
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   18
            Left            =   1500
            TabIndex        =   59
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   17
            Left            =   1140
            TabIndex        =   58
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   1860
            TabIndex        =   57
            Top             =   285
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   1500
            TabIndex        =   56
            Top             =   285
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   1140
            TabIndex        =   55
            Top             =   285
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   2220
            TabIndex        =   54
            Top             =   45
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   1860
            TabIndex        =   53
            Top             =   45
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   1500
            TabIndex        =   52
            Top             =   45
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   1140
            TabIndex        =   51
            Top             =   45
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   41
            Left            =   2220
            TabIndex        =   50
            Top             =   1245
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   2220
            TabIndex        =   49
            Top             =   285
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   37
            Left            =   780
            TabIndex        =   48
            Top             =   1245
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   36
            Left            =   420
            TabIndex        =   47
            Top             =   1245
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   35
            Left            =   60
            TabIndex        =   46
            Top             =   1245
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   30
            Left            =   780
            TabIndex        =   45
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   29
            Left            =   420
            TabIndex        =   44
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   28
            Left            =   60
            TabIndex        =   43
            Top             =   1005
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   23
            Left            =   780
            TabIndex        =   42
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   22
            Left            =   420
            TabIndex        =   41
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   21
            Left            =   60
            TabIndex        =   40
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   780
            TabIndex        =   39
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   15
            Left            =   420
            TabIndex        =   38
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   60
            TabIndex        =   37
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   780
            TabIndex        =   36
            Top             =   285
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   420
            TabIndex        =   35
            Top             =   285
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   60
            TabIndex        =   34
            Top             =   285
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   780
            TabIndex        =   33
            Top             =   45
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   32
            Top             =   45
            Width           =   255
         End
         Begin VB.Label lblNumbers 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   31
            Top             =   45
            Width           =   255
         End
      End
      Begin VB.Label lblToday 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Today:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DC9670&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2500
         Width           =   2535
      End
      Begin VB.Shape shpCalendar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00DC9670&
         Height          =   2895
         Left            =   0
         Top             =   0
         Width           =   2730
      End
   End
End
Attribute VB_Name = "DMdatepicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Programmer:        Donckers Frank
'                    DarkManSoft@Gmail.com
'
' Description:       User Control DatePicker
'
' Updates:
'                    09/03/2010
'                    Refreshing date when clicked in monthlist
'                    Selection for formats dd/mm/yyyy and mm/dd/yyyy added
'                    Selection for firstweekday added

'=====================================================
' Enum Languages
'=====================================================
Public Enum Language
    [Englisch] = 0
    [Nederlands] = 1
    [Francais] = 2
    [Deutch] = 3
    [Italiano] = 4
    [Espagnol] = 5
End Enum
'=====================================================
' Enum Weekdays
'=====================================================
Public Enum Weekdays
    [Sunday] = 1
    [monday] = 2
    [Tuesday] = 3
    [Wednesday] = 4
    [Thursday] = 5
    [Friday] = 6
    [Saterday] = 7
End Enum

'=====================================================
' Enum DateFormats
'=====================================================
Public Enum DateFormats
    [dd/mm/yyyy] = 0
    [mm/dd/yyyy] = 1
End Enum

'=====================================================
' Enum Styles
'=====================================================
Public Enum Styles
    [Flat] = 0
    [3D] = 1
End Enum

'=====================================================
' Events
'=====================================================
Event Change()
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'=====================================================
' Default Property Values
'=====================================================
' DropDown
Const m_def_DropDnForeColor = &HEA692E
Const m_def_DropDnBackColor = &HFFC0C0
Const m_def_DropDnBorderColor = &HEA692E
Const m_def_DropDnButtonBackColor = &HEA692E
Const m_def_DropDnButtonForeColor = &H80000005
Const m_def_DropDnButtonBorderColor = &HEA692E
' Calendar
Const m_def_CalendarForeColor = &HDC9670
Const m_def_CalendarBorderColor = &HDC9670
Const m_def_CalendarBackColor = &HEA692E
' Month Drop down
Const m_def_DropDnMonthForeColor = &HEA692E
Const m_def_DropDnMonthBackColor = &HFFC0C0
Const m_def_DropDnMonthBorderColor = &HEA692E
Const m_def_DropDnMonthButtonBackColor = &HEA692E
Const m_def_DropDnMonthButtonForeColor = &H80000005
Const m_def_DropDnMonthButtonBorderColor = &HEA692E
' Month List
Const m_def_MonthListBackColor = &HFBD2CA
Const m_def_MonthListForeColor = &HDC9670
Const m_def_MonthListSelectedBackColor = &HDC9670
Const m_def_MonthListSelectedForeColor = &H80000005
Const m_def_MonthListBorderColor = &HDC9670
' Year
Const m_def_SpinYearForeColor = &HEA692E
Const m_def_SpinYearBackColor = &HFFC0C0
Const m_def_SpinYearBorderColor = &HEA692E
Const m_def_SpinYearButtonBackColor = &HEA692E
Const m_def_SpinYearButtonForeColor = &H80000005
Const m_def_SpinYearButtonBorderColor = &HEA692E
' Days
Const m_def_DaysBackColor = &H833606
Const m_def_DaysBorderColor = &HEA692E
Const m_def_DayNamesForeColor = &H80000005
Const m_def_DayNamesBackColor = &HEA692E
Const m_def_DayNamesBorderColor = &HEA692E
Const m_def_DaysForeColor = &HFFC0C0
Const m_def_DaySelectedBackColor = &HFFC0C0
Const m_def_DaySelectedForeColor = &H833606
Const m_def_CurrentDayBOrderColor = &HFFFFFF
' DisabledColors
Const m_def_DisabledBackColor = &HCFF0F2
Const m_def_DisabledBorderColor = &H81BECB
Const m_def_DisabledForeColor = &H81BECB
' Language
Const m_def_Language = 0
'Style
Const m_def_Style = 1
' DateFormat
Const m_def_DateFormat = 0
' FirstWeekDay
Const m_def_FirstWeekDay = 1

'=====================================================
' Property Variables
'=====================================================
' DropDown
Dim m_DropDnForeColor As OLE_COLOR
Dim m_DropDnBackColor As OLE_COLOR
Dim m_DropDnBorderColor As OLE_COLOR
Dim m_DropDnButtonBackColor As OLE_COLOR
Dim m_DropDnButtonForeColor As OLE_COLOR
Dim m_DropDnButtonBorderColor As OLE_COLOR
' Calendar
Dim m_CalendarForeColor As OLE_COLOR
Dim m_CalendarBackColor As OLE_COLOR
Dim m_CalendarBorderColor As OLE_COLOR
' Month
Dim m_DropDnMonthForeColor As OLE_COLOR
Dim m_DropDnMonthBackColor As OLE_COLOR
Dim m_DropDnMonthBorderColor As OLE_COLOR
Dim m_DropDnMonthButtonBackColor As OLE_COLOR
Dim m_DropDnMonthButtonForeColor As OLE_COLOR
Dim m_DropDnMonthButtonBorderColor As OLE_COLOR
' Monthlist
Dim m_MonthListBackColor As OLE_COLOR
Dim m_MonthListForeColor As OLE_COLOR
Dim m_MonthListSelectedBackColor As OLE_COLOR
Dim m_MonthListSelectedForeColor As OLE_COLOR
Dim m_MonthListBorderColor As OLE_COLOR
' Year
Dim m_SpinYearForeColor As OLE_COLOR
Dim m_SpinYearBackColor As OLE_COLOR
Dim m_SpinYearBorderColor As OLE_COLOR
Dim m_SpinYearButtonBackColor As OLE_COLOR
Dim m_SpinYearButtonForeColor As OLE_COLOR
Dim m_SpinYearButtonBorderColor As OLE_COLOR
' Days
Dim m_DaysForeColor As OLE_COLOR
Dim m_DaysBackColor As OLE_COLOR
Dim m_DaysBorderColor As OLE_COLOR
Dim m_DayNamesForeColor As OLE_COLOR
Dim m_DayNamesBackColor As OLE_COLOR
Dim m_DayNamesBorderColor As OLE_COLOR
Dim m_DaySelectedBackColor As OLE_COLOR
Dim m_DaySelectedForeColor As OLE_COLOR
Dim m_CurrentDayBOrderColor As OLE_COLOR
' DisabledColors
Dim m_DisabledBackColor As OLE_COLOR
Dim m_DisabledBorderColor As OLE_COLOR
Dim m_DisabledForeColor As OLE_COLOR
Dim m_Enabled As Boolean
'Language
Dim m_Language As Language
'StartDate
Dim m_StartDate As String
'Style
Dim m_Style As Styles
' DateFormat
Dim m_DateFormat As String
' FirstWeekDay
Dim m_FirstWeekDay As Weekdays

'=====================================================
' Program Variables
'=====================================================
Dim OldScaleMode As Byte
Dim cControl As Control
Dim StartCol As Double, EndCol As Double
Dim RedI As Single, BlueI As Single, GreenI As Single
Dim RedStart As Integer, GreenStart As Integer, BlueStart As Integer
Dim RedEnd As Double, GreenEnd As Double, BlueEnd As Double
Dim i, ii, iii As Integer
Dim NewColor As Single
Dim MidX, MidY As Integer
Dim MonthNow As Byte
Dim YearNow As Integer
Dim ArrMonth(6, 12) As String
Dim ArrDay(6, 7) As String
Dim ArrDayName(6, 7) As String
Dim GetMonth As Byte
Dim CurrentDay As Byte
Dim DayNumber As Byte
Dim DayPos, MonthPos As Byte
Dim FirstDayPos(8, 8) As Byte


'=====================================================
' Raiseevent change on change date
'=====================================================
Private Sub txtXText_Change()
    RaiseEvent Change
End Sub

'=====================================================
' Select Daynumbers
'=====================================================
Private Sub lblNumbers_DblClick(Index As Integer)
    lblNumbers_Click Index
    btnCombo_Click
End Sub
Private Sub lblNumbers_Click(Index As Integer)
    DayNumber = Index
    If lblNumbers(Index).Caption = "" Then Exit Sub
    For i = 0 To 11
        If ArrMonth(m_Language, i) = txtMonth.Caption Then
            GetMonth = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 41
        lblNumbers(i).BackStyle = 0
        lblNumbers(i).BackColor = m_DaysBackColor
        lblNumbers(i).ForeColor = m_DaysForeColor
    Next i
    lblNumbers(DayNumber).BackStyle = 1
    lblNumbers(DayNumber).BackColor = m_DaySelectedBackColor
    lblNumbers(DayNumber).ForeColor = m_DaySelectedForeColor
    Dim NewDate As String
    Dim sMonth, sDay As String
    sMonth = GetMonth
    sDay = lblNumbers(Index).Caption
    If Val(sDay) < 10 Then sDay = "0" & sDay
    If GetMonth < 10 Then sMonth = "0" & GetMonth
    If m_DateFormat = 0 Then
        NewDate = sDay & "/" & sMonth & "/" & txtSpin.Caption
    Else
        NewDate = sMonth & "/" & sDay & "/" & txtSpin.Caption
    End If
    txtXtext.Caption = NewDate
End Sub
'=====================================================
' Set daynumbers
'=====================================================
Private Sub CalculateCalendar()
    Dim GetDay As Byte
    Dim J As Integer
    Dim MonthStart
    Dim NumDays As Long
    Dim FirstDayWD
    Dim NewDate As Date
    GetMonth = 1
    For i = 0 To 11
        If ArrMonth(m_Language, i) = txtMonth.Caption Then
            GetMonth = i + 1
            Exit For
        End If
    Next i
    For i = 1 To 41
        If Day(Date) = i Then
            GetDay = i
        End If
    Next
    NewDate = GetDay & "/" & GetMonth & "/" & txtSpin.Caption
    MonthStart = DateSerial(Year(NewDate), Month(NewDate), 1)
    ' Number of days in the month
    NumDays = DateDiff("d", MonthStart, DateAdd("m", 1, MonthStart))
    ' First weekday basef on m_FirstWeekDay
    If m_FirstWeekDay = 1 Then
        FirstDayWD = Weekday(MonthStart, m_FirstWeekDay) + 1
    Else
        FirstDayWD = Weekday(MonthStart, m_FirstWeekDay - 1)
    End If
    For i = 0 To 41
        lblNumbers(i).Caption = ""
        lblNumbers(i).Font.Bold = False
    Next
    ' Put days on control
    ' Set day to bold if it is the selected date
    ' Sunday
    If FirstDayWD = 1 Then
        J = 6
        For i = 1 To NumDays
            lblNumbers(J).Caption = i
            If Day(Date) = i And GetMonth = Month(Date) Then
                lblNumbers(J).Font.Bold = True
                CurrentDay = J
            End If
            J = J + 1
        Next
    ' Monday
    ElseIf FirstDayWD = 2 Then
        J = 0
        For i = 1 To NumDays
            lblNumbers(J).Caption = i
            If Day(Date) = i And GetMonth = Month(Date) Then
                lblNumbers(J).Font.Bold = True
                CurrentDay = J
            End If
            J = J + 1
        Next
    ' Tuesday
    ElseIf FirstDayWD = 3 Then
        J = 1
        For i = 1 To NumDays
            lblNumbers(J).Caption = i
            If Day(Date) = i And GetMonth = Month(Date) Then
                lblNumbers(J).Font.Bold = True
                CurrentDay = J
            End If
            J = J + 1
        Next
    ' Wednesday
    ElseIf FirstDayWD = 4 Then
        J = 2
        For i = 1 To NumDays
            lblNumbers(J).Caption = i
            If Day(Date) = i And GetMonth = Month(Date) Then
                lblNumbers(J).Font.Bold = True
                CurrentDay = J
            End If
            J = J + 1
        Next
    ' Thursday
    ElseIf FirstDayWD = 5 Then
        J = 3
        For i = 1 To NumDays
            lblNumbers(J).Caption = i
            If Day(Date) = i And GetMonth = Month(Date) Then
                lblNumbers(J).Font.Bold = True
                CurrentDay = J
            End If
            J = J + 1
        Next
    ' Friday
    ElseIf FirstDayWD = 6 Then
        J = 4
        For i = 1 To NumDays
            lblNumbers(J).Caption = i
            If Day(Date) = i And GetMonth = Month(Date) Then
                lblNumbers(J).Font.Bold = True
                CurrentDay = J
            End If
            J = J + 1
        Next
    ' Saturday
    ElseIf FirstDayWD = 7 Then
        J = 5
        For i = 1 To NumDays
            lblNumbers(J).Caption = i
            If Day(Date) = i And GetMonth = Month(Date) Then
                lblNumbers(J).Font.Bold = True
                CurrentDay = J
            End If
            J = J + 1
        Next
    End If
    ' Put circle around currend day
    If GetMonth = Month(Date) Then
        shpDayNow.Visible = True
        shpDayNow.Left = lblNumbers(CurrentDay).Left - 30
        shpDayNow.Top = lblNumbers(CurrentDay).Top - 15
        shpDayNow.ZOrder 0
    Else
        shpDayNow.Visible = False
    End If
End Sub

'=====================================================
' Dropdown usercontrol (simulates combobox)
'=====================================================
Private Sub btnCombo_Click()
    If picCalendar.Visible = False Then
        picCalendar.Visible = True
        UserControl.Height = picCalendar.Top + picCalendar.Height
        If shpBorder.Width < picCalendar.Width Then
            UserControl.Width = picCalendar.Width
        End If
    Else
        UserControl.Width = shpBorder.Width
        picCalendar.Visible = False
        picMonthList.Visible = False
        UserControl.Height = shpBorder.Height
    End If
End Sub

'=====================================================
' Button Up/Down Calendar(simulates button combo)
'=====================================================
Private Sub btnCombo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnCombo, ShiftColors(DropDnButtonBackColor, 170), DropDnButtonBackColor, False, False, DropDnButtonForeColor, DropDnButtonBorderColor
End Sub
Private Sub btnCombo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnCombo, ShiftColors(DropDnButtonBackColor, 170), DropDnButtonBackColor, True, False, DropDnButtonForeColor, DropDnButtonBorderColor
End Sub

'=====================================================
' Button Month (simulates button combo)
'=====================================================
Private Sub btnMonthDown_Click()
    If picMonthList.Visible = False Then
        picYear.Enabled = False
        picMonthList.Visible = True
        For i = 0 To 41
            lblNumbers(i).Enabled = False
        Next i
    Else
        picYear.Enabled = True
        picMonthList.Visible = False
        For i = 0 To 41
            lblNumbers(i).Enabled = True
        Next i
    End If
End Sub
Private Sub btnMonthDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnMonthDown, ShiftColors(m_DropDnMonthButtonBackColor, 170), m_DropDnMonthButtonBackColor, False, False, m_DropDnMonthButtonForeColor, m_DropDnMonthButtonBorderColor
End Sub
Private Sub btnMonthDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnMonthDown, ShiftColors(m_DropDnMonthButtonBackColor, 170), m_DropDnMonthButtonBackColor, True, False, m_DropDnMonthButtonForeColor, m_DropDnMonthButtonBorderColor
End Sub

'=====================================================
' List Month (simulates combolist)
'=====================================================
Private Sub lblMonth_Click(Index As Integer)
    For i = 0 To 11
        lblMonth(i).BackStyle = 0
        lblMonth(i).BackColor = m_MonthListBackColor
        lblMonth(i).ForeColor = m_MonthListForeColor
    Next i
    lblMonth(Index).BackStyle = 1
    lblMonth(Index).ForeColor = m_MonthListSelectedForeColor
    lblMonth(Index).BackColor = m_MonthListSelectedBackColor
    MonthNow = Index
    picMonthList.Refresh
    txtMonth = lblMonth(Index).Caption
    For i = 0 To 41
        lblNumbers(i).BackStyle = 0
        lblNumbers(i).ForeColor = m_DaysForeColor
    Next i
    DoEvents
    Dim MonthNr As String
    If Index < 9 Then
        MonthNr = "0" & Index + 1
    Else
        MonthNr = Index + 1
    End If
    If m_DateFormat = 0 Then
        txtXtext = Left$(txtXtext, 2) & "\" & MonthNr & "\" & txtSpin
    Else
        txtXtext = MonthNr & "\" & Mid$(txtXtext, 4, 2) & "\" & txtSpin
    End If
    If UserControl.Ambient.DisplayName = "DMdatepicker8" Then
        StartDate = StartDate
    End If
    For i = 0 To 41
        If Val(lblNumbers(i)) = Val(Mid$(txtXtext, DayPos, 2)) Then lblNumbers_Click (i)
    Next i
End Sub
Private Sub lblMonth_DblClick(Index As Integer)
    lblMonth_Click Index
    picYear.Enabled = True
    picMonthList.Visible = False
    For i = 0 To 41
        lblNumbers(i).Enabled = True
    Next i
End Sub
Private Sub txtMonth_Change()
    If Trim$(txtMonth.Caption) <> "" Then Call CalculateCalendar
End Sub


'=====================================================
' Button Down Year (simulates spinbutton down)
'=====================================================
Private Sub btnSpinDown_Click()
    If Val(txtSpin.Caption) - 1 > 1950 Then txtSpin.Caption = Val(txtSpin.Caption) - 1
End Sub
Private Sub btnSpinDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnSpinDown, ShiftColors(m_SpinYearButtonBackColor, 170), m_SpinYearButtonBackColor, False, False, m_SpinYearButtonForeColor, m_SpinYearButtonBorderColor
    tmrDown.Enabled = True
End Sub
Private Sub btnSpinDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnSpinDown, ShiftColors(m_SpinYearButtonBackColor, 170), m_SpinYearButtonBackColor, True, False, m_SpinYearButtonForeColor, m_SpinYearButtonBorderColor
    tmrDown.Enabled = False
End Sub
Private Sub tmrDown_Timer()
    If Val(txtSpin.Caption) - 1 > 1950 Then txtSpin.Caption = Val(txtSpin.Caption) - 1
End Sub


'=====================================================
' Button Up Year (simulates spinbutton up)
'=====================================================
Private Sub btnSpinUp_Click()
    If Val(txtSpin.Caption) + 1 > 1950 Then txtSpin.Caption = Val(txtSpin.Caption) + 1
End Sub
Private Sub btnSpinUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnSpinUp, ShiftColors(m_SpinYearButtonBackColor, 170), m_SpinYearButtonBackColor, False, True, m_SpinYearButtonForeColor, m_SpinYearButtonBorderColor
    tmrUp.Enabled = True
End Sub
Private Sub btnSpinUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnSpinUp, ShiftColors(m_SpinYearButtonBackColor, 170), m_SpinYearButtonBackColor, True, True, m_SpinYearButtonForeColor, m_SpinYearButtonBorderColor
    tmrUp.Enabled = False
End Sub
Private Sub tmrUp_Timer()
    If Val(txtSpin.Caption) + 1 > 1950 Then txtSpin.Caption = Val(txtSpin.Caption) + 1
End Sub
Private Sub txtSpin_Change()
    If Trim$(txtSpin.Caption) <> "" Then Call CalculateCalendar
End Sub

'=====================================================
' Draw buttons
'=====================================================
Public Sub DrawButton(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, btnUp As Boolean, ArrowUp As Boolean, TextColor As OLE_COLOR, BordersColor As OLE_COLOR)  'Horizontal gradient
    On Error Resume Next
    DoEvents
    OldScaleMode = ctlControl.ScaleMode
    ctlControl.ScaleMode = 3
    If m_Style = Flat Or EndColor = &H8000000F Then
        If btnUp = True Then
            ctlControl.BackColor = EndColor
        Else
            ctlControl.BackColor = ShiftColors(EndColor, -50)
        End If
       GoTo DrawArrows
    End If
    If btnUp = True Then
        Call InitializeCol(ctlControl, StartColor, EndColor, False)
    Else
        Call InitializeCol(ctlControl, EndColor, StartColor, False)
    End If
    For i = 0 To ctlControl.ScaleHeight
        NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
        ctlControl.Line (0, i)-(ctlControl.ScaleWidth, i), NewColor
    Next
    DoEvents
DrawArrows:
    MidX = Round(ctlControl.ScaleWidth / 2)
    MidY = Round(ctlControl.ScaleHeight / 2)
    If ArrowUp = True Then
        ctlControl.Line (MidX - 3, MidY)-(MidX + 3, MidY), TextColor
        ctlControl.Line (MidX - 2, MidY - 1)-(MidX + 2, MidY - 1), TextColor
        ctlControl.Line (MidX - 1, MidY - 2)-(MidX + 1, MidY - 2), TextColor
        ctlControl.Line (MidX, MidY - 3)-(MidX, MidY - 3), TextColor
    Else
        ctlControl.Line (MidX - 3, MidY - 2)-(MidX + 3, MidY - 2), TextColor
        ctlControl.Line (MidX - 2, MidY - 1)-(MidX + 2, MidY - 1), TextColor
        ctlControl.Line (MidX - 1, MidY)-(MidX + 1, MidY), TextColor
        ctlControl.Line (MidX, MidY + 1)-(MidX, MidY + 1), TextColor
    End If
DrawBorders:
    ctlControl.Line (0, 0)-(ctlControl.ScaleWidth, 0), BordersColor
    ctlControl.Line (0, ctlControl.ScaleHeight - 1)-(ctlControl.ScaleWidth, ctlControl.ScaleHeight - 1), BordersColor
    ctlControl.Line (0, 0)-(0, ctlControl.ScaleHeight), BordersColor
    ctlControl.Line (ctlControl.ScaleWidth - 1, 0)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight), BordersColor
    ctlControl.Refresh
    ctlControl.ScaleMode = OldScaleMode
End Sub
'=====================================================
' Draw buttons with arrows
'=====================================================
Public Sub DrawPicBack(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, TextColor As OLE_COLOR)  'Horizontal gradient
    On Error Resume Next
    If m_Style = Flat Or EndColor = &H8000000F Then
       ctlControl.BackColor = EndColor
       Exit Sub
    End If
    Call InitializeCol(ctlControl, StartColor, EndColor, False)
    DoEvents
    OldScaleMode = ctlControl.ScaleMode
    ctlControl.ScaleMode = 3
    For i = 0 To ctlControl.ScaleHeight
        NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
        ctlControl.Line (0, i)-(ctlControl.ScaleWidth, i), NewColor
    Next
    ctlControl.Refresh
    ctlControl.ScaleMode = OldScaleMode
End Sub

'=====================================================
' Initialize colors for usercontrol
'=====================================================
Function InitializeCol(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, Clear As Boolean)
    OldScaleMode = ctlControl.ScaleMode
    ctlControl.ScaleMode = 3
    StartCol = StartColor
    EndCol = EndColor
    RedStart = StartCol Mod 256
    RedEnd = EndCol Mod 256
    RedI = (RedEnd - RedStart) / (ctlControl.ScaleHeight)
    GreenStart = (StartCol And &HFF00FF00) / 256
    GreenEnd = (EndCol And &HFF00FF00) / 256
    GreenI = (GreenEnd - GreenStart) / (ctlControl.ScaleHeight)
    BlueStart = (StartCol And &HFFFF0000) / (65536)
    BlueEnd = (EndCol And &HFFFF0000) / (65536)
    BlueI = (BlueEnd - BlueStart) / (ctlControl.ScaleHeight)
    ctlControl.ScaleMode = OldScaleMode
    If Clear = True Then ctlControl.Cls
End Function

'=====================================================
' Shift colors
'=====================================================
Private Function ShiftColors(ByVal MyColor As Long, ByVal Base As Long) As Long
    Dim R As Long, G As Long, B As Long, Delta As Long
    R = (MyColor And &HFF)
    G = ((MyColor \ &H100) Mod &H100)
    B = ((MyColor \ &H10000) Mod &H100)
    Delta = &HFF - Base
    B = Base + B * Delta \ &HFF
    G = Base + G * Delta \ &HFF
    R = Base + R * Delta \ &HFF
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255
    ShiftColors = R + 256& * G + 65536 * B
End Function

'=====================================================
' Set Array that's used for dayheadings and daynames based on the FirstWeekDay
'=====================================================
Private Sub SetFirstDayArray()
    FirstDayPos(1, 1) = 6
    FirstDayPos(1, 2) = 0
    FirstDayPos(1, 3) = 1
    FirstDayPos(1, 4) = 2
    FirstDayPos(1, 5) = 3
    FirstDayPos(1, 6) = 4
    FirstDayPos(1, 7) = 5
    FirstDayPos(2, 1) = 0
    FirstDayPos(2, 2) = 1
    FirstDayPos(2, 3) = 2
    FirstDayPos(2, 4) = 3
    FirstDayPos(2, 5) = 4
    FirstDayPos(2, 6) = 5
    FirstDayPos(2, 7) = 6
    FirstDayPos(3, 1) = 1
    FirstDayPos(3, 2) = 2
    FirstDayPos(3, 3) = 3
    FirstDayPos(3, 4) = 4
    FirstDayPos(3, 5) = 5
    FirstDayPos(3, 6) = 6
    FirstDayPos(3, 7) = 0
    FirstDayPos(4, 1) = 2
    FirstDayPos(4, 2) = 3
    FirstDayPos(4, 3) = 4
    FirstDayPos(4, 4) = 5
    FirstDayPos(4, 5) = 6
    FirstDayPos(4, 6) = 0
    FirstDayPos(4, 7) = 1
    FirstDayPos(5, 1) = 3
    FirstDayPos(5, 2) = 4
    FirstDayPos(5, 3) = 5
    FirstDayPos(5, 4) = 6
    FirstDayPos(5, 5) = 0
    FirstDayPos(5, 6) = 1
    FirstDayPos(5, 7) = 2
    FirstDayPos(6, 1) = 4
    FirstDayPos(6, 2) = 5
    FirstDayPos(6, 3) = 6
    FirstDayPos(6, 4) = 0
    FirstDayPos(6, 5) = 1
    FirstDayPos(6, 6) = 2
    FirstDayPos(6, 7) = 3
    FirstDayPos(7, 1) = 5
    FirstDayPos(7, 2) = 6
    FirstDayPos(7, 3) = 0
    FirstDayPos(7, 4) = 1
    FirstDayPos(7, 5) = 2
    FirstDayPos(7, 6) = 3
    FirstDayPos(7, 7) = 4
End Sub

'=====================================================
' Set language for monthnames and daynames
'=====================================================
Private Sub SetLanguage()
    Select Case m_Language
        Case 0 'Englisch
            lblToday = "Today: " & Format(Now, "dd/mm/yyyy")
            ArrMonth(0, 0) = "January"
            ArrMonth(0, 1) = "February"
            ArrMonth(0, 2) = "March"
            ArrMonth(0, 3) = "April"
            ArrMonth(0, 4) = "May"
            ArrMonth(0, 5) = "June"
            ArrMonth(0, 6) = "July"
            ArrMonth(0, 7) = "August"
            ArrMonth(0, 8) = "September"
            ArrMonth(0, 9) = "October"
            ArrMonth(0, 10) = "November"
            ArrMonth(0, 11) = "December"
            ArrDay(0, 0) = "M"
            ArrDay(0, 1) = "T"
            ArrDay(0, 2) = "W"
            ArrDay(0, 3) = "T"
            ArrDay(0, 4) = "F"
            ArrDay(0, 5) = "S"
            ArrDay(0, 6) = "S"
            ArrDayName(0, 0) = "Monday"
            ArrDayName(0, 1) = "Tuesday"
            ArrDayName(0, 2) = "Wednesday"
            ArrDayName(0, 3) = "Thursday"
            ArrDayName(0, 4) = "Friday"
            ArrDayName(0, 5) = "Saterday"
            ArrDayName(0, 6) = "Sunday"
        Case 1 'Nederlands
            lblToday = "Vandaag: " & Format(Now, "dd/mm/yyyy")
            ArrMonth(1, 0) = "Januari"
            ArrMonth(1, 1) = "Februari"
            ArrMonth(1, 2) = "Maart"
            ArrMonth(1, 3) = "April"
            ArrMonth(1, 4) = "Mei"
            ArrMonth(1, 5) = "Juni"
            ArrMonth(1, 6) = "Juli"
            ArrMonth(1, 7) = "Augustus"
            ArrMonth(1, 8) = "September"
            ArrMonth(1, 9) = "Oktober"
            ArrMonth(1, 10) = "November"
            ArrMonth(1, 11) = "December"
            ArrDay(1, 0) = "M"
            ArrDay(1, 1) = "D"
            ArrDay(1, 2) = "W"
            ArrDay(1, 3) = "D"
            ArrDay(1, 4) = "V"
            ArrDay(1, 5) = "Z"
            ArrDay(1, 6) = "Z"
            ArrDayName(1, 0) = "Maandag"
            ArrDayName(1, 1) = "Dinsdag"
            ArrDayName(1, 2) = "Woensdag"
            ArrDayName(1, 3) = "Donderdag"
            ArrDayName(1, 4) = "Vrijdag"
            ArrDayName(1, 5) = "Zaterdag"
            ArrDayName(1, 6) = "Zondag"
        Case 2 'Francais
            lblToday = "Aujourd'hui: " & Format(Now, "dd/mm/yyyy")
            ArrMonth(2, 0) = "Janvier"
            ArrMonth(2, 1) = "Février"
            ArrMonth(2, 2) = "Mars"
            ArrMonth(2, 3) = "Avril"
            ArrMonth(2, 4) = "Mai"
            ArrMonth(2, 5) = "Juin"
            ArrMonth(2, 6) = "Juillet"
            ArrMonth(2, 7) = "Août"
            ArrMonth(2, 8) = "Septembre"
            ArrMonth(2, 9) = "Octobre"
            ArrMonth(2, 10) = "Novembre"
            ArrMonth(2, 11) = "Décembre"
            ArrDay(2, 0) = "L"
            ArrDay(2, 1) = "M"
            ArrDay(2, 2) = "M"
            ArrDay(2, 3) = "J"
            ArrDay(2, 4) = "V"
            ArrDay(2, 5) = "S"
            ArrDay(2, 6) = "D"
            ArrDayName(2, 0) = "Lundi"
            ArrDayName(2, 1) = "Mardi"
            ArrDayName(2, 2) = "Mercredi"
            ArrDayName(2, 3) = "Jeudi"
            ArrDayName(2, 4) = "Vendredi"
            ArrDayName(2, 5) = "Samedi"
            ArrDayName(2, 6) = "Dimanche"
        Case 3 'Deutch
            lblToday = "Heute: " & Format(Now, "dd/mm/yyyy")
            ArrMonth(3, 0) = "Januar"
            ArrMonth(3, 1) = "Februar"
            ArrMonth(3, 2) = "März"
            ArrMonth(3, 3) = "April"
            ArrMonth(3, 4) = "Mai"
            ArrMonth(3, 5) = "Juni"
            ArrMonth(3, 6) = "Juli"
            ArrMonth(3, 7) = "August"
            ArrMonth(3, 8) = "September"
            ArrMonth(3, 9) = "OKtober"
            ArrMonth(3, 10) = "November"
            ArrMonth(3, 11) = "Dezember"
            ArrDay(3, 0) = "M"
            ArrDay(3, 1) = "D"
            ArrDay(3, 2) = "W"
            ArrDay(3, 3) = "D"
            ArrDay(3, 4) = "F"
            ArrDay(3, 5) = "S"
            ArrDay(3, 6) = "S"
            ArrDayName(3, 0) = "Montag"
            ArrDayName(3, 1) = "Dienstag"
            ArrDayName(3, 2) = "Mittwoch"
            ArrDayName(3, 3) = "Donnerstag"
            ArrDayName(3, 4) = "Freitag"
            ArrDayName(3, 5) = "Samstag"
            ArrDayName(3, 6) = "Sonntag"
       Case 4 'Italiano
            lblToday = "Oggi: " & Format(Now, "dd/mm/yyyy")
            ArrMonth(4, 0) = "Gennaio "
            ArrMonth(4, 1) = "Febbraio "
            ArrMonth(4, 2) = "Marzo "
            ArrMonth(4, 3) = "Aprile"
            ArrMonth(4, 4) = "Maggio"
            ArrMonth(4, 5) = "Giugno "
            ArrMonth(4, 6) = "Luglio "
            ArrMonth(4, 7) = "Agosto "
            ArrMonth(4, 8) = "Settembre "
            ArrMonth(4, 9) = "Ottobre "
            ArrMonth(4, 10) = "Novembre "
            ArrMonth(4, 11) = "Dicembre "
            ArrDay(4, 0) = "L"
            ArrDay(4, 1) = "M"
            ArrDay(4, 2) = "M"
            ArrDay(4, 3) = "G"
            ArrDay(4, 4) = "V"
            ArrDay(4, 5) = "S"
            ArrDay(4, 6) = "D"
            ArrDayName(4, 0) = "Lunedì"
            ArrDayName(4, 1) = "Martedì"
            ArrDayName(4, 2) = "Mercoledì"
            ArrDayName(4, 3) = "Giovedì"
            ArrDayName(4, 4) = "Venerdì"
            ArrDayName(4, 5) = "Sabato"
            ArrDayName(4, 6) = "Domenica"
        Case 5 'Espagnol
            lblToday = "Hoy: " & Format(Now, "dd/mm/yyyy")
            ArrMonth(5, 0) = "Enero"
            ArrMonth(5, 1) = "Febrero"
            ArrMonth(5, 2) = "Marzo "
            ArrMonth(5, 3) = "Abril"
            ArrMonth(5, 4) = "Mayo"
            ArrMonth(5, 5) = "Junio"
            ArrMonth(5, 6) = "Julio"
            ArrMonth(5, 7) = "Agosto"
            ArrMonth(5, 8) = "Septiembre"
            ArrMonth(5, 9) = "Octubre "
            ArrMonth(5, 10) = "Noviembre "
            ArrMonth(5, 11) = "Diciembre "
            ArrDay(5, 0) = "L"
            ArrDay(5, 1) = "M"
            ArrDay(5, 2) = "M"
            ArrDay(5, 3) = "J"
            ArrDay(5, 4) = "V"
            ArrDay(5, 5) = "S"
            ArrDay(5, 6) = "D"
            ArrDayName(5, 0) = "Lunes"
            ArrDayName(5, 1) = "Martes"
            ArrDayName(5, 2) = "Miércoles"
            ArrDayName(5, 3) = "Jueves"
            ArrDayName(5, 4) = "Viernes"
            ArrDayName(5, 5) = "Sábado"
            ArrDayName(5, 6) = "Domingo"
    End Select
    ' Switch headinglabels for days and daynames based on m_FirstWeekDay
    ' via temporary array
    ' Also see the sub SetFirstDayArray
    Dim ArrTmpD(7) As String
    Dim ArrTmpDN(7) As String
    ' Fill the temporary array with standard values
    For i = 0 To 6
        ArrTmpD(i) = ArrDay(m_Language, i)
        lblDay(i).Tag = i
    Next i
    For i = 0 To 6
        ArrTmpDN(i) = ArrDayName(m_Language, i)
    Next i
    ' switch standard values to the temporary
    For i = 0 To 6
        ArrDay(m_Language, 0) = ArrTmpD(FirstDayPos(m_FirstWeekDay, 1))
        ArrDay(m_Language, 1) = ArrTmpD(FirstDayPos(m_FirstWeekDay, 2))
        ArrDay(m_Language, 2) = ArrTmpD(FirstDayPos(m_FirstWeekDay, 3))
        ArrDay(m_Language, 3) = ArrTmpD(FirstDayPos(m_FirstWeekDay, 4))
        ArrDay(m_Language, 4) = ArrTmpD(FirstDayPos(m_FirstWeekDay, 5))
        ArrDay(m_Language, 5) = ArrTmpD(FirstDayPos(m_FirstWeekDay, 6))
        ArrDay(m_Language, 6) = ArrTmpD(FirstDayPos(m_FirstWeekDay, 7))
        ArrDayName(m_Language, 0) = ArrTmpDN(FirstDayPos(m_FirstWeekDay, 1))
        ArrDayName(m_Language, 1) = ArrTmpDN(FirstDayPos(m_FirstWeekDay, 2))
        ArrDayName(m_Language, 2) = ArrTmpDN(FirstDayPos(m_FirstWeekDay, 3))
        ArrDayName(m_Language, 3) = ArrTmpDN(FirstDayPos(m_FirstWeekDay, 4))
        ArrDayName(m_Language, 4) = ArrTmpDN(FirstDayPos(m_FirstWeekDay, 5))
        ArrDayName(m_Language, 5) = ArrTmpDN(FirstDayPos(m_FirstWeekDay, 6))
        ArrDayName(m_Language, 6) = ArrTmpDN(FirstDayPos(m_FirstWeekDay, 7))
    Next i
    ' Update monthnames in list
    For i = 0 To 11
        lblMonth(i).Caption = ArrMonth(Language, i)
    Next i
    ' Update daynames
    For i = 0 To 6
        lblDay(i) = ArrDay(Language, i)
    Next i
End Sub

'=====================================================
' Set colors for every control on the usercontrol
'=====================================================
Private Sub SetCalendarColors()
    ' Drop down
    If Enabled = False Then
        shpBorder.BorderColor = m_DisabledBorderColor
        shpBorder.FillColor = m_DisabledBackColor
        txtXtext.BackColor = m_DisabledBackColor
        txtXtext.ForeColor = m_DisabledForeColor
        DrawPicBack picDropDown, ShiftColors(m_DisabledBackColor, 150), m_DisabledBackColor, m_DisabledForeColor
        DrawButton btnCombo, ShiftColors(m_DisabledBackColor, 100), m_DisabledBackColor, True, False, m_DisabledForeColor, m_DisabledBorderColor
        Exit Sub
    Else
        UserControl.Enabled = True
        shpBorder.BorderColor = m_DropDnBorderColor
        shpBorder.FillColor = m_DropDnBackColor
        txtXtext.BackColor = m_DropDnBackColor
        txtXtext.ForeColor = m_DropDnForeColor
        DrawPicBack picDropDown, ShiftColors(m_DropDnBackColor, 150), m_DropDnBackColor, m_DropDnForeColor
        DrawButton btnCombo, ShiftColors(m_DropDnButtonBackColor, 100), m_DropDnButtonBackColor, True, False, m_DropDnButtonForeColor, m_DropDnButtonBorderColor
    End If
    ' Calendar
    shpCalendar.BorderColor = m_CalendarBorderColor
    shpCalendar.BackColor = m_CalendarBackColor
    DrawPicBack picCalendar, ShiftColors(m_CalendarBackColor, 170), m_CalendarBackColor, m_CalendarForeColor
    ' Month Drop down
    shpMonth.BorderColor = m_DropDnMonthBorderColor
    'picMonth.BackColor = m_DropDnMonthBackColor
    txtMonth.ForeColor = m_DropDnMonthForeColor
    DrawPicBack picMonth, ShiftColors(m_DropDnMonthBackColor, 170), m_DropDnMonthBackColor, m_DropDnMonthForeColor
    DrawButton btnMonthDown, ShiftColors(m_DropDnMonthButtonBackColor, 170), m_DropDnMonthButtonBackColor, True, False, m_DropDnMonthButtonForeColor, m_DropDnMonthButtonBorderColor
    ' Month List
    shpMonthNames.BorderColor = m_MonthListBorderColor
    DrawPicBack picMonthList, ShiftColors(m_MonthListBackColor, 170), m_MonthListBackColor, m_MonthListForeColor
    picMonthList.ForeColor = m_MonthListForeColor
    For i = 0 To 11
        lblMonth(i).ForeColor = m_MonthListForeColor
    Next i
    ' Days
    shpDays.BorderColor = m_DaysBorderColor
    DrawPicBack picDays, ShiftColors(m_DayNamesBackColor, 60), m_DayNamesBackColor, m_DayNamesForeColor
    DrawPicBack picDayNames, ShiftColors(m_DayNamesBackColor, 170), m_DayNamesBackColor, m_DayNamesForeColor
    shpDayNamesBorder.BorderColor = m_DayNamesBorderColor
    shpDayNow.BorderColor = m_CurrentDayBOrderColor
    lblToday.ForeColor = m_CalendarForeColor
    ' Year
    shpYear.BorderColor = m_SpinYearBorderColor
    DrawButton btnSpinUp, ShiftColors(m_SpinYearButtonBackColor, 170), m_SpinYearButtonBackColor, True, True, m_SpinYearButtonForeColor, m_SpinYearButtonBorderColor
    DrawButton btnSpinDown, ShiftColors(m_SpinYearButtonBackColor, 170), m_SpinYearButtonBackColor, True, False, m_SpinYearButtonForeColor, m_SpinYearButtonBorderColor
    DrawPicBack picYear, ShiftColors(m_SpinYearBackColor, 170), m_SpinYearBackColor, m_SpinYearForeColor
    txtSpin.ForeColor = m_SpinYearForeColor
    ' days
    For i = 0 To 6
        lblDay(i).ForeColor = m_DayNamesForeColor
        lblDay(i).BackColor = m_DayNamesBackColor
    Next i
    For i = 0 To 41
        lblNumbers(i).ForeColor = m_DaysForeColor
    Next i
    DrawPicBack picDays, ShiftColors(m_DaysBackColor, 60), m_DaysBackColor, m_DaysForeColor
End Sub


'=====================================================
' Usercontrol handling
'=====================================================
Private Sub UserControl_Resize()
    If picCalendar.Visible = False Then
        shpBorder.Width = UserControl.Width
        shpBorder.Height = UserControl.Height
        txtXtext.Width = (shpBorder.Width - 45) - btnCombo.Width
        picDropDown.Width = shpBorder.Width
        picDropDown.Height = UserControl.Height
        txtXtext.Top = (shpBorder.Height / 2) - (txtXtext.Height / 2) + 15
        picCalendar.Top = shpBorder.Height - 15
    Else
        UserControl.Height = shpBorder.Height + shpCalendar.Height + 30
    End If
    btnCombo.Height = shpBorder.Height '- 30
    btnCombo.Top = 0 '15
    btnCombo.Left = (shpBorder.Width - btnCombo.Width) '- 15
    If m_Enabled = True Then
        DrawButton btnCombo, ShiftColors(DropDnButtonBackColor, 170), DropDnButtonBackColor, True, False, DropDnButtonForeColor, DropDnButtonBorderColor
        DrawPicBack picDropDown, ShiftColors(m_DropDnBackColor, 150), m_DropDnBackColor, m_DropDnForeColor
    Else
        DrawPicBack picDropDown, ShiftColors(m_DisabledBackColor, 150), m_DisabledBackColor, m_DisabledForeColor
        DrawButton btnCombo, ShiftColors(m_DisabledBackColor, 100), m_DisabledBackColor, True, False, m_DisabledForeColor, m_DisabledBorderColor
    End If
End Sub

'=============================================================================================
' Usercontrol properties
'=============================================================================================

'=====================================================
' InitProperties
'=====================================================
Private Sub UserControl_InitProperties()
    'Calendar
    picCalendar.Height = 2890
    picCalendar.Width = 2730
    picCalendar.Visible = False
    ' Month
    picMonth.Height = 285
    picMonth.Width = 1335
    picMonth.Visible = False
    picMonthList.Height = 2500
    picMonthList.Width = 1335
    picMonthList.Visible = False
    ' DropDown
    m_DropDnForeColor = m_def_DropDnForeColor
    m_DropDnBackColor = m_def_DropDnBackColor
    m_DropDnBorderColor = m_def_DropDnBorderColor
    m_DropDnButtonBackColor = m_def_DropDnButtonBackColor
    m_DropDnButtonForeColor = m_def_DropDnButtonForeColor
    m_DropDnButtonBorderColor = m_def_DropDnButtonBorderColor
    ' Calendar
    m_CalendarBackColor = m_def_CalendarBackColor
    m_CalendarBorderColor = m_def_CalendarBorderColor
    m_CalendarForeColor = m_def_CalendarForeColor
    ' Year
    m_SpinYearForeColor = m_def_SpinYearForeColor
    m_SpinYearBackColor = m_def_SpinYearBackColor
    m_SpinYearBorderColor = m_def_SpinYearBorderColor
    m_SpinYearButtonBackColor = m_def_SpinYearButtonBackColor
    m_SpinYearButtonForeColor = m_def_SpinYearButtonForeColor
    m_SpinYearButtonBorderColor = m_def_SpinYearButtonBorderColor
    ' Month Drop down
    m_DropDnMonthForeColor = m_def_DropDnMonthForeColor
    m_DropDnMonthBackColor = m_def_DropDnMonthBackColor
    m_DropDnMonthBorderColor = m_def_DropDnMonthBorderColor
    m_DropDnMonthButtonBackColor = m_def_DropDnMonthButtonBackColor
    m_DropDnMonthButtonForeColor = m_def_DropDnMonthButtonForeColor
    m_DropDnMonthButtonBorderColor = m_def_DropDnMonthButtonBorderColor
    ' Month List
    m_MonthListBackColor = m_def_MonthListBackColor
    m_MonthListForeColor = m_def_MonthListForeColor
    m_MonthListSelectedBackColor = m_def_MonthListSelectedBackColor
    m_MonthListSelectedForeColor = m_def_MonthListSelectedForeColor
    m_MonthListBorderColor = m_def_MonthListBorderColor
    m_Language = m_def_Language
    ' Days
    m_DaysForeColor = m_def_DaysForeColor
    m_DaysBackColor = m_def_DaysBackColor
    m_DaysBorderColor = m_def_DaysBorderColor
    m_DayNamesForeColor = m_def_DayNamesForeColor
    m_DayNamesBackColor = m_def_DayNamesBackColor
    m_DaySelectedBackColor = m_def_DaySelectedBackColor
    m_DaySelectedForeColor = m_def_DaySelectedForeColor
    m_DayNamesBorderColor = m_def_DayNamesBorderColor
    m_CurrentDayBOrderColor = m_def_CurrentDayBOrderColor
    ' Disabledcolos
    m_DisabledBackColor = m_def_DisabledBackColor
    m_DisabledBorderColor = m_def_DisabledBorderColor
    m_DisabledForeColor = m_def_DisabledForeColor
    m_Enabled = True
    
    m_Style = m_def_Style
    m_DateFormat = m_def_DateFormat
    txtXtext = Format(Date, "dd/mm/yyyy")
    m_FirstWeekDay = m_def_FirstWeekDay
End Sub

'=====================================================
' ReadProperties
'=====================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' DropDown
    m_DropDnForeColor = PropBag.ReadProperty("DropDnForeColor", m_def_DropDnForeColor)
    m_DropDnBackColor = PropBag.ReadProperty("DropDnBackColor", m_def_DropDnBackColor)
    m_DropDnBorderColor = PropBag.ReadProperty("DropDnBorderColor", m_def_DropDnBorderColor)
    m_DropDnButtonBackColor = PropBag.ReadProperty("DropDnButtonBackColor", m_def_DropDnButtonBackColor)
    m_DropDnButtonForeColor = PropBag.ReadProperty("DropDnButtonForeColor", m_def_DropDnButtonForeColor)
    m_DropDnButtonBorderColor = PropBag.ReadProperty("DropDnButtonBorderColor", m_def_DropDnButtonBorderColor)
    ' Calendar
    m_CalendarForeColor = PropBag.ReadProperty("CalendarForeColor", m_def_CalendarForeColor)
    m_CalendarBackColor = PropBag.ReadProperty("CalendarBackColor", m_def_CalendarBackColor)
    m_CalendarBorderColor = PropBag.ReadProperty("CalendarBorderColor", m_def_CalendarBorderColor)
    'Days
    m_DaysForeColor = PropBag.ReadProperty("DaysForeColor", m_def_DaysForeColor)
    m_DaysBackColor = PropBag.ReadProperty("DaysBackColor", m_def_DaysBackColor)
    m_DaysBorderColor = PropBag.ReadProperty("DaysBorderColor", m_def_DaysBorderColor)
    m_DayNamesForeColor = PropBag.ReadProperty("DayNamesForeColor", m_def_DayNamesForeColor)
    m_DayNamesBackColor = PropBag.ReadProperty("DayNamesBackColor", m_def_DayNamesBackColor)
    m_DaySelectedBackColor = PropBag.ReadProperty("DaySelectedBackColor", m_def_DaySelectedBackColor)
    m_DaySelectedForeColor = PropBag.ReadProperty("DaySelectedForeColor", m_def_DaySelectedForeColor)
    m_CurrentDayBOrderColor = PropBag.ReadProperty("CurrentDayBOrderColor", m_def_CurrentDayBOrderColor)
    ' Daynames
    m_DayNamesBorderColor = PropBag.ReadProperty("DayNamesBorderColor", m_def_DayNamesBorderColor)
    ' Month Drop down
    m_DropDnMonthForeColor = PropBag.ReadProperty("DropDnMonthForeColor", m_def_DropDnMonthForeColor)
    m_DropDnMonthBackColor = PropBag.ReadProperty("DropDnMonthBackColor", m_def_DropDnMonthBackColor)
    m_DropDnMonthBorderColor = PropBag.ReadProperty("DropDnMonthBorderColor", m_def_DropDnMonthBorderColor)
    m_DropDnMonthButtonBackColor = PropBag.ReadProperty("DropDnMonthButtonBackColor", m_def_DropDnMonthButtonBackColor)
    m_DropDnMonthButtonForeColor = PropBag.ReadProperty("DropDnMonthButtonForeColor", m_def_DropDnMonthButtonForeColor)
    m_DropDnMonthButtonBorderColor = PropBag.ReadProperty("DropDnMonthButtonBorderColor", m_def_DropDnMonthButtonBorderColor)
    ' Month List
    m_MonthListBackColor = PropBag.ReadProperty("MonthListBackColor", m_def_MonthListBackColor)
    m_MonthListForeColor = PropBag.ReadProperty("MonthListForeColor", m_def_MonthListForeColor)
    m_MonthListSelectedBackColor = PropBag.ReadProperty("MonthListSelectedBackColor", m_def_MonthListSelectedBackColor)
    m_MonthListSelectedForeColor = PropBag.ReadProperty("MonthListSelectedForeColor", m_def_MonthListSelectedForeColor)
    m_MonthListBorderColor = PropBag.ReadProperty("MonthListBorderColor", m_def_MonthListBorderColor)
    ' Year
    m_SpinYearForeColor = PropBag.ReadProperty("SpinYearForeColor", m_def_SpinYearForeColor)
    m_SpinYearBackColor = PropBag.ReadProperty("SpinYearBackColor", m_def_SpinYearBackColor)
    m_SpinYearBorderColor = PropBag.ReadProperty("SpinYearBorderColor", m_def_SpinYearBorderColor)
    m_SpinYearButtonBackColor = PropBag.ReadProperty("SpinYearButtonBackColor", m_def_SpinYearButtonBackColor)
    m_SpinYearButtonForeColor = PropBag.ReadProperty("SpinYearButtonForeColor", m_def_SpinYearButtonForeColor)
    m_SpinYearButtonBorderColor = PropBag.ReadProperty("SpinYearButtonBorderColor", m_def_SpinYearButtonBorderColor)
    ' DisabledColors
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_DisabledBackColor = PropBag.ReadProperty("DisabledBackColor", m_def_DisabledBackColor)
    m_DisabledBorderColor = PropBag.ReadProperty("DisabledBorderColor", m_def_DisabledBorderColor)
    m_DisabledForeColor = PropBag.ReadProperty("DisabledforeColor", m_def_DisabledForeColor)
    m_StartDate = PropBag.ReadProperty("StartDate", "")
    m_Language = PropBag.ReadProperty("Language", m_def_Language)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_DateFormat = PropBag.ReadProperty("DateFormat", m_def_DateFormat)
    m_FirstWeekDay = PropBag.ReadProperty("FirstWeekDay", m_def_FirstWeekDay)
    If DateFormat = 0 Then
        DayPos = 1
        MonthPos = 4
    Else
        DayPos = 4
        MonthPos = 1
    End If
    ' DrawButtons
    SetFirstDayArray
    SetCalendarColors
    ' set monthnames and daynames
    SetLanguage
    ' get startdate
    If StartDate <> "" Then
        txtXtext = m_StartDate
    Else
        txtXtext = Format(Now, "dd/mm/yyyy")
    End If
    GetMonth = Val(Mid$(txtXtext, MonthPos, 2))
    txtMonth.Caption = ArrMonth(m_Language, GetMonth - 1)
    'lblMonth_Click (GetMonth - 1)  <== don't need to do this anymore
    lblMonth(GetMonth - 1).BackStyle = 1
    lblMonth(GetMonth - 1).ForeColor = m_MonthListSelectedForeColor
    lblMonth(GetMonth - 1).BackColor = m_MonthListSelectedBackColor
    ' calculate days in month and put them on the calendar
    CalculateCalendar
    If Val(Mid$(txtXtext, DayPos, 2)) > 0 Then
        For i = 0 To 41
            If Val(lblNumbers(i).Caption) = Val(Mid$(txtXtext, DayPos, 2)) Then
                lblNumbers(i).BackStyle = 1
                lblNumbers(i).BackColor = m_DaySelectedBackColor
                lblNumbers(i).ForeColor = m_DaySelectedForeColor
            End If
        Next i
    End If
    UserControl.Enabled = m_Enabled
End Sub

'=====================================================
' WriteProperties
'=====================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Days
    Call PropBag.WriteProperty("DaysForeColor", m_DaysForeColor, m_def_DaysForeColor)
    Call PropBag.WriteProperty("DaysBackColor", m_DaysBackColor, m_def_DaysBackColor)
    Call PropBag.WriteProperty("DaysBorderColor", m_DaysBorderColor, m_def_DaysBorderColor)
    Call PropBag.WriteProperty("DayNamesForeColor", m_DayNamesForeColor, m_def_DayNamesForeColor)
    Call PropBag.WriteProperty("DayNamesBackColor", m_DayNamesBackColor, m_def_DayNamesBackColor)
    Call PropBag.WriteProperty("DaySelectedBackColor", m_DaySelectedBackColor, m_def_DaySelectedBackColor)
    Call PropBag.WriteProperty("DaySelectedForeColor", m_DaySelectedForeColor, m_def_DaySelectedForeColor)
    Call PropBag.WriteProperty("DayNamesBorderColor", m_DayNamesBorderColor, m_def_DayNamesBorderColor)
    Call PropBag.WriteProperty("CurrentDayBOrderColor", m_CurrentDayBOrderColor, m_def_CurrentDayBOrderColor)
    ' Calendar
    Call PropBag.WriteProperty("CalendarBorderColor", m_CalendarBorderColor, m_def_CalendarBorderColor)
    Call PropBag.WriteProperty("CalendarBackColor", m_CalendarBackColor, m_def_CalendarBackColor)
    Call PropBag.WriteProperty("CalendarForeColor", m_CalendarForeColor, m_def_CalendarForeColor)
    ' Drop down
    Call PropBag.WriteProperty("DropDnForeColor", m_DropDnForeColor, m_def_DropDnForeColor)
    Call PropBag.WriteProperty("DropDnBackColor", m_DropDnBackColor, m_def_DropDnBackColor)
    Call PropBag.WriteProperty("DropDnButtonBackColor", m_DropDnButtonBackColor, m_def_DropDnButtonBackColor)
    Call PropBag.WriteProperty("DropDnButtonForeColor", m_DropDnButtonForeColor, m_def_DropDnButtonForeColor)
    Call PropBag.WriteProperty("DropDnButtonBorderColor", m_DropDnButtonBorderColor, m_def_DropDnButtonBorderColor)
    Call PropBag.WriteProperty("DropDnBorderColor", m_DropDnBorderColor, m_def_DropDnBorderColor)
    ' Month Drop down
    Call PropBag.WriteProperty("DropDnMonthForeColor", m_DropDnMonthForeColor, m_def_DropDnMonthForeColor)
    Call PropBag.WriteProperty("DropDnMonthBackColor", m_DropDnMonthBackColor, m_def_DropDnMonthBackColor)
    Call PropBag.WriteProperty("DropDnMonthBorderColor", m_DropDnMonthBorderColor, m_def_DropDnMonthBorderColor)
    Call PropBag.WriteProperty("DropDnMonthButtonBackColor", m_DropDnMonthButtonBackColor, m_def_DropDnMonthButtonBackColor)
    Call PropBag.WriteProperty("DropDnMonthButtonForeColor", m_DropDnMonthButtonForeColor, m_def_DropDnMonthButtonForeColor)
    Call PropBag.WriteProperty("DropDnMonthButtonBorderColor", m_DropDnMonthButtonBorderColor, m_def_DropDnMonthButtonBorderColor)
    ' Month List
    Call PropBag.WriteProperty("MonthListBackColor", m_MonthListBackColor, m_def_MonthListBackColor)
    Call PropBag.WriteProperty("MonthListForeColor", m_MonthListForeColor, m_def_MonthListForeColor)
    Call PropBag.WriteProperty("MonthListSelectedBackColor", m_MonthListSelectedBackColor, m_def_MonthListSelectedBackColor)
    Call PropBag.WriteProperty("MonthListSelectedForeColor", m_MonthListSelectedForeColor, m_def_MonthListSelectedForeColor)
    Call PropBag.WriteProperty("MonthListBorderColor", m_MonthListBorderColor, m_def_MonthListBorderColor)
    ' Year
    Call PropBag.WriteProperty("SpinYearForeColor", m_SpinYearForeColor, m_def_SpinYearForeColor)
    Call PropBag.WriteProperty("SpinYearBackColor", m_SpinYearBackColor, m_def_SpinYearBackColor)
    Call PropBag.WriteProperty("SpinYearBorderColor", m_SpinYearBorderColor, m_def_SpinYearBorderColor)
    Call PropBag.WriteProperty("SpinYearButtonBackColor", m_SpinYearButtonBackColor, m_def_SpinYearButtonBackColor)
    Call PropBag.WriteProperty("SpinYearButtonForeColor", m_SpinYearButtonForeColor, m_def_SpinYearButtonForeColor)
    Call PropBag.WriteProperty("SpinYearButtonBorderColor", m_SpinYearButtonBorderColor, m_def_SpinYearButtonBorderColor)
    ' DisabledColors
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("DisabledBackColor", m_DisabledBackColor, m_def_DisabledBackColor)
    Call PropBag.WriteProperty("DisabledForeColor", m_DisabledForeColor, m_def_DisabledForeColor)
    Call PropBag.WriteProperty("DisabledBorderColor", m_DisabledBorderColor, m_def_DisabledBorderColor)
    ' rest
    Call PropBag.WriteProperty("Language", m_Language, m_def_Language)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("StartDate", m_StartDate, "")
    Call PropBag.WriteProperty("DateFormat", m_DateFormat, m_def_DateFormat)
    Call PropBag.WriteProperty("FirstWeekDay", m_FirstWeekDay, m_def_FirstWeekDay)
End Sub

'=====================================================
' Get,Set and Let
'=====================================================

' Return the Text.
Public Property Get Text() As String
    Text = txtXtext.Caption
End Property

' Set the Text.
Public Property Let Text(ByVal New_Text As String)
    txtXtext.Caption() = New_Text
    PropertyChanged "Text"
End Property

' return the daynumber
Public Property Get SelectedDay() As String
    SelectedDay = Mid$(txtXtext.Caption, DayPos, 2)
End Property
Public Property Let SelectedDay(ByVal New_SelectedDay As String)
    PropertyChanged "SelectedDay"
End Property

' return the daynumber
Public Property Get SelectedDateShort() As String
    SelectedDateShort = Left(txtXtext.Caption, 6) & Right$(txtXtext.Caption, 2)
End Property
Public Property Let SelectedDateShort(ByVal New_SelectedDateShort As String)
    PropertyChanged "SelectedDateShort"
End Property

' return the dayname
Public Property Get SelectedDayName() As String
    Select Case DayNumber
        Case 0, 7, 14, 21, 28, 35
            SelectedDayName = ArrDayName(Language, 0)
        Case 1, 8, 15, 22, 29, 36
            SelectedDayName = ArrDayName(Language, 1)
        Case 2, 9, 16, 23, 30, 37
            SelectedDayName = ArrDayName(Language, 2)
        Case 3, 10, 17, 24, 31, 38
            SelectedDayName = ArrDayName(Language, 3)
        Case 4, 11, 18, 25, 32, 39
            SelectedDayName = ArrDayName(Language, 4)
        Case 5, 12, 19, 25, 33, 40
            SelectedDayName = ArrDayName(Language, 5)
        Case 6, 13, 20, 26, 34, 41
            SelectedDayName = ArrDayName(Language, 6)
    End Select
End Property
Public Property Let SelectedDayName(ByVal New_SelectedDayName As String)
    PropertyChanged "SelectedDayName"
End Property

' return the Monthnumber
Public Property Get SelectedMonth() As String
    SelectedMonth = Mid$(txtXtext.Caption, MonthPos, 2)
End Property
Public Property Let SelectedMonth(ByVal New_SelectedMonth As String)
    PropertyChanged "SelectedMonth"
End Property

' return the Monthname
Public Property Get SelectedMonthName() As String
    If Trim$(txtXtext.Caption) = "" Then
        SelectedMonthName = ""
        Exit Sub
    End If
    SelectedMonthName = ArrMonth(Language, Val(Mid$(txtXtext.Caption, MonthPos, 2)) - 1) '1)
End Property
Public Property Let SelectedMonthName(ByVal New_SelectedMonthName As String)
    PropertyChanged "SelectedMonthName"
End Property

' return the year
Public Property Get SelectedYear() As String
    If Trim$(txtXtext.Caption) = "" Then
        SelectedYear = ""
        Exit Sub
    End If
    SelectedYear = Right$(txtXtext.Caption, 4)
End Property
Public Property Let SelectedYear(ByVal New_SelectedYear As String)
    PropertyChanged "SelectedYear"
End Property

' return the short year
Public Property Get SelectedYearShort() As String
    SelectedYearShort = Right$(txtXtext.Caption, 2)
End Property
Public Property Let SelectedYearShort(ByVal New_SelectedYearShort As String)
    PropertyChanged "SelectedYearShort"
End Property

' return the full date
Public Property Get SelectedFullDate() As String
    If Trim$(txtXtext.Caption) = "" Then
        SelectedFullDate = ""
        Exit Sub
    End If
    If m_DateFormat = 0 Then
        SelectedFullDate = SelectedDayName & " " & Mid$(txtXtext.Caption, DayPos, 2) & " " & ArrMonth(Language, Val(Mid$(Text, MonthPos, 2)) - 1) & " " & Right$(txtXtext, 4)
    Else
        SelectedFullDate = ArrMonth(Language, Val(Mid$(Text, MonthPos, 2)) - 1) & " " & SelectedDayName & " " & Mid$(txtXtext.Caption, DayPos, 2) & " " & Right$(txtXtext, 4)
    End If
End Property
Public Property Let SelectedFullDate(ByVal New_SelectedFullDate As String)
    PropertyChanged "SelectedFullDate"
End Property

' Return the Text.
Public Property Get StartDate() As String
    StartDate = m_StartDate
End Property

' Set the Text.
Public Property Let StartDate(ByVal New_StartDate As String)
    If IsDate(New_StartDate) = False Then GoTo ErrHandler
    ' Make sure Month is ok for dd/mm/yyyy or mm/dd/yyyy
    If Mid$(New_StartDate, MonthPos, 2) > 12 Then GoTo ErrHandler
    m_StartDate = New_StartDate
    PropertyChanged "StartDate"
    Exit Property
ErrHandler:
    MsgBox New_StartDate & vbCrLf & " Is an invalid date", vbOKOnly + vbExclamation, "Error"
    New_StartDate = ""
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CurrentDayBOrderColor() As OLE_COLOR
    CurrentDayBOrderColor = m_CurrentDayBOrderColor
End Property
Public Property Let CurrentDayBOrderColor(ByVal New_CurrentDayBOrderColor As OLE_COLOR)
    m_CurrentDayBOrderColor = New_CurrentDayBOrderColor
    PropertyChanged "CurrentDayBOrderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CalendarBorderColor() As OLE_COLOR
    CalendarBorderColor = m_CalendarBorderColor
End Property
Public Property Let CalendarBorderColor(ByVal New_CalendarBorderColor As OLE_COLOR)
    m_CalendarBorderColor = New_CalendarBorderColor
    PropertyChanged "CalendarBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DaysForeColor() As OLE_COLOR
    DaysForeColor = m_DaysForeColor
End Property
Public Property Let DaysForeColor(ByVal New_DaysForeColor As OLE_COLOR)
    m_DaysForeColor = New_DaysForeColor
    PropertyChanged "DaysForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DaysBackColor() As OLE_COLOR
    DaysBackColor = m_DaysBackColor
End Property
Public Property Let DaysBackColor(ByVal New_DaysBackColor As OLE_COLOR)
    m_DaysBackColor = New_DaysBackColor
    PropertyChanged "DaysBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DaysBorderColor() As OLE_COLOR
    DaysBorderColor = m_DaysBorderColor
End Property
Public Property Let DaysBorderColor(ByVal New_DaysBorderColor As OLE_COLOR)
    m_DaysBorderColor = New_DaysBorderColor
    PropertyChanged "DaysBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DayNamesForeColor() As OLE_COLOR
    DayNamesForeColor = m_DayNamesForeColor
End Property
Public Property Let DayNamesForeColor(ByVal New_DayNamesForeColor As OLE_COLOR)
    m_DayNamesForeColor = New_DayNamesForeColor
    PropertyChanged "DayNamesForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DayNamesBackColor() As OLE_COLOR
    DayNamesBackColor = m_DayNamesBackColor
End Property
Public Property Let DayNamesBackColor(ByVal New_DayNamesBackColor As OLE_COLOR)
    m_DayNamesBackColor = New_DayNamesBackColor
    PropertyChanged "DayNamesBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DaySelectedBackColor() As OLE_COLOR
    DaySelectedBackColor = m_DaySelectedBackColor
End Property
Public Property Let DaySelectedBackColor(ByVal New_DaySelectedBackColor As OLE_COLOR)
    m_DaySelectedBackColor = New_DaySelectedBackColor
    PropertyChanged "DaySelectedBackColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DaySelectedForeColor() As OLE_COLOR
    DaySelectedForeColor = m_DaySelectedForeColor
End Property
Public Property Let DaySelectedForeColor(ByVal New_DaySelectedForeColor As OLE_COLOR)
    m_DaySelectedForeColor = New_DaySelectedForeColor
    PropertyChanged "DaySelectedForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get Language() As Language
    Language = m_Language
End Property
Public Property Let Language(ByVal New_Language As Language)
    m_Language = New_Language
    PropertyChanged "Language"
    SetLanguage
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DayNamesBorderColor() As OLE_COLOR
    DayNamesBorderColor = m_DayNamesBorderColor
End Property
Public Property Let DayNamesBorderColor(ByVal New_DayNamesBorderColor As OLE_COLOR)
    m_DayNamesBorderColor = New_DayNamesBorderColor
    PropertyChanged "DayNamesBorderColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnButtonBackColor() As OLE_COLOR
    DropDnButtonBackColor = m_DropDnButtonBackColor
End Property
Public Property Let DropDnButtonBackColor(ByVal New_DropDnButtonBackColor As OLE_COLOR)
    m_DropDnButtonBackColor = New_DropDnButtonBackColor
    PropertyChanged "DropDnButtonBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnButtonBorderColor() As OLE_COLOR
    DropDnButtonBorderColor = m_DropDnButtonBorderColor
End Property
Public Property Let DropDnButtonBorderColor(ByVal New_DropDnButtonBorderColor As OLE_COLOR)
    m_DropDnButtonBorderColor = New_DropDnButtonBorderColor
    PropertyChanged "DropDnButtonBorderColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CalendarBackColor() As OLE_COLOR
    CalendarBackColor = m_CalendarBackColor
End Property
Public Property Let CalendarBackColor(ByVal New_CalendarBackColor As OLE_COLOR)
    m_CalendarBackColor = New_CalendarBackColor
    PropertyChanged "CalendarBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnForeColor() As OLE_COLOR
    DropDnForeColor = m_DropDnForeColor
End Property
Public Property Let DropDnForeColor(ByVal New_DropDnForeColor As OLE_COLOR)
    m_DropDnForeColor = New_DropDnForeColor
    PropertyChanged "DropDnForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnBackColor() As OLE_COLOR
    DropDnBackColor = m_DropDnBackColor
End Property
Public Property Let DropDnBackColor(ByVal New_DropDnBackColor As OLE_COLOR)
    m_DropDnBackColor = New_DropDnBackColor
    PropertyChanged "DropDnBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnButtonForeColor() As OLE_COLOR
    DropDnButtonForeColor = m_DropDnButtonForeColor
End Property
Public Property Let DropDnButtonForeColor(ByVal New_DropDnButtonForeColor As OLE_COLOR)
    m_DropDnButtonForeColor = New_DropDnButtonForeColor
    PropertyChanged "DropDnButtonForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CalendarForeColor() As OLE_COLOR
    CalendarForeColor = m_CalendarForeColor
End Property
Public Property Let CalendarForeColor(ByVal New_CalendarForeColor As OLE_COLOR)
    m_CalendarForeColor = New_CalendarForeColor
    PropertyChanged "CalendarForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnBorderColor() As OLE_COLOR
    DropDnBorderColor = m_DropDnBorderColor
End Property

Public Property Let DropDnBorderColor(ByVal New_DropDnBorderColor As OLE_COLOR)
    m_DropDnBorderColor = New_DropDnBorderColor
    PropertyChanged "DropDnBorderColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnMonthForeColor() As OLE_COLOR
    DropDnMonthForeColor = m_DropDnMonthForeColor
End Property
Public Property Let DropDnMonthForeColor(ByVal New_DropDnMonthForeColor As OLE_COLOR)
    m_DropDnMonthForeColor = New_DropDnMonthForeColor
    PropertyChanged "DropDnMonthForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnMonthBackColor() As OLE_COLOR
    DropDnMonthBackColor = m_DropDnMonthBackColor
End Property
Public Property Let DropDnMonthBackColor(ByVal New_DropDnMonthBackColor As OLE_COLOR)
    m_DropDnMonthBackColor = New_DropDnMonthBackColor
    PropertyChanged "DropDnMonthBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnMonthBorderColor() As OLE_COLOR
    DropDnMonthBorderColor = m_DropDnMonthBorderColor
End Property
Public Property Let DropDnMonthBorderColor(ByVal New_DropDnMonthBorderColor As OLE_COLOR)
    m_DropDnMonthBorderColor = New_DropDnMonthBorderColor
    PropertyChanged "DropDnMonthBorderColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnMonthButtonBackColor() As OLE_COLOR
    DropDnMonthButtonBackColor = m_DropDnMonthButtonBackColor
End Property
Public Property Let DropDnMonthButtonBackColor(ByVal New_DropDnMonthButtonBackColor As OLE_COLOR)
    m_DropDnMonthButtonBackColor = New_DropDnMonthButtonBackColor
    PropertyChanged "DropDnMonthButtonBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnMonthButtonForeColor() As OLE_COLOR
    DropDnMonthButtonForeColor = m_DropDnMonthButtonForeColor
End Property
Public Property Let DropDnMonthButtonForeColor(ByVal New_DropDnMonthButtonForeColor As OLE_COLOR)
    m_DropDnMonthButtonForeColor = New_DropDnMonthButtonForeColor
    PropertyChanged "DropDnMonthButtonForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropDnMonthButtonBorderColor() As OLE_COLOR
    DropDnMonthButtonBorderColor = m_DropDnMonthButtonBorderColor
End Property
Public Property Let DropDnMonthButtonBorderColor(ByVal New_DropDnMonthButtonBorderColor As OLE_COLOR)
    m_DropDnMonthButtonBorderColor = New_DropDnMonthButtonBorderColor
    PropertyChanged "DropDnMonthButtonBorderColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SpinYearForeColor() As OLE_COLOR
    SpinYearForeColor = m_SpinYearForeColor
End Property
Public Property Let SpinYearForeColor(ByVal New_SpinYearForeColor As OLE_COLOR)
    m_SpinYearForeColor = New_SpinYearForeColor
    PropertyChanged "SpinYearForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SpinYearBackColor() As OLE_COLOR
    SpinYearBackColor = m_SpinYearBackColor
End Property
Public Property Let SpinYearBackColor(ByVal New_SpinYearBackColor As OLE_COLOR)
    m_SpinYearBackColor = New_SpinYearBackColor
    PropertyChanged "SpinYearBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SpinYearBorderColor() As OLE_COLOR
    SpinYearBorderColor = m_SpinYearBorderColor
End Property
Public Property Let SpinYearBorderColor(ByVal New_SpinYearBorderColor As OLE_COLOR)
    m_SpinYearBorderColor = New_SpinYearBorderColor
    PropertyChanged "SpinYearBorderColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SpinYearButtonBackColor() As OLE_COLOR
    SpinYearButtonBackColor = m_SpinYearButtonBackColor
End Property
Public Property Let SpinYearButtonBackColor(ByVal New_SpinYearButtonBackColor As OLE_COLOR)
    m_SpinYearButtonBackColor = New_SpinYearButtonBackColor
    PropertyChanged "SpinYearButtonBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SpinYearButtonForeColor() As OLE_COLOR
    SpinYearButtonForeColor = m_SpinYearButtonForeColor
End Property
Public Property Let SpinYearButtonForeColor(ByVal New_SpinYearButtonForeColor As OLE_COLOR)
    m_SpinYearButtonForeColor = New_SpinYearButtonForeColor
    PropertyChanged "SpinYearButtonForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SpinYearButtonBorderColor() As OLE_COLOR
    SpinYearButtonBorderColor = m_SpinYearButtonBorderColor
End Property
Public Property Let SpinYearButtonBorderColor(ByVal New_SpinYearButtonBorderColor As OLE_COLOR)
    m_SpinYearButtonBorderColor = New_SpinYearButtonBorderColor
    PropertyChanged "SpinYearButtonBorderColor"
    SetCalendarColors
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MonthListBackColor() As OLE_COLOR
    MonthListBackColor = m_MonthListBackColor
End Property
Public Property Let MonthListBackColor(ByVal New_MonthListBackColor As OLE_COLOR)
    m_MonthListBackColor = New_MonthListBackColor
    PropertyChanged "MonthListBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MonthListForeColor() As OLE_COLOR
    MonthListForeColor = m_MonthListForeColor
End Property
Public Property Let MonthListForeColor(ByVal New_MonthListForeColor As OLE_COLOR)
    m_MonthListForeColor = New_MonthListForeColor
    PropertyChanged "MonthListForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MonthListSelectedBackColor() As OLE_COLOR
    MonthListSelectedBackColor = m_MonthListSelectedBackColor
End Property
Public Property Let MonthListSelectedBackColor(ByVal New_MonthListSelectedBackColor As OLE_COLOR)
    m_MonthListSelectedBackColor = New_MonthListSelectedBackColor
    PropertyChanged "MonthListSelectedBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MonthListSelectedForeColor() As OLE_COLOR
    MonthListSelectedForeColor = m_MonthListSelectedForeColor
End Property
Public Property Let MonthListSelectedForeColor(ByVal New_MonthListSelectedForeColor As OLE_COLOR)
    m_MonthListSelectedForeColor = New_MonthListSelectedForeColor
    PropertyChanged "MonthListSelectedForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MonthListBorderColor() As OLE_COLOR
    MonthListBorderColor = m_MonthListBorderColor
End Property
Public Property Let MonthListBorderColor(ByVal New_MonthListBorderColor As OLE_COLOR)
    m_MonthListBorderColor = New_MonthListBorderColor
    PropertyChanged "MonthListBorderColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Style() As Styles
    Style = m_Style
End Property
Public Property Let Style(ByVal New_Style As Styles)
    m_Style = New_Style
    PropertyChanged "Style"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl.Enabled = m_Enabled
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DisabledBackColor() As OLE_COLOR
    DisabledBackColor = m_DisabledBackColor
End Property
Public Property Let DisabledBackColor(ByVal New_DisabledBackColor As OLE_COLOR)
    m_DisabledBackColor = New_DisabledBackColor
    PropertyChanged "DisabledBackColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DisabledForeColor() As OLE_COLOR
    DisabledForeColor = m_DisabledForeColor
End Property
Public Property Let DisabledForeColor(ByVal New_DisabledForeColor As OLE_COLOR)
    m_DisabledForeColor = New_DisabledForeColor
    PropertyChanged "DisabledForeColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DisabledBorderColor() As OLE_COLOR
    DisabledBorderColor = m_DisabledBorderColor
End Property
Public Property Let DisabledBorderColor(ByVal New_DisabledBorderColor As OLE_COLOR)
    m_DisabledBorderColor = New_DisabledBorderColor
    PropertyChanged "DisabledBorderColor"
    SetCalendarColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get DateFormat() As DateFormats
    DateFormat = m_DateFormat
End Property

Public Property Let DateFormat(ByVal New_DateFormat As DateFormats)
    m_DateFormat = New_DateFormat
    PropertyChanged "DateFormat"
    If DateFormat = 0 Then
        DayPos = 1
        MonthPos = 4
    Else
        DayPos = 4
        MonthPos = 1
    End If
End Property

Public Property Get FirstWeekDay() As Weekdays
    FirstWeekDay = m_FirstWeekDay
End Property
Public Property Let FirstWeekDay(ByVal New_FirstWeekDay As Weekdays)
    m_FirstWeekDay = New_FirstWeekDay
    PropertyChanged "FirstWeekDay"
    SetCalendarColors
    CalculateCalendar
End Property

