VERSION 5.00
Begin VB.Form frmDatePick 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin Project1.DMdatepicker DMdatepicker7 
      Height          =   375
      Left            =   11640
      TabIndex        =   8
      Top             =   4440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      DropDnBackColor =   8438015
      Enabled         =   0   'False
   End
   Begin Project1.DMdatepicker DMdatepicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      CalendarBackColor=   16761024
      CalendarForeColor=   15362350
      DropDnMonthButtonForeColor=   16777215
      SpinYearButtonForeColor=   16777215
      StartDate       =   "01/03/2010"
      DateFormat      =   "1"
   End
   Begin Project1.DMdatepicker DMdatepicker4 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   840
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      DaysForeColor   =   0
      DaysBackColor   =   192
      DaysBorderColor =   0
      DayNamesForeColor=   0
      DayNamesBackColor=   192
      DaySelectedBackColor=   0
      DaySelectedForeColor=   192
      DayNamesBorderColor=   0
      CurrentDayBOrderColor=   5725166
      CalendarBorderColor=   192
      CalendarBackColor=   0
      CalendarForeColor=   192
      DropDnForeColor =   8421631
      DropDnBackColor =   0
      DropDnButtonBackColor=   192
      DropDnButtonForeColor=   0
      DropDnButtonBorderColor=   0
      DropDnBorderColor=   192
      DropDnMonthForeColor=   8421631
      DropDnMonthBackColor=   0
      DropDnMonthBorderColor=   0
      DropDnMonthButtonBackColor=   192
      DropDnMonthButtonForeColor=   0
      DropDnMonthButtonBorderColor=   0
      MonthListBackColor=   192
      MonthListForeColor=   0
      MonthListSelectedBackColor=   0
      MonthListSelectedForeColor=   8421631
      MonthListBorderColor=   0
      SpinYearForeColor=   8421631
      SpinYearBackColor=   0
      SpinYearBorderColor=   192
      SpinYearButtonBackColor=   192
      SpinYearButtonForeColor=   0
      SpinYearButtonBorderColor=   192
      Language        =   1
      StartDate       =   "01/03/2010"
      FirstWeekDay    =   2
   End
   Begin Project1.DMdatepicker DMdatepicker3 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      DaysForeColor   =   15396349
      DaysBackColor   =   742118
      DaysBorderColor =   3365105
      DayNamesForeColor=   15396349
      DayNamesBackColor=   1003510
      DaySelectedBackColor=   5221999
      DaySelectedForeColor=   32768
      DayNamesBorderColor=   475089
      CurrentDayBOrderColor=   3569228
      CalendarBorderColor=   32768
      CalendarBackColor=   12648384
      CalendarForeColor=   475089
      DropDnForeColor =   475089
      DropDnBackColor =   9155832
      DropDnButtonBackColor=   5221999
      DropDnButtonForeColor=   14479330
      DropDnButtonBorderColor=   3569228
      DropDnBorderColor=   612831
      DropDnMonthForeColor=   16777215
      DropDnMonthBackColor=   9155832
      DropDnMonthBorderColor=   3569228
      DropDnMonthButtonBackColor=   998378
      DropDnMonthButtonForeColor=   12640511
      DropDnMonthButtonBorderColor=   16512
      MonthListBackColor=   9155832
      MonthListForeColor=   612831
      MonthListSelectedBackColor=   612831
      MonthListSelectedForeColor=   9155832
      MonthListBorderColor=   16512
      SpinYearForeColor=   16777215
      SpinYearBackColor=   9155832
      SpinYearBorderColor=   16512
      SpinYearButtonBackColor=   998378
      SpinYearButtonForeColor=   12640511
      SpinYearButtonBorderColor=   16512
      Language        =   2
      Style           =   0
      StartDate       =   "01/03/2010"
      FirstWeekDay    =   2
   End
   Begin Project1.DMdatepicker DMdatepicker5 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   4440
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      DaysForeColor   =   16761087
      DaysBackColor   =   8388736
      DaysBorderColor =   12583104
      DayNamesForeColor=   16761087
      DayNamesBackColor=   12583104
      DaySelectedBackColor=   16761087
      DaySelectedForeColor=   8388736
      DayNamesBorderColor=   8388736
      CurrentDayBOrderColor=   8388736
      CalendarBorderColor=   8388736
      CalendarBackColor=   16761087
      CalendarForeColor=   12583104
      DropDnForeColor =   16711935
      DropDnBackColor =   16761087
      DropDnButtonBackColor=   12583104
      DropDnButtonForeColor=   16761087
      DropDnButtonBorderColor=   8388736
      DropDnBorderColor=   12583104
      DropDnMonthForeColor=   16711935
      DropDnMonthBackColor=   16761087
      DropDnMonthBorderColor=   12583104
      DropDnMonthButtonBackColor=   12583104
      DropDnMonthButtonForeColor=   16761087
      DropDnMonthButtonBorderColor=   8388736
      MonthListBackColor=   16711935
      MonthListForeColor=   8388736
      MonthListSelectedBackColor=   8388736
      MonthListSelectedForeColor=   16761087
      MonthListBorderColor=   12583104
      SpinYearForeColor=   16711935
      SpinYearBackColor=   16761087
      SpinYearBorderColor=   12583104
      SpinYearButtonBackColor=   12583104
      SpinYearButtonForeColor=   16761087
      SpinYearButtonBorderColor=   8388736
      Language        =   3
      StartDate       =   "01/03/2010"
      FirstWeekDay    =   2
   End
   Begin Project1.DMdatepicker DMdatepicker2 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   4440
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      DaysForeColor   =   12640511
      DaysBackColor   =   4210816
      DaysBorderColor =   4210752
      DayNamesForeColor=   12632256
      DayNamesBackColor=   16384
      DaySelectedBackColor=   0
      DaySelectedForeColor=   14737632
      DayNamesBorderColor=   4210688
      CurrentDayBOrderColor=   255
      CalendarBorderColor=   0
      CalendarBackColor=   4210752
      CalendarForeColor=   33023
      DropDnForeColor =   64
      DropDnBackColor =   4210752
      DropDnButtonBackColor=   16576
      DropDnButtonForeColor=   12640511
      DropDnButtonBorderColor=   16512
      DropDnBorderColor=   16576
      DropDnMonthForeColor=   64
      DropDnMonthBackColor=   4210752
      DropDnMonthBorderColor=   16576
      DropDnMonthButtonBackColor=   16576
      DropDnMonthButtonForeColor=   12640511
      DropDnMonthButtonBorderColor=   16512
      MonthListBackColor=   16384
      MonthListForeColor=   64
      MonthListSelectedBackColor=   32768
      MonthListSelectedForeColor=   14737632
      MonthListBorderColor=   16384
      SpinYearForeColor=   64
      SpinYearBackColor=   4210752
      SpinYearBorderColor=   16576
      SpinYearButtonBackColor=   16576
      SpinYearButtonForeColor=   12640511
      SpinYearButtonBorderColor=   16512
      Language        =   4
      StartDate       =   "01/03/2010"
      FirstWeekDay    =   2
   End
   Begin Project1.DMdatepicker DMdatepicker6 
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   4440
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      DaysForeColor   =   192
      DaysBackColor   =   0
      DaysBorderColor =   192
      DayNamesForeColor=   192
      DayNamesBackColor=   0
      DaySelectedBackColor=   192
      DaySelectedForeColor=   0
      DayNamesBorderColor=   192
      CurrentDayBOrderColor=   0
      CalendarBorderColor=   0
      CalendarBackColor=   192
      CalendarForeColor=   0
      DropDnForeColor =   0
      DropDnBackColor =   192
      DropDnButtonBackColor=   0
      DropDnButtonForeColor=   255
      DropDnButtonBorderColor=   0
      DropDnBorderColor=   0
      DropDnMonthForeColor=   0
      DropDnMonthBackColor=   192
      DropDnMonthBorderColor=   192
      DropDnMonthButtonBackColor=   0
      DropDnMonthButtonForeColor=   192
      DropDnMonthButtonBorderColor=   192
      MonthListBackColor=   0
      MonthListForeColor=   192
      MonthListSelectedBackColor=   192
      MonthListSelectedForeColor=   0
      MonthListBorderColor=   192
      SpinYearForeColor=   0
      SpinYearBackColor=   192
      SpinYearBorderColor=   0
      SpinYearButtonBackColor=   0
      SpinYearButtonForeColor=   192
      SpinYearButtonBorderColor=   0
      Language        =   5
      StartDate       =   "01/03/2010"
      FirstWeekDay    =   2
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "3D Dutch"
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "firstweekday = Sunday"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DateFormat = mm/dd/yyyy"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "3D English"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "3D Spanish"
      Height          =   255
      Left            =   8760
      TabIndex        =   14
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "3D German"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "3D Dutch"
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "3D Italian"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Disabled"
      Height          =   255
      Left            =   11640
      TabIndex        =   10
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Flat French"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   9480
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "frmDatePick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Programmer:        Donckers Frank
'                    DarkManSoft@Gmail.com
'
' Description:       TestForm User Control DatePicker



Private Sub DMdatepicker1_Change()
    Label1 = "Date: " & DMdatepicker1.Text & vbCrLf & "Dateshort: " & DMdatepicker1.SelectedDateShort & vbCrLf
    Label1 = Label1 & "Day: " & DMdatepicker1.SelectedDay & vbCrLf & "DayName: " & DMdatepicker1.SelectedDayName
    Label1 = Label1 & vbCrLf & "Month: " & DMdatepicker1.SelectedMonth & vbCrLf & "Monthname: " & DMdatepicker1.SelectedMonthName
    Label1 = Label1 & vbCrLf & "Year: " & DMdatepicker1.SelectedYear & vbCrLf & "Yearshort: " & DMdatepicker1.SelectedYearShort
    Label1 = Label1 & vbCrLf & "FullDate: " & DMdatepicker1.SelectedFullDate
End Sub


Private Sub DMdatepicker4_Change()
    Label2 = "Date: " & DMdatepicker4.Text & vbCrLf & "Dateshort: " & DMdatepicker4.SelectedDateShort & vbCrLf
    Label2 = Label2 & "Day: " & DMdatepicker4.SelectedDay & vbCrLf & "DayName: " & DMdatepicker4.SelectedDayName
    Label2 = Label2 & vbCrLf & "Month: " & DMdatepicker4.SelectedMonth & vbCrLf & "Monthname: " & DMdatepicker4.SelectedMonthName
    Label2 = Label2 & vbCrLf & "Year: " & DMdatepicker4.SelectedYear & vbCrLf & "Yearshort: " & DMdatepicker4.SelectedYearShort
    Label2 = Label2 & vbCrLf & "FullDate: " & DMdatepicker4.SelectedFullDate

End Sub

