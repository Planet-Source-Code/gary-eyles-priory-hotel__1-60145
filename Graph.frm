VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "The Priory Hotel Booking Systen"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14325
   Icon            =   "Graph.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   658
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Selections"
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9720
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Graph.frx":23D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Graph.frx":47B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Graph.frx":6B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Graph.frx":8F78
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox List1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   8040
      Width           =   2595
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Undo"
      Height          =   255
      Left            =   7080
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data cCust 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer Details"
      Top             =   7920
      Visible         =   0   'False
      Width           =   3180
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Graph.frx":B35A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Graph.frx":D73C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Graph.frx":FB1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Graph.frx":11F00
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox fDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Text            =   "23/10/03"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox sDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "23/9/03"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Data tHolidays 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Holidays"
      Top             =   8280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox wDays 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2880
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   9
      Top             =   7560
      Width           =   4695
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   9570
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14499
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "00:33"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   2
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox tNames 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Left            =   3240
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.PictureBox tnums 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4680
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   5
      Top             =   1920
      Width           =   8295
   End
   Begin VB.PictureBox tcal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   4680
      ScaleHeight     =   441
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   633
      TabIndex        =   0
      Top             =   2760
      Width           =   9495
   End
   Begin VB.Data tDates 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Dates"
      Top             =   8640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   780
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1376
      ButtonWidth     =   2037
      ButtonHeight    =   1376
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Customer"
            Key             =   "New Customer"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Old Customer"
            Key             =   "OldCustomer"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sidebar"
            Key             =   "sidebar"
            ImageIndex      =   4
            Value           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Key             =   "Settings"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Day"
            Key             =   "PrintBookings"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Calendar"
            Key             =   "PrintCalendar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Holidays"
            Key             =   "holidays"
         EndProperty
      EndProperty
   End
   Begin The_Priory_Hotel.uSideBar uSideBar 
      Height          =   7575
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   3135
      _extentx        =   5530
      _extenty        =   13361
   End
   Begin VB.Label Label5 
      Caption         =   "Day +/-"
      Height          =   255
      Left            =   10080
      TabIndex        =   22
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Month +/-"
      Height          =   255
      Left            =   10080
      TabIndex        =   21
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Num Days +/-"
      Height          =   255
      Left            =   10080
      TabIndex        =   20
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8400
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Selections="
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   840
      Width           =   1335
   End
   Begin MSForms.SpinButton SpinButton4 
      Height          =   255
      Left            =   11280
      TabIndex        =   10
      Top             =   1320
      Width           =   615
      Size            =   "1085;450"
   End
   Begin MSForms.SpinButton SpinButton2 
      Height          =   255
      Left            =   11280
      TabIndex        =   4
      Top             =   1080
      Width           =   615
      Size            =   "1085;450"
   End
   Begin MSForms.SpinButton SpinButton3 
      Height          =   255
      Left            =   11280
      TabIndex        =   7
      Top             =   1560
      Width           =   615
      ForeColor       =   8421631
      Size            =   "1085;450"
      Orientation     =   1
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
      Size            =   "1085;661"
   End
   Begin VB.Label ExtraInfo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   960
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
   End
   Begin VB.Menu selection 
      Caption         =   "&Selection"
      Enabled         =   0   'False
      Begin VB.Menu undoSel 
         Caption         =   "&Undo"
      End
      Begin VB.Menu ClearSel 
         Caption         =   "&Clear Selections"
      End
   End
   Begin VB.Menu BookingSel 
      Caption         =   "&Booking"
      Enabled         =   0   'False
      Begin VB.Menu moveto 
         Caption         =   "Move to"
         Begin VB.Menu moveroom 
            Caption         =   "1 Snowhill"
            Index           =   0
         End
         Begin VB.Menu moveroom 
            Caption         =   "2 Square View"
            Index           =   1
         End
         Begin VB.Menu moveroom 
            Caption         =   "3 Ladyewell"
            Index           =   2
         End
         Begin VB.Menu moveroom 
            Caption         =   "4 Littledale"
            Index           =   3
         End
         Begin VB.Menu moveroom 
            Caption         =   "5 Nickynook"
            Index           =   4
         End
         Begin VB.Menu moveroom 
            Caption         =   "6 Wyresdale"
            Index           =   5
         End
         Begin VB.Menu moveroom 
            Caption         =   "7 /3a Arbor"
            Index           =   6
         End
      End
      Begin VB.Menu DelBooking 
         Caption         =   "&Delete Booking"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu Aboutbox 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim usDate As Date
Dim ufDate As Date
Dim numdays As Long
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFacename As String * 33
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Type selDate
    Uni As Long
    stdate As Date
    fnDate As Date
    roomNo As Long
    roomNoName As String
    custid As Long
    moveable As Boolean
    noPeople As Long
    CustName As String
    Address As String
    Quote As String
    'bookmark As Recordset.bookmark
End Type

Private Type gColors
    BookedColors As Long
    SelectBooked As Long
    SameIDSelectBooked As Long
    Column1 As Long
    Column2 As Long
    ColumnText1 As Long
    ColumnText2 As Long
    Month1 As Long
    Month2 As Long
    Holiday As Long
    NonMovable As Long
End Type

Dim gCols As gColors
Dim stSel As Boolean
Dim custSel As selDate
Dim C As ExToolTip
Dim MMove As Boolean
Dim calSel() As selDate
Dim calSelNum As Long

Private Type booked
    custid As Long
    startdate As Date
    enddate As Date
    numdays As Long
    special As Long
    lastone As Boolean
    Uni As Long
    Room As String
    moveable As Boolean
    cname As String
    nump As Integer
End Type

Private Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
         Const vbHiMetric As Integer = 8
         Dim PicRatio As Double
         Dim PrnWidth As Double
         Dim PrnHeight As Double
         Dim PrnRatio As Double
         Dim PrnPicWidth As Double
         Dim PrnPicHeight As Double

         ' Determine if picture should be printed in landscape or portrait
         ' and set the orientation
         If Pic.Height >= Pic.Width Then
            Prn.Orientation = vbPRORPortrait   ' Taller than wide
         Else
            Prn.Orientation = vbPRORLandscape  ' Wider than tall
         End If

         ' Calculate device independent Width to Height ratio for picture
         PicRatio = Pic.Width / Pic.Height

         ' Calculate the dimentions of the printable area in HiMetric
         PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
         PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
         ' Calculate device independent Width to Height ratio for printer
         PrnRatio = PrnWidth / PrnHeight

         ' Scale the output to the printable area
         If PicRatio >= PrnRatio Then
            ' Scale picture to fit full width of printable area
            PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
            PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
               Prn.ScaleMode)
         Else
            ' Scale picture to fit full height of printable area
            PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
            PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
               Prn.ScaleMode)
         End If

         ' Print the picture using the PaintPicture method
         Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub

Private Function ChangeDate(stdate As Date, endate As Date, theroom As String, tuni As Long) As booked
'Call tst

Dim cc As Long
Dim cc2 As Long
'For cc = 1 To tDates.Recordset.RecordCount
tDates.Recordset.MoveFirst
Do
    If tDates.Recordset.Fields("Room") = theroom Then
        If tuni <> tDates.Recordset.Fields("Uni") Then
            If tDates.Recordset.Fields("start date") > stdate And tDates.Recordset.Fields("end date") < endate Then
                        ChangeDate.Uni = tDates.Recordset.Fields("Uni")
                        ChangeDate.startdate = tDates.Recordset.Fields("start date")
                        ChangeDate.enddate = tDates.Recordset.Fields("end date")
                        ChangeDate.Room = tDates.Recordset.Fields("Room")
                        ChangeDate.moveable = tDates.Recordset.Fields("Moveable")
            End If
            If stdate >= tDates.Recordset.Fields("start date") And stdate <= tDates.Recordset.Fields("end date") Then
                        ChangeDate.Uni = tDates.Recordset.Fields("Uni")
                        ChangeDate.startdate = tDates.Recordset.Fields("start date")
                        ChangeDate.enddate = tDates.Recordset.Fields("end date")
                        ChangeDate.Room = tDates.Recordset.Fields("Room")
                        ChangeDate.moveable = tDates.Recordset.Fields("Moveable")
                        Exit Function
            End If
            If endate >= tDates.Recordset.Fields("start date") And endate <= tDates.Recordset.Fields("end date") Then
                        ChangeDate.Uni = tDates.Recordset.Fields("Uni")
                        ChangeDate.startdate = tDates.Recordset.Fields("start date")
                        ChangeDate.enddate = tDates.Recordset.Fields("end date")
                        ChangeDate.Room = tDates.Recordset.Fields("Room")
                        ChangeDate.moveable = tDates.Recordset.Fields("Moveable")
                        Exit Function
            End If
        End If
    End If
'Next
tDates.Recordset.MoveNext
Loop Until tDates.Recordset.EOF
Debug.Print
End Function

Private Sub ApplyDemoValues() '//--Default Values used in Demo
               
        '//-- Remember each time a new Tooltip is created if the
        '     user doesn't specify any parameters values like this ones,
        '     the default values are going to be the Extooltip default values.(See ExTooltip Class_Initialize).
        
        C.DelayTime = 5
        C.KillTime = 5
        C.BackColor = QBColor(14)
        C.TextColor = 0
        C.GradientColorStart = 0
        C.GradientColorEnd = 0
        C.BackStyle = 1
        C.Font.Name = "Arial"
        'C.Shadow = IIf(CheckShadow.Value = 1, True, False)
        'C.ToolTipStyle = IIf(CheckBalloon.Value = 1, 1, 0)

End Sub

Sub Reset()
usDate = CDate(sDate)
ufDate = CDate(fDate)
numdays = ufDate - usDate + 1
Text1 = numdays

If uSideBar.Bar("Boards").State = BarState_Expanded Then
    uSideBar.Bar("Boards").Item("StartDate").Text = "Start Date: " & sDate
    uSideBar.Bar("Boards").Item("EndDate").Text = "End Date: " & fDate
    uSideBar.Bar("Boards").Item("NumDays").Text = "Num Days: " & Text1
End If

Dim qdftemp As QueryDef
Set qdftemp = tDates.Database.CreateQueryDef("", "PARAMETERS StDate DateTime, EnDate DateTime; SELECT Dates.Uni, Dates.ID, Dates.[Start Date], Dates.[End Date], Dates.Room, Dates.[No People], Dates.Moveable, Dates.[Date made], Dates.[Time made], Dates.[Quote] From Dates WHERE (((Dates.[Start Date])>[stDate]) AND ((Dates.[End Date])<[enDate]));")
qdftemp.Parameters!stdate = usDate - numdays
qdftemp.Parameters!endate = ufDate + numdays
tDates.Recordset.Requery qdftemp

Set qdftemp = tHolidays.Database.CreateQueryDef("", "PARAMETERS stDate DateTime, enDate DateTime; SELECT Holidays.ID, Holidays.Comments, Holidays.Yearly, Holidays.Birthday, Holidays.[Hol Date], IIf([yearly]=True,DateSerial(Year([stdate]),Month([hol date]),Day([hol date])),[hol date]) AS Expr3, IIf([yearly]=True,DateSerial(Year([endate]),Month([hol date]),Day([hol date])),[hol date]) AS Expr4 From Holidays WHERE (((IIf([yearly]=True,DateSerial(Year([stdate]),Month([hol date]),Day([hol date])),[hol date]))>=[stdate] And (IIf([yearly]=True,DateSerial(Year([stdate]),Month([hol date]),Day([hol date])),[hol date]))<=[endate])) OR (((IIf([yearly]=True,DateSerial(Year([endate]),Month([hol date]),Day([hol date])),[hol date]))>=[stdate] And (IIf([yearly]=True,DateSerial(Year([endate]),Month([hol date]),Day([hol date])),[hol date]))<=[endate]));")
qdftemp.Parameters!stdate = usDate
qdftemp.Parameters!endate = ufDate
tHolidays.Recordset.Requery qdftemp

'uSideBar.Bar("Holidays").RemoveALL
'uSideBar.Redraw
'Dim cc As Long

List1.Text = ""
If tHolidays.Recordset.RecordCount > 0 Then
Do
'    cc = cc + 1
'    If tHolidays.Recordset.Fields("Birthday") = True Then
'        uSideBar.Bar("Holidays").AddItem , tHolidays.Recordset.Fields("Hol Date") & " " & tHolidays.Recordset.Fields("Comments") & " " & Year(usDate) - Year(tHolidays.Recordset.Fields("Hol Date"))
'    Else
'        uSideBar.Bar("Holidays").AddItem , tHolidays.Recordset.Fields("Hol Date") & " " & tHolidays.Recordset.Fields("Comments")
'    End If

        If tHolidays.Recordset.Fields("Birthday") = True Then
            List1.Text = List1.Text & tHolidays.Recordset.Fields("Hol Date") & " " & tHolidays.Recordset.Fields("Comments") & " " & Year(usDate) - Year(tHolidays.Recordset.Fields("Hol Date")) & vbCrLf
        Else
            List1.Text = List1.Text & tHolidays.Recordset.Fields("Hol Date") & " " & tHolidays.Recordset.Fields("Comments") & vbCrLf
        End If

tHolidays.Recordset.MoveNext
Loop Until tHolidays.Recordset.EOF
End If

'tHolidays.Recordset.MoveFirst
'If frmHol.Visible = True Then
'    frmHol.List1.Clear
'    Do
'        If tHolidays.Recordset.Fields("Birthday") = True Then
'            frmHol.List1.AddItem tHolidays.Recordset.Fields("Hol Date") & " " & tHolidays.Recordset.Fields("Comments") & " " & Year(usDate) - Year(tHolidays.Recordset.Fields("Hol Date"))
'        Else
'            frmHol.List1.AddItem tHolidays.Recordset.Fields("Hol Date") & " " & tHolidays.Recordset.Fields("Comments")
'        End If
'    tHolidays.Recordset.MoveNext
'    Loop Until tHolidays.Recordset.EOF
'Else
'    Unload frmHol
'End If

tcal_Paint
End Sub

Private Sub SetDefaultColours()
gCols.BookedColors = QBColor(14)
gCols.SameIDSelectBooked = QBColor(12)
gCols.SelectBooked = QBColor(12)
gCols.NonMovable = QBColor(15)
gCols.Holiday = RGB(250, 0, 250)
gCols.Column2 = RGB(190, 190, 190)
gCols.Column1 = RGB(200, 200, 200)
gCols.ColumnText1 = RGB(230, 230, 230)
gCols.ColumnText2 = RGB(210, 210, 210)
End Sub

Private Sub Aboutbox_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub ClearSel_Click()
Command2_Click
End Sub

Private Sub Command2_Click()
ReDim calSel(0)
calSelNum = 1
Label2.Caption = 0
tcal_Paint
selection.Enabled = False
End Sub

Private Sub Command3_Click()
If calSelNum > 1 Then
    calSelNum = calSelNum - 1
    Label2.Caption = calSelNum - 1
    ReDim Preserve calSel(calSelNum - 1)
    tcal_Paint
End If
If calSelNum = 1 Then
    selection.Enabled = False
End If
End Sub

Private Sub DelBooking_Click()
If custSel.custid > 0 Then
    tDates.Recordset.MoveFirst
    Do
    tDates.Recordset.MoveNext
    Loop Until tDates.Recordset.Fields("Uni") = custSel.Uni
    'tDates.Recordset.bookmark = custSel.bookmark
    tDates.Recordset.Delete
        custSel.stdate = vbNull
        custSel.fnDate = vbNull
        custSel.roomNo = 0
        custSel.custid = 0
        custSel.Uni = 0
        'custSel.bookmark = 0
        ExtraInfo.Caption = ""
        BookingSel.Enabled = False
    Call Reset
    tcal_Paint
End If
End Sub

Private Sub Exit_Click()
Unload Me
Unload frmHol
End Sub

Private Sub ExtraInfo_Click()
Dim tmpstr As String

If custSel.Uni > 0 Then
    On Error Resume Next
    tDates.Recordset.MoveFirst
    tDates.Recordset.FindFirst "Uni=" & custSel.Uni
    tDates.Recordset.edit
    
    Debug.Print tDates.Recordset.Fields("Quote")
    If IsNull(tDates.Recordset.Fields("Quote")) Then
        tDates.Recordset.Fields("Quote").Value = ""
    End If
    tmpstr = InputBox("Enter new quote", , (tDates.Recordset.Fields("Quote")))
    tDates.Recordset.Fields("Quote") = IIf(tmpstr = "", tDates.Recordset.Fields("Quote"), tmpstr)
    tDates.Recordset.Update
    ExtraInfo.Caption = tDates.Recordset.Fields("Quote")
End If
End Sub

Private Sub Form_Activate()
Toolbar1.Refresh
Call Reset
tcal_Paint
uSideBar.Redraw
End Sub


Private Sub Form_Load()
ReDim calSel(0)
calSelNum = 1
sDate.Text = Date
fDate.Text = DateSerial(Year(Date), Month(Date) + 1, Day(Date))
tDates.DatabaseName = App.Path & "\Hotel.mdb"
tHolidays.DatabaseName = App.Path & "\Hotel.mdb"
cCust.DatabaseName = App.Path & "\Hotel.mdb"

Call SetDefaultColours
'Toolbar1.Buttons(1).Height = 10
'toolbar1.Buttons.Add

Set C = New ExToolTip

    Dim Tell As Long

    With uSideBar
        ' Manually stop all redraws
        .IgnoreRedraw = True
        .AddBar "Boards", "Dates:"
        '.Bar("Boards").AddItem "Label", "", ItemType_ControlPlaceHolder
        
        ' Add an control into the "boars"-bar
'        .Bar("Boards").AddItem("Text", "Text", ItemType_Object).Control = sDate
'        .Bar("Boards").AddItem("Text", "Text", ItemType_Object).Control = fDate
        .Bar("Boards").AddItem "StartDate", "Start date: " & sDate, ItemType_ControlPlaceHolder
        .Bar("Boards").AddItem "EndDate", "End date: " & fDate, ItemType_ControlPlaceHolder
        .Bar("Boards").AddItem "NumDays", "Num Days: " & Text1, ItemType_ControlPlaceHolder
        
        .AddBar "Details", "Customer Details:"
        .Bar("Details").AddItem "Title", "Title: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "Forename", "Forename: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "Surname", "Surname: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "Address1", "Address1: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "Address2", "Address2: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "Address3", "Address3: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "County", "County: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "Town", "Town: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "Country", "Country: ", ItemType_ControlPlaceHolder
        .Bar("Details").AddItem "Postcode", "Postcode: ", ItemType_ControlPlaceHolder
        
        .AddBar "Holidays", "Holidays: "
        .Bar("Holidays").AddItem("Text", "Text", ItemType_Object).Control = List1
               
        .AddBar "Extra", "Extra:"
        .Bar("Extra").AddItem "numStays", "No Stays: ", ItemType_ControlPlaceHolder
        
'        .AddBar "Holidays", "Boards:"
'        .Bar("Holidays").AddItem "Label", "Write something here:", ItemType_ControlPlaceHolder
'        .Bar("Holidays").AddItem("Text", "Text", ItemType_Object).Control = List1

        'For Tell = 1 To 2
        '    .Bar("Holidays").AddItem "HOLS" & Tell, "Element number " & Tell, ItemType_Link
        'Next
        
        .Bar("Holidays").MaxItems = 5
        
        ' Redraw it
        List1.BackColor = .Bar("Holidays").BackColor
        .IgnoreRedraw = False
        .Redraw
    End With

End Sub

Private Sub Form_Resize()
On Error Resume Next
Call Reset
If Form1.WindowState <> 1 Then
    tcal.Width = ScaleWidth - tcal.Left - 10
    tcal.Height = ScaleHeight - tcal.TOP - 10 - StatusBar1.Height - wDays.Height
    wDays.Left = tcal.Left
    wDays.TOP = tcal.Height + tcal.TOP
    SpinButton2.Left = ScaleWidth - SpinButton2.Width - 50
    SpinButton3.Left = ScaleWidth - SpinButton3.Width - 50
    SpinButton4.Left = ScaleWidth - SpinButton4.Width - 50
    uSideBar.Height = ScaleHeight - StatusBar1.Height
    'CoolBar1.Width = ScaleWidth
    Label3.Left = SpinButton2.Left - Label3.Width
    Label4.Left = SpinButton4.Left - Label4.Width
    Label5.Left = SpinButton3.Left - Label5.Width
    Toolbar1.Width = ScaleWidth
    tcal_Paint
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmHol
End Sub

Private Sub moveroom_Click(Index As Integer)
Dim canmove As booked
Dim tmpmove As booked
Dim rallFrom() As booked

Dim Roomfrom As String
Dim RoomTo As String
Dim canval As Long
Dim ccval As Long
Dim rs As Object
Dim DadDate As Boolean
Dim DoneChangeBack As Boolean
Dim croom As Long

Dim qdftemp As QueryDef
Set qdftemp = tDates.Database.CreateQueryDef("", "PARAMETERS StDate DateTime, EnDate DateTime; SELECT Dates.Uni, Dates.ID, Dates.[Start Date], Dates.[End Date], Dates.Room, Dates.[No People], Dates.Moveable, Dates.[Date made], Dates.[Time made], Dates.[Quote] From Dates WHERE (((Dates.[Start Date])>[stDate]) AND ((Dates.[End Date])<[enDate]));")
qdftemp.Parameters!stdate = usDate - numdays
qdftemp.Parameters!endate = ufDate + numdays
tDates.Recordset.Requery qdftemp

Set rs = tDates.Recordset.Clone
Dim tmpstr As String
Dim baddate As Boolean

Roomfrom = custSel.roomNoName
'Roomfrom = Room
'On Error GoTo tcancel
'RoomTo = Combo56
'RoomTo = "1 Snowhill"
RoomTo = moveroom(Index).Caption

changeback:
    
    If canval = 0 Then
        ReDim Preserve rallFrom(canval)
        canval = canval
        rallFrom(canval).custid = custSel.custid
        rallFrom(canval).enddate = custSel.fnDate
        rallFrom(canval).startdate = custSel.stdate
        rallFrom(canval).Room = custSel.roomNoName
        rallFrom(canval).Uni = custSel.Uni
        rallFrom(canval).moveable = custSel.moveable
    End If
    
Do
    For ccval = 0 To UBound(rallFrom)
        If (rallFrom(ccval).startdate < Date And rallFrom(ccval).enddate >= Date) Or rallFrom(ccval).moveable = False Then
            'MsgBox "THIS ONE " & ccval
            baddate = True
        End If
        
        canmove = ChangeDate(rallFrom(ccval).startdate, rallFrom(ccval).enddate, RoomTo, rallFrom(ccval).Uni)
        If canmove.Uni > 0 Then
            Exit For
        ElseIf canmove.Uni = 0 Then
            rs.MoveLast
            rs.FindFirst "[UNI] = " & Str(rallFrom(ccval).Uni)
            If Not rs.EOF Then tDates.Recordset.Bookmark = rs.Bookmark
            tDates.Recordset.edit
            tDates.Recordset.Fields("Room") = RoomTo
            tDates.Recordset.Update
        End If
    Next
         
    If canmove.Uni > 0 Then
        rs.MoveLast
        rs.FindFirst "[UNI] = " & Str(rallFrom(ccval).Uni)
        If Not rs.EOF Then tDates.Recordset.Bookmark = rs.Bookmark
        tDates.Recordset.edit
        tDates.Recordset.Fields("Room") = Left(Roomfrom, 1) & " TEMP"
        tDates.Recordset.Update
        Do
            If (canmove.startdate < Date And canmove.enddate >= Date) Or canmove.moveable = False Then
                'MsgBox "HERE"
                baddate = True
            End If
            
            tmpmove = ChangeDate(canmove.startdate, canmove.enddate, Roomfrom, canmove.Uni)
            If tmpmove.Uni > 0 Then
                canval = canval + 1
                ReDim Preserve rallFrom(canval)
                rallFrom(canval) = tmpmove
                rs.MoveLast
                rs.FindFirst "[UNI] = " & Str(tmpmove.Uni)
                If Not rs.EOF Then tDates.Recordset.Bookmark = rs.Bookmark
                tDates.Recordset.edit
                tDates.Recordset.Fields("Room") = Left(Roomfrom, 1) & "TEMP"
                tDates.Recordset.Update
            ElseIf tmpmove.Uni = 0 Then
                rs.MoveLast
                rs.FindFirst "[UNI] = " & Str(canmove.Uni)
                If Not rs.EOF Then tDates.Recordset.Bookmark = rs.Bookmark
                tDates.Recordset.edit
                tDates.Recordset.Fields("Room") = Roomfrom
                tDates.Recordset.Update
            End If
        Loop Until tmpmove.Uni = 0
    End If
Loop Until canmove.Uni = 0

If baddate = True And DoneChangeBack = False Then
    tmpstr = Roomfrom
    Roomfrom = RoomTo
    RoomTo = tmpstr
    Erase rallFrom
    canval = 0
    DoneChangeBack = True
    GoTo changeback
End If

custSel.roomNoName = RoomTo
custSel.roomNo = CInt(Mid(RoomTo, 1, 1))
'Call tst(IIf(Check67.Value = -1, True, False))
'Form_Resize
'Debug.Print "END " & baddate
Call Reset

Exit Sub
tcancel:

End Sub

Private Sub SpinButton1_SpinDown()
sDate = CDate(sDate) - 1
Call Reset
End Sub

Private Sub SpinButton1_SpinUp()
sDate = CDate(sDate) + 1
Call Reset
End Sub

Private Sub SpinButton2_SpinDown()
fDate = CDate(fDate) - 1
Call Reset
End Sub

Private Sub SpinButton2_SpinUp()
fDate = CDate(fDate) + 1
Call Reset
End Sub


Private Sub SpinButton3_SpinDown()
fDate = CDate(fDate) - 1
sDate = CDate(sDate) - 1
Call Reset
End Sub

Private Sub SpinButton3_SpinUp()
fDate = CDate(fDate) + 1
sDate = CDate(sDate) + 1
Call Reset
End Sub

Private Sub SpinButton4_SpinDown()
Dim tmpdate As Date
tmpdate = CDate(fDate)
fDate = DateSerial(Year(tmpdate), Month(tmpdate) - 1, Day(tmpdate))
tmpdate = CDate(sDate)
sDate = DateSerial(Year(tmpdate), Month(tmpdate) - 1, Day(tmpdate))
Call Reset
tcal_Paint
End Sub

Private Sub SpinButton4_SpinUp()
Dim tmpdate As Date
tmpdate = CDate(fDate)
fDate = DateSerial(Year(tmpdate), Month(tmpdate) + 1, Day(tmpdate))
tmpdate = CDate(sDate)
sDate = DateSerial(Year(tmpdate), Month(tmpdate) + 1, Day(tmpdate))
Call Reset
tcal_Paint
End Sub

Private Sub tcal_DblClick()
If custSel.custid <> 0 Then
    Load CustDetails
    CustDetails.Data1.Recordset.MoveFirst
    CustDetails.Data1.Recordset.FindFirst "ID=" & custSel.custid
    CustDetails.Show vbModal, Me
    tDates.Refresh
    Call Reset
    tcal_Paint
End If
End Sub

Private Sub tcal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xx As Long
Dim yy As Long
Dim xad As Long
Dim yad As Long
Dim wDate As Date
xad = tcal.ScaleWidth / numdays
yad = tcal.ScaleHeight / 7

Do
    xad = xad - 1
Loop Until xad * numdays < tcal.ScaleWidth
Do
    yad = yad - 1
Loop Until yad * 7 < tcal.ScaleHeight

ReDim Preserve calSel(calSelNum)
If Button = 1 Then
    yy = Y \ yad
    xx = X \ xad
    wDate = usDate + xx
    If wDate >= Date Then
        stSel = True
        calSel(calSelNum).stdate = wDate
        calSel(calSelNum).fnDate = wDate
        calSel(calSelNum).roomNo = yy + 1
    End If
    'Debug.Print "DOWN" & calSel(calSelNum).roomNo
    tcal_Paint
End If
End Sub

Private Sub tcal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xx As Long
Dim yy As Long
Dim xad As Long
Dim yad As Long
Dim wDate As Date
xad = tcal.ScaleWidth / numdays
yad = tcal.ScaleHeight / 7
On Error Resume Next
Do
    xad = xad - 1
Loop Until xad * numdays < tcal.ScaleWidth
Do
    yad = yad - 1
Loop Until yad * 7 < tcal.ScaleHeight

If Button = 1 Then
    yy = Y \ yad
    xx = X \ xad
    wDate = usDate + xx
    If calSel(calSelNum).fnDate <> wDate Then
        calSel(calSelNum).fnDate = wDate
        tcal_Paint
    End If
End If
End Sub

Private Sub tcal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xx As Long
Dim yy As Long
Dim xad As Long
Dim yad As Long
Dim fcol As Long
Dim wDate As Date
Dim monthcol As Boolean
Dim tmpRoom As String
Dim qdftemp As QueryDef
Dim cc As Long
Dim cc2 As Long
Dim croom As Long
Dim another As Boolean

xad = tcal.ScaleWidth / numdays
yad = tcal.ScaleHeight / 7

Do
    xad = xad - 1
Loop Until xad * numdays < tcal.ScaleWidth
Do
    yad = yad - 1
Loop Until yad * 7 < tcal.ScaleHeight

If Button = 1 Then
    yy = Y \ yad
    xx = X \ xad
    wDate = usDate + xx
    

ReDim Preserve calSel(calSelNum)
    calSel(calSelNum).fnDate = wDate
    If calSel(calSelNum).stdate > calSel(calSelNum).fnDate Then
        calSel(calSelNum).fnDate = calSel(calSelNum).stdate
        calSel(calSelNum).stdate = wDate
    End If
    
    If tDates.Recordset.RecordCount > 0 Then
    tDates.Recordset.MoveFirst
    Do
    If CLng(Left(tDates.Recordset.Fields("Room"), 1)) = calSel(calSelNum).roomNo Then
        If ((calSel(calSelNum).stdate >= tDates.Recordset.Fields("Start date") And calSel(calSelNum).stdate <= tDates.Recordset.Fields("End date")) Or (calSel(calSelNum).fnDate >= tDates.Recordset.Fields("Start date") And calSel(calSelNum).fnDate <= tDates.Recordset.Fields("End date"))) Or ((tDates.Recordset.Fields("Start date") >= calSel(calSelNum).stdate) And (tDates.Recordset.Fields("End date") <= calSel(calSelNum).fnDate)) Then
            If calSelNum > 0 Then
                ReDim Preserve calSel(UBound(calSel) - 1)
                calSelNum = calSelNum - 1
                Exit Do
            End If
        End If
    End If
        tDates.Recordset.MoveNext
    Loop Until tDates.Recordset.EOF
    End If
    
'For cc2 = UBound(calSel) To 1 Step -1
cc2 = UBound(calSel)
Do
another = False

    For cc = 1 To UBound(calSel) - 1
        If calSel(cc).roomNo = calSel(cc2).roomNo Then
        
        If calSel(cc2).stdate >= calSel(cc).stdate And _
           calSel(cc2).stdate <= calSel(cc).fnDate And _
           calSel(cc2).fnDate > calSel(cc).fnDate Then
            
            calSel(cc).fnDate = calSel(cc2).fnDate
            'ReDim Preserve calSel(UBound(calSel) - 1)
            calSelNum = calSelNum - 1
            another = True
            'Exit For
        End If
        If calSel(cc2).stdate >= calSel(cc).stdate And _
           calSel(cc2).stdate <= calSel(cc).fnDate And _
           calSel(cc2).fnDate >= calSel(cc).stdate And _
           calSel(cc2).fnDate <= calSel(cc).fnDate Then
            
            'ReDim Preserve calSel(UBound(calSel) - 1)
            calSelNum = calSelNum - 1
            another = True
            'Exit For
        End If
'Debug.Print UBound(calSel)
        If calSel(cc2).fnDate >= calSel(cc).stdate And _
           calSel(cc2).fnDate <= calSel(cc).fnDate And _
           calSel(cc2).stdate < calSel(cc).stdate Then
            
            calSel(cc).stdate = calSel(cc2).stdate
            'ReDim Preserve calSel(UBound(calSel) - 1)
            calSelNum = calSelNum - 1
            another = True
            'Exit For
        End If
        
        If calSel(cc2).stdate - 1 = calSel(cc).fnDate Then
            calSel(cc).fnDate = calSel(cc2).fnDate
            'ReDim Preserve calSel(UBound(calSel) - 1)
            calSelNum = calSelNum - 1
            another = True
            'Exit For
        End If
        
        If calSel(cc2).fnDate + 1 = calSel(cc).stdate Then
            calSel(cc).stdate = calSel(cc2).stdate
            'ReDim Preserve calSel(UBound(calSel) - 1)
            calSelNum = calSelNum - 1
            another = True
            'Exit For
        End If
        End If
    Next
    
    If another = True Then
        ReDim Preserve calSel(UBound(calSel) - 1)
        cc2 = cc2 - 1
        another = False
    End If
    
    If stSel = True Then
        calSelNum = calSelNum + 1
    End If
    stSel = False
    If calSelNum > 1 Then
        selection.Enabled = True
    End If
'Next
Loop Until cc2 = 1 Or another = False

    
    'Debug.Print calSel(calSelNum - 1).roomNo
    Label2.Caption = calSelNum - 1
    
    'Debug.Print xx, yy, wDate
    '"PARAMETERS wkDate DateTime, wkRoom Short; SELECT Dates.Uni, Dates.ID, Dates.[Start Date], Dates.[End Date], Dates.Room, Dates.[No People], Dates.Moveable, Dates.[Date made], Dates.[Time made], Int(Left([Room],1)) AS Expr1 From Dates WHERE (((Dates.[Start Date])<=[wkdate]) AND ((Dates.[End Date])>=[wkdate]) AND ((Int(Left([Room],1)))=[wkRoom]));"
    Set qdftemp = tDates.Database.CreateQueryDef("", "PARAMETERS wkDate DateTime, wkRoom Short; SELECT Dates.Uni, Dates.ID, Dates.[Start Date], Dates.[End Date], Dates.Room, Dates.[No People], Dates.Moveable, Dates.[Date made], Dates.[Time made], Dates.[Quote], Int(Left([Room],1)) AS Expr1 From Dates WHERE (((Dates.[Start Date])<=[wkdate]) AND ((Dates.[End Date])>=[wkdate]) AND ((Int(Left([Room],1)))=[wkRoom]));")
    qdftemp.Parameters!wkDate = wDate
    qdftemp.Parameters!wkRoom = Int(yy + 1)
    tDates.Recordset.Requery qdftemp
       
    If tDates.Recordset.RecordCount = 0 Then
        'Debug.Print "EMPTY"
        custSel.stdate = vbNull
        custSel.fnDate = vbNull
        custSel.roomNoName = vbNull
        custSel.roomNo = 0
        custSel.custid = 0
        custSel.Uni = 0
        custSel.moveable = True
        'custSel.bookmark = 0
        ExtraInfo.Caption = ""
        BookingSel.Enabled = False
    Else
        'Debug.Print "FULL"
        BookingSel.Enabled = True
        custSel.stdate = tDates.Recordset.Fields("Start date")
        custSel.fnDate = tDates.Recordset.Fields("End date")
        custSel.roomNo = CLng(Left(tDates.Recordset.Fields("Room"), 1))
        For croom = 0 To 6
            If croom <> custSel.roomNo - 1 Then
                moveroom(croom).Enabled = True
            Else
                moveroom(croom).Enabled = False
            End If
        Next
        custSel.roomNoName = tDates.Recordset.Fields("Room")
        custSel.Uni = tDates.Recordset.Fields("Uni")
        custSel.custid = tDates.Recordset.Fields("ID")
        custSel.moveable = tDates.Recordset.Fields("Moveable")
        'custSel.bookmark = tDates.Recordset.bookmark
        If tDates.Recordset.Fields("Quote") <> vbNull Then
            ExtraInfo.Caption = tDates.Recordset.Fields("Quote")
        Else
            ExtraInfo.Caption = ""
        End If
        
        cCust.Recordset.MoveFirst
        cCust.Recordset.FindFirst "ID=" & tDates.Recordset("ID")
        uSideBar.Bar("Details").Item("Title").Text = "Title: " & cCust.Recordset.Fields("Title")
        uSideBar.Bar("Details").Item("Forename").Text = "Forname: " & cCust.Recordset.Fields("First name")
        uSideBar.Bar("Details").Item("Surname").Text = "Surname: " & cCust.Recordset.Fields("Last name")
        uSideBar.Bar("Details").Item("Address1").Text = "Address1: " & cCust.Recordset.Fields("Address 1")
        uSideBar.Bar("Details").Item("Address2").Text = "Address2: " & cCust.Recordset.Fields("Address 2")
        uSideBar.Bar("Details").Item("Address3").Text = "Address3: " & cCust.Recordset.Fields("Address 3")
        uSideBar.Bar("Details").Item("County").Text = "County: " & cCust.Recordset.Fields("County")
        uSideBar.Bar("Details").Item("Town").Text = "Town: " & cCust.Recordset.Fields("Town")
        uSideBar.Bar("Details").Item("Country").Text = "Country: " & cCust.Recordset.Fields("Country")
        uSideBar.Bar("Details").Item("Postcode").Text = "Postcode: " & cCust.Recordset.Fields("Postcode")
    End If

    'tDates.Recordset.Requery
    'Debug.Print "D" & tDates.Recordset.RecordCount
    Dim qdftemp2 As QueryDef
    Set qdftemp2 = tDates.Database.CreateQueryDef("", "PARAMETERS StDate DateTime, EnDate DateTime; SELECT Dates.Uni, Dates.ID, Dates.[Start Date], Dates.[End Date], Dates.Room, Dates.[No People], Dates.Moveable, Dates.[Date made], Dates.[Time made], Dates.[Quote] From Dates WHERE (((Dates.[Start Date])>[stDate]) AND ((Dates.[End Date])<[enDate]));")
    qdftemp2.Parameters!stdate = usDate - numdays
    qdftemp2.Parameters!endate = ufDate + numdays
    tDates.Recordset.Requery qdftemp2
    tcal_Paint
End If

If Button = 2 Then
tcal_MouseUp 1, Shift, X, Y
If custSel.custid > 0 Then
    PopupMenu BookingSel
Else
    If selection.Enabled = True Then
        PopupMenu selection
    End If
End If
End If

'Command2_Click
End Sub

Private Sub tcal_Paint()
Dim xx As Long
Dim yy As Long
Dim xad As Long
Dim yad As Long
Dim fcol As Long
Dim wDate As Date
Dim monthcol As Boolean
Dim tmpRoom As String
Dim tmpdate As Date
Dim aHol As Boolean
Dim nc As Long
Dim tmpname As String
Dim F As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
Dim FONTSIZE As Integer



'    Dim qdfTemp As QueryDef
'    Set qdfTemp = tDates.Database.CreateQueryDef("", "PARAMETERS StDate DateTime, EnDate DateTime; SELECT Dates.Uni, Dates.ID, Dates.[Start Date], Dates.[End Date], Dates.Room, Dates.[No People], Dates.Moveable, Dates.[Date made], Dates.[Time made], Dates.[Quote] From Dates WHERE (((Dates.[Start Date])>[stDate]) AND ((Dates.[End Date])<[enDate]));")
'    qdfTemp.Parameters!stDate = usDate - numDays
'    qdfTemp.Parameters!endate = ufDate + numDays
'    tDates.Recordset.Requery qdfTemp
    
xad = tcal.ScaleWidth / numdays
yad = tcal.ScaleHeight / 7

Do
xad = xad - 1
Loop Until xad * numdays < tcal.ScaleWidth
Do
yad = yad - 1
Loop Until yad * 7 < tcal.ScaleHeight

tnums.Width = xad * numdays
wDays.Width = xad * numdays
tnums.Cls
tNames.Height = yad * 7
tNames.Cls
tcal.Cls
wDays.Cls

For yy = 0 To 6
    
    If yy = 0 Then tmpRoom = "1 Snowhill" & vbCrLf & "(Double bed)"
    If yy = 1 Then tmpRoom = "2 Square View" & vbCrLf & "(Double bed)"
    If yy = 2 Then tmpRoom = "3 Ladywell" & vbCrLf & "*(Double bed)"
    If yy = 3 Then tmpRoom = "4 Littledale" & vbCrLf & "(Double bed)"
    If yy = 4 Then tmpRoom = "5 Nickynook" & vbCrLf & "(Single bed +" & vbCrLf & "Kingsize bed)"
    If yy = 5 Then tmpRoom = "6 Wyresdale" & vbCrLf & "(Kingsize bed)"
    If yy = 6 Then tmpRoom = "7/3a Arbor" & vbCrLf & "*(single bed)"
    If tNames.TextHeight(tmpRoom) > yad Then
        If yy = 0 Then tmpRoom = "1 Snowhill"
        If yy = 1 Then tmpRoom = "2 Square View"
        If yy = 2 Then tmpRoom = "3 Ladywell"
        If yy = 3 Then tmpRoom = "4 Littledale"
        If yy = 4 Then tmpRoom = "5 Nickynook"
        If yy = 5 Then tmpRoom = "6 Wyresdale"
        If yy = 6 Then tmpRoom = "7/3a Arbor"
    End If

    tNames.CurrentX = 0
    tNames.CurrentY = yy * yad + yad / 2 - tNames.TextHeight(tmpRoom) / 2
    tNames.Print tmpRoom
    
    wDate = sDate
    For xx = 0 To numdays - 1
aHol = False
If tHolidays.Recordset.RecordCount > 0 Then
tHolidays.Recordset.MoveFirst
Do
tmpdate = tHolidays.Recordset.Fields("Hol Date")
If tHolidays.Recordset.Fields("Yearly") Or tmpdate = wDate Then
    If wDate = DateSerial(Year(wDate), Month(tmpdate), Day(tmpdate)) Or tmpdate = wDate Then
        aHol = True
    End If
End If
tHolidays.Recordset.MoveNext
Loop Until tHolidays.Recordset.EOF
End If

    If yy = 0 Then

        If monthcol = True Then
            tnums.ForeColor = RGB(200, 0, 0)
            wDays.ForeColor = tnums.ForeColor
        Else
            tnums.ForeColor = 0
            wDays.ForeColor = tnums.ForeColor
        End If
                
        
            
        If Month(wDate - 1) <> Month(wDate) Then
            tnums.CurrentX = xx * xad
            If tnums.CurrentX + tnums.TextWidth(Format(wDate + 1, "mmm yyyy")) > numdays * xad Then
                tnums.CurrentX = numdays * xad - tnums.TextWidth(Format(wDate + 1, "mmm yyyy"))
            End If
            tnums.CurrentY = tnums.TextHeight("X")
            tnums.Print Format(wDate + 1, "mmm yyyy")
        End If
        If xx = 0 Then
            tnums.CurrentX = 0
            tnums.CurrentY = 0
            tnums.Print Format(wDate, "mmm yyyy")
        End If
        
        If aHol = True Then
            wDays.ForeColor = RGB(250, 0, 250)
            tnums.ForeColor = RGB(250, 0, 250)
        End If
        
        tnums.CurrentX = xx * xad - (tnums.TextWidth(Trim(Day(wDate))) / 2) + (xad / 2)
        tnums.CurrentY = tnums.Height - tnums.TextHeight("X")
        wDays.CurrentX = xx * xad - (wDays.TextWidth(Trim(Day(wDate))) / 2) + (xad / 2)
        wDays.CurrentY = wDays.Height - wDays.TextHeight("X")
                        
        tnums.Print Trim(Day(wDate))
        wDays.Print Left(Format(wDate, "ddd"), 1)
        
        If Month(wDate + 1) <> Month(wDate) Then
            If monthcol = False Then
                monthcol = True
            Else
                monthcol = False
            End If
        End If
    
    End If
    
    wDate = wDate + 1
    
        If xx Mod 2 = 0 Then
            fcol = gCols.Column1
            tcal.ForeColor = gCols.ColumnText1
        Else
            fcol = gCols.Column2
            tcal.ForeColor = gCols.ColumnText2
        End If
        tcal.Line (xx * xad, yy * yad)-(xx * xad + xad, yy * yad + yad), fcol, BF
    
    tcal.CurrentX = xx * xad + xad / 2 - (tcal.TextWidth(Trim(Day(wDate - 1))) / 2)
    tcal.CurrentY = yy * yad + yad / 2 - (tcal.TextHeight(Trim(Day(wDate - 1))) / 2)
    If aHol = True Then
        tcal.ForeColor = gCols.Holiday
    End If
    tcal.Print Trim(Day(wDate - 1))
    
    Next
    tcal.Line (0, yy * yad)-(numdays * xad, yy * yad), RGB(100, 100, 100)
Next

Dim rNum As Long
Dim stroom As Long
Dim enroom As Long

If tDates.Recordset.RecordCount > 0 Then
tDates.Recordset.MoveFirst
    Do
        rNum = CLng(Mid(tDates.Recordset.Fields("Room"), 1, 1))
        stroom = tDates.Recordset.Fields("Start date") - usDate
        enroom = tDates.Recordset.Fields("End date") - usDate
            tcal.Line (stroom * xad + 5, rNum * yad - yad + 5)-(enroom * xad + xad - 5, rNum * yad - 5), gCols.BookedColors, BF
        If tDates.Recordset.Fields("Moveable") = False Then
            tcal.FillColor = gCols.NonMovable
            tcal.FillStyle = 0
            tcal.Line (stroom * xad + 5, rNum * yad - yad + 5)-(enroom * xad + xad - 5, rNum * yad - 5), QBColor(15), B
            tcal.FillColor = gCols.SameIDSelectBooked
            tcal.FillStyle = 1
        End If
        If tDates.Recordset.Fields("ID") = custSel.custid Then
            'tcal.FillStyle = 7
            tcal.Line (stroom * xad + 5, rNum * yad - yad + 5)-(enroom * xad + xad - 5, rNum * yad - 5), gCols.SameIDSelectBooked, B
            'tcal.FillStyle = 1
        End If
        
If tcal.BackColor = QBColor(15) Then
tcal.ForeColor = QBColor(15)
cCust.Recordset.MoveFirst
cCust.Recordset.FindFirst "ID=" & tDates.Recordset.Fields("ID")
  
  FONTSIZE = Val(tcal.FONTSIZE)

    F.lfEscapement = 10 * Val(90) 'rotation angle, in tenths
    FontName = "Arial" + Chr$(0) 'null terminated
    F.lfFacename = FontName
    F.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(F)
    hPrevFont = SelectObject(tcal.hdc, hFont)
    tcal.CurrentY = rNum * yad - 10
    tcal.CurrentX = stroom * xad + 4
    tcal.Print cCust.Recordset.Fields("Last name")
  
'  Clean up, restore original font
  hFont = SelectObject(tcal.hdc, hPrevFont)
  DeleteObject hFont
  
        'tcal.Font = "Arial Rotated"
        'cCust.Recordset.MoveFirst
        'cCust.Recordset.FindFirst "ID=" & tDates.Recordset.Fields("ID")
        'tcal.CurrentY = rNum * yad - 5 - tcal.TextHeight("X")
        'tcal.ForeColor = 0
        'tmpname = cCust.Recordset.Fields("Last name")
        '    For nc = 1 To Len(tmpname)
        '        tcal.FONTSIZE = 7
        '        tcal.CurrentX = stroom * xad + 8
        '        If nc > 1 Then
        '        tcal.CurrentY = tcal.CurrentY - tcal.TextWidth(Mid(tmpname, nc, 1))
        '        tcal.CurrentY = tcal.CurrentY - tcal.TextHeight(Mid(tmpname, nc - 1, 1))
        '        End If
        '        tcal.ForeColor = QBColor(15)
        '        tcal.Print (Mid(tmpname, nc, 1))
        '        tcal.FONTSIZE = 8
        '    Next
            
        tDates.Font = "MS Sans Serif"
End If

        tDates.Recordset.MoveNext

    Loop Until tDates.Recordset.EOF
End If

stroom = custSel.stdate - usDate
enroom = custSel.fnDate - usDate
tcal.DrawWidth = 2
'tcal.FillStyle = 7
tcal.Line (stroom * xad + 5, custSel.roomNo * yad - yad + 5)-(enroom * xad + xad - 5, custSel.roomNo * yad - 5), gCols.SelectBooked, BF
'tcal.FillStyle = 1
tcal.DrawWidth = 1

Dim tmpc As Long
For tmpc = 1 To UBound(calSel)
    If calSel(tmpc).stdate <> 0 Then
    stroom = calSel(tmpc).stdate - usDate
    enroom = calSel(tmpc).fnDate - usDate
    tcal.Line (stroom * xad + 5, calSel(tmpc).roomNo * yad - yad + 5)-(enroom * xad + xad - 5, calSel(tmpc).roomNo * yad - 5), QBColor(9), B
    End If
Next

tcal.Line (xad * numdays, 0)-(tcal.ScaleWidth, tcal.ScaleHeight), tcal.BackColor, BF
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cc As Long
Dim newRoom As String
Dim qdftemp As QueryDef

If Button.Key = "sidebar" Then
    If uSideBar.Visible = True Then
        tcal.Left = tcal.Left - uSideBar.Width
        tNames.Left = tNames.Left - uSideBar.Width
        wDays.Left = wDays.Left - uSideBar.Width
        tnums.Left = tnums.Left - uSideBar.Width
        uSideBar.Visible = False
        List1.Visible = False
        'Button.Value = tbrUnpressed
        Toolbar1.Buttons("sidebar").Value = tbrUnpressed
    Else
        uSideBar.Visible = True
        List1.Visible = True
        tcal.Left = tcal.Left + uSideBar.Width
        tNames.Left = tNames.Left + uSideBar.Width
        wDays.Left = wDays.Left + uSideBar.Width
        tnums.Left = tnums.Left + uSideBar.Width
        'Button.Value = tbrPressed
        Toolbar1.Buttons("sidebar").Value = tbrPressed
    End If
    
    If frmHol.Visible = False Then
        'frmHol.Visible = True
        Toolbar1.Buttons("holidays").Value = tbrUnpressed
        'Button.Value = tbrPressed
    Else
        'Unload frmHol
        'Button.Value = tbrUnpressed
        Toolbar1.Buttons("holidays").Value = tbrPressed
    End If
    
    Form_Resize
ElseIf Button.Key = "holidays" Then
    If frmHol.Visible = False Then
        frmHol.Visible = True
        Toolbar1.Buttons("holidays").Value = tbrPressed
        'Button.Value = tbrPressed
    Else
        Unload frmHol
        'Button.Value = tbrUnpressed
        Toolbar1.Buttons("holidays").Value = tbrUnpressed
    End If
        If uSideBar.Visible = True Then
            Toolbar1.Buttons("sidebar").Value = tbrPressed
        Else
            Toolbar1.Buttons("sidebar").Value = tbrUnpressed
        End If
ElseIf Button.Key = "New Customer" Then
    NewCustDetails.Show vbModal, Me
    If Tag <> "" Then
        If calSelNum > 1 Then
            For cc = 1 To UBound(calSel)
                tDates.Recordset.AddNew
                tDates.Recordset.Fields("ID") = CLng(Tag)
                tDates.Recordset.Fields("Start Date") = calSel(cc).stdate
                tDates.Recordset.Fields("End Date") = calSel(cc).fnDate
                If calSel(cc).roomNo = 1 Then newRoom = "1 Snowhill"
                If calSel(cc).roomNo = 2 Then newRoom = "2 Square View"
                If calSel(cc).roomNo = 3 Then newRoom = "3 Ladyewell"
                If calSel(cc).roomNo = 4 Then newRoom = "4 Littledale"
                If calSel(cc).roomNo = 5 Then newRoom = "5 Nickynook"
                If calSel(cc).roomNo = 6 Then newRoom = "6 Wyresdale"
                If calSel(cc).roomNo = 7 Then newRoom = "7 /3a Arbor"
                tDates.Recordset.Fields("Room") = newRoom
                tDates.Recordset.Fields("Date Made") = Date
                tDates.Recordset.Fields("Time Made") = Time
                tDates.Recordset.Update
            Next
            'Clear selections
            cCust.Recordset.Requery
            tDates.Recordset.Requery
            Command2_Click
            Call Reset
        End If
        Tag = ""
    End If
ElseIf Button.Key = "OldCustomer" Then
    If calSelNum > 1 Then tOldCust.Show vbModal, Me
    If Tag <> "" Then
        If calSelNum > 1 Then
            For cc = 1 To UBound(calSel)
                tDates.Recordset.AddNew
                tDates.Recordset.Fields("ID") = CLng(Tag)
                tDates.Recordset.Fields("Start Date") = calSel(cc).stdate
                tDates.Recordset.Fields("End Date") = calSel(cc).fnDate
                If calSel(cc).roomNo = 1 Then newRoom = "1 Snowhill"
                If calSel(cc).roomNo = 2 Then newRoom = "2 Square View"
                If calSel(cc).roomNo = 3 Then newRoom = "3 Ladyewell"
                If calSel(cc).roomNo = 4 Then newRoom = "4 Littledale"
                If calSel(cc).roomNo = 5 Then newRoom = "5 Nickynook"
                If calSel(cc).roomNo = 6 Then newRoom = "6 Wyresdale"
                If calSel(cc).roomNo = 7 Then newRoom = "7 /3a Arbor"
                tDates.Recordset.Fields("Room") = newRoom
                tDates.Recordset.Fields("Date Made") = Date
                tDates.Recordset.Fields("Time Made") = Time
                tDates.Recordset.Update
            Next
            'Clear selections
            cCust.Recordset.Requery
            tDates.Recordset.Requery
            Command2_Click
            Call Reset
        End If
        Tag = ""
    End If
ElseIf Button.Key = "PrintBookings" Then
    Dim useDate As Date
    useDate = CDate(InputBox("Enter date to use", "Enter Date", Date))
    Set qdftemp = tDates.Database.CreateQueryDef("", "PARAMETERS stdate DateTime; SELECT Dates.Uni, Dates.ID, Dates.Room, Dates.[Start Date], Dates.[End Date], Dates.[No People], Dates.Moveable, Dates.[Date made], Dates.[Time made], Dates.Quote From Dates Where (((Dates.[Start Date]) = [stdate])) ORDER BY Dates.Room;")
    qdftemp.Parameters!stdate = useDate
    tDates.Recordset.Requery qdftemp
    Dim tmpDates(7) As selDate
    Dim tmpRN As Long
    
    If tDates.Recordset.RecordCount > 0 Then
        Do
            
            tmpRN = Mid(tDates.Recordset.Fields("Room"), 1, 1)
            tmpDates(tmpRN).Uni = tDates.Recordset.Fields("Uni")
            tmpDates(tmpRN).custid = tDates.Recordset.Fields("ID")
            tmpDates(tmpRN).stdate = tDates.Recordset.Fields("Start date")
            tmpDates(tmpRN).fnDate = tDates.Recordset.Fields("End date")
            tmpDates(tmpRN).noPeople = tDates.Recordset.Fields("No People")
            tmpDates(tmpRN).roomNoName = tDates.Recordset.Fields("Room")
            tmpDates(tmpRN).Quote = IIf(IsNull(tDates.Recordset.Fields("Quote")), "", tDates.Recordset.Fields("Quote"))
            cCust.Recordset.MoveFirst
            cCust.Recordset.FindFirst "ID=" & tmpDates(tmpRN).custid
            tmpDates(tmpRN).CustName = cCust.Recordset.Fields("Customer Name")
            tmpDates(tmpRN).Address = cCust.Recordset.Fields("Address 1") & vbCrLf
            tmpDates(tmpRN).Address = tmpDates(tmpRN).Address & cCust.Recordset.Fields("Address 2") & vbCrLf
            tmpDates(tmpRN).Address = tmpDates(tmpRN).Address & cCust.Recordset.Fields("Address 3") & vbCrLf
            tmpDates(tmpRN).Address = tmpDates(tmpRN).Address & cCust.Recordset.Fields("Town") & vbCrLf
            tmpDates(tmpRN).Address = tmpDates(tmpRN).Address & cCust.Recordset.Fields("County") & vbCrLf
            tmpDates(tmpRN).Address = tmpDates(tmpRN).Address & cCust.Recordset.Fields("Postcode") & vbCrLf
            tmpDates(tmpRN).Address = tmpDates(tmpRN).Address & cCust.Recordset.Fields("Country") '& vbCrLf
            
            tDates.Recordset.MoveNext
        Loop Until tDates.Recordset.EOF
            Printer.Orientation = vbPRORPortrait
            Printer.FONTSIZE = 30
            Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("The Priory Hotel") / 2
            Printer.Print "The Priory Hotel"
            Printer.FONTSIZE = 20
            Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Format(useDate, "dddd dd mmmm, yyyy")) / 2
            Printer.Print Format(useDate, "dddd d mmmm, yyyy")
            Printer.Print
        
        For tmpRN = 1 To 7
            Printer.FONTSIZE = 15
            If tmpDates(tmpRN).custid > 0 Then
            Printer.CurrentX = Printer.ScaleWidth * 0.5
            If tmpRN = 7 And tmpDates(tmpRN).custid = tmpDates(3).custid Then Exit For
            If tmpRN = 3 And tmpDates(tmpRN).custid = tmpDates(7).custid Then
                Printer.Print tmpDates(tmpRN).roomNoName & " + " & tmpDates(7).roomNoName
            Else
                Printer.Print tmpDates(tmpRN).roomNoName
            End If
            Printer.CurrentY = Printer.CurrentY - Printer.TextHeight("X")
            Printer.FONTSIZE = 10
            Printer.FontBold = True
            Printer.Print tmpDates(tmpRN).CustName
            Printer.FontBold = False
            Printer.FONTSIZE = 8
            Printer.Print
            Printer.Print tmpDates(tmpRN).Address
            Printer.CurrentY = Printer.CurrentY - Printer.TextHeight("X") * 7
            Printer.CurrentX = Printer.ScaleWidth * 0.3
            Printer.Print "Number People:- " & tmpDates(tmpRN).noPeople
            Printer.CurrentX = Printer.ScaleWidth * 0.3
            Printer.Print "Comments:- " & tmpDates(tmpRN).Quote
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            'Printer.Print
            Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY), 0
            Printer.Print
            End If
        Next
        
        'BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth, Picture1.ScaleHeight, Printer.hdc, 0, 0, vbSrcCopy
        'Picture1.Refresh
        'Printer.KillDoc
        Printer.EndDoc
    Else
        MsgBox "No Records"
    End If
    
    Call Reset
ElseIf Button.Key = "PrintCalendar" Then
    Printer.Orientation = vbPRORLandscape
    Printer.ScaleMode = 7
    'Debug.Print Printer.ScaleWidth & " " & Printer.ScaleHeight
    gCols.Column1 = QBColor(15)
    gCols.Column2 = RGB(200, 200, 200)
    gCols.BookedColors = 0
    tcal.Width = tcal.ScaleX(Printer.ScaleWidth, vbCentimeters, vbPixels) - tNames.ScaleWidth
    tcal.Height = tcal.ScaleY(Printer.ScaleHeight, vbCentimeters, vbPixels) - tnums.ScaleHeight - wDays.ScaleHeight
    tnums.Width = tcal.Width
    tNames.Height = tcal.Height
    tcal.BackColor = QBColor(15)
    tNames.BackColor = QBColor(15)
    tnums.BackColor = QBColor(15)
    tcal_Paint
    
    'Printer.ScaleMode = 3
    'Debug.Print Printer.ScaleWidth & " " & Printer.ScaleHeight
    'Debug.Print tcal.ScaleWidth
    Printer.PaintPicture tcal.Image, tNames.ScaleX(tNames.ScaleWidth, vbPixels, vbCentimeters), tnums.ScaleY(tnums.ScaleHeight, vbPixels, vbCentimeters)
    Printer.PaintPicture tNames.Image, 0, tnums.ScaleY(tnums.ScaleHeight, vbPixels, vbCentimeters)
    Printer.PaintPicture tnums.Image, tNames.ScaleX(tNames.ScaleWidth, vbPixels, vbCentimeters), 0
    Printer.PaintPicture wDays.Image, tNames.ScaleX(tNames.ScaleWidth, vbPixels, vbCentimeters), tnums.ScaleY(tnums.ScaleHeight, vbPixels, vbCentimeters) + tcal.ScaleY(tcal.ScaleHeight, vbPixels, vbCentimeters)
    'Printer.KillDoc
    Printer.EndDoc
    Form_Resize
    Call SetDefaultColours
    tcal.BackColor = vbButtonFace
    tNames.BackColor = vbButtonFace
    tnums.BackColor = vbButtonFace
    tcal_Paint
End If
'tDates.Recordset.Close
'tdates.Recordset.OpenRecordset("Dates", dbOpendynaset, dbRunAsync)
End Sub

Private Sub undoSel_Click()
Command3_Click
End Sub
