VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form CustDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Details"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   781
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PostcodesXY"
      Top             =   6840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Countries"
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Alpha County"
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   240
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   11295
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "CustDetails.frx":0000
         Height          =   5415
         Left            =   120
         OleObjectBlob   =   "CustDetails.frx":0014
         TabIndex        =   35
         Top             =   240
         Width           =   10935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   11295
      Begin VB.PictureBox pMap 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   7200
         Picture         =   "CustDetails.frx":09E7
         ScaleHeight     =   367
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   229
         TabIndex        =   36
         Top             =   240
         Width           =   3470
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "CustDetails.frx":465A
         DataField       =   "Country"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   3960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         ListField       =   "Country"
         Text            =   "DBCombo3"
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "CustDetails.frx":466E
         DataField       =   "County"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   2880
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         BackColor       =   14737632
         ListField       =   "County"
         Text            =   "DBCombo2"
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "ID"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Text            =   "Uni"
         Top             =   360
         Width           =   3750
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "First Name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "First Name"
         Top             =   1080
         Width           =   3750
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Last Name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "Last Name"
         Top             =   1440
         Width           =   3750
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Address 1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Text            =   "Address 1"
         Top             =   1800
         Width           =   3750
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Address 2"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "Address 2"
         Top             =   2160
         Width           =   3750
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Address 3"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "Address 3"
         Top             =   2520
         Width           =   3750
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Postcode"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "Postcode"
         Top             =   3600
         Width           =   3750
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Telephone"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Text            =   "Telephone"
         Top             =   4320
         Width           =   3750
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Email"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Text            =   "Email"
         Top             =   4680
         Width           =   3750
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Customer Name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Text            =   "Customer Name"
         Top             =   5400
         Visible         =   0   'False
         Width           =   3750
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Comments"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Text            =   "Comments"
         Top             =   5040
         Width           =   3750
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "CustDetails.frx":4682
         DataField       =   "Town"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   3240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         ListField       =   "Town"
         BoundColumn     =   ""
         Text            =   "DBCombo1"
      End
      Begin VB.Label Label1 
         Caption         =   "Uni"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Title"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Forename"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Surname"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Address 1"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Address 2"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Address 3"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "County"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Town"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Postcode"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Telephone"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Email"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Country"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Comments"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   5040
         Width           =   1095
      End
      Begin MSForms.ComboBox ComboBox1 
         Bindings        =   "CustDetails.frx":46A0
         CausesValidation=   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   0
         Top             =   720
         Width           =   3735
         VariousPropertyBits=   746604571
         BackColor       =   14737632
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6588;450"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Dates"
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Alpha Towns"
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer Details"
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSForms.TabStrip TabStrip1 
      Height          =   6615
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   11535
      ListIndex       =   0
      Size            =   "20346;11668"
      Items           =   "Tab1;Tab2;"
      TipStrings      =   ";;"
      Names           =   "Tab1;Tab2;"
      NewVersion      =   -1  'True
      TabsAllocated   =   2
      Tags            =   ";;"
      TabData         =   2
      Accelerator     =   ";;"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      TabState        =   "3;3"
   End
   Begin VB.Label Label15 
      Caption         =   "Title"
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   8640
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "CustDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Color1 As Long
Dim Color2 As Long
Dim FirstLoad As Boolean

Private Sub Combo1_Change()

End Sub

Private Sub ComboBox1_Change()
Label15.Caption = ComboBox1.Text
End Sub

Private Sub ComboBox1_GotFocus()
ComboBox1.BackColor = Color1
'ComboBox1.DropDown
End Sub

Private Sub ComboBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
ComboBox1.DropDown
End Sub

Private Sub ComboBox1_LostFocus()
ComboBox1.BackColor = Color2
End Sub

Private Sub Command1_Click()
Data1.Recordset.CancelUpdate
Unload Me
End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
Unload Me
End Sub

Private Sub DBCombo1_GotFocus()
DBCombo1.BackColor = Color1
SendKeys "%+{DOWN}"
End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub DBCombo1_LostFocus()
DBCombo1.BackColor = Color2
End Sub

Private Sub DBCombo2_GotFocus()
DBCombo2.BackColor = Color1
SendKeys "%+{DOWN}"
End Sub

Private Sub DBCombo2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub DBCombo2_LostFocus()
DBCombo2.BackColor = Color2
End Sub

Private Sub DBCombo3_GotFocus()
DBCombo3.BackColor = Color1
SendKeys "%+{DOWN}"
End Sub

Private Sub DBCombo3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub DBCombo3_LostFocus()
DBCombo3.BackColor = Color2
End Sub

Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex = 0 Then Cancel = True
If ColIndex = 1 Then Cancel = True
If ColIndex = 2 Then Cancel = True
If ColIndex = 3 Then Cancel = True
If ColIndex = 4 Then Cancel = True
If ColIndex = 6 Then Cancel = True
If ColIndex = 7 Then Cancel = True
If ColIndex = 8 Then Cancel = True
End Sub

Private Sub Form_Activate()
Dim qdftemp As QueryDef
Set qdftemp = Data3.Database.CreateQueryDef("", "PARAMETERS custID Long; SELECT Dates.Uni, Dates.ID, Dates.[Start Date], Dates.[End Date], Dates.Room, Dates.[No People], Dates.Moveable, Dates.[Date made], Dates.[Time made], Dates.Quote From Dates WHERE (((Dates.ID)=[custid]));")
qdftemp.Parameters!custid = CLng(Text1.Text)
Data3.Recordset.Requery qdftemp

DBGrid1.Columns(0).Width = 500
DBGrid1.Columns(1).Width = 500
DBGrid1.Columns(9).Width = 2000
'DBGrid1.Columns(1).Enabled = False
'DBCombo1.ListField = "Town"

Data1.Recordset.edit
If FirstLoad = False Then
'    Data2.Recordset.Sort = "Town"
'    Do
'    ComboBox2.AddItem Data2.Recordset.Fields("Town")
'    Data2.Recordset.MoveNext
'    Loop Until Data2.Recordset.EOF
End If
FirstLoad = True

'Data6.Recordset.MoveFirst
Dim xx As Long, yy As Long
'Do
'xx = Data6.Recordset.Fields("x") / 2850
'yy = Data6.Recordset.Fields("y") / 2950
'yy = pMap.ScaleHeight - yy
'pMap.Circle (xx, yy), 0
'Data6.Recordset.MoveNext
'Loop Until Data6.Recordset.EOF


If Text10.Text <> "" Then
    Dim qdftemp2 As QueryDef
    Dim tmppost As String
    Set qdftemp2 = Data6.Database.CreateQueryDef("", "PARAMETERS tPost Text ( 255 ); SELECT PostcodesXY.postcode, PostcodesXY.x, PostcodesXY.y, PostcodesXY.latitude, PostcodesXY.longitude, Mid([tpost],1,Len([postcode])) AS Expr1 From PostcodesXY WHERE (((Mid([tpost],1,Len([postcode])))=[postcode]));")
    qdftemp2.Parameters!tpost = Text10.Text
    Data6.Recordset.Requery qdftemp2
    'Debug.Print Data6.Recordset.RecordCount
If Data6.Recordset.RecordCount > 0 Then
    Data6.Recordset.MoveFirst
    Do
        If Len(Data6.Recordset.Fields("Postcode")) > Len(tmppost) Then
            xx = Data6.Recordset.Fields("x") / 2850
            yy = Data6.Recordset.Fields("y") / 2950
            yy = pMap.ScaleHeight - yy
        End If
        Data6.Recordset.MoveNext
    Loop Until Data6.Recordset.EOF
    pMap.ForeColor = QBColor(0)
    pMap.FillStyle = 0
    pMap.Circle (xx, yy), 3
    
End If
End If
End Sub

Private Sub Form_Load()
ComboBox1.AddItem "Dr"
ComboBox1.AddItem "Father"
ComboBox1.AddItem "Mr"
ComboBox1.AddItem "Mrs"
ComboBox1.AddItem "Mr & Mrs"
ComboBox1.AddItem "Miss"
ComboBox1.AddItem "Rev"
TabStrip1.Tabs(0).Caption = "Customer Details"
TabStrip1.Tabs(1).Caption = "Customer Bookings"
Data1.DatabaseName = App.Path & "\Hotel.mdb"
Data2.DatabaseName = App.Path & "\Hotel.mdb"
Data3.DatabaseName = App.Path & "\Hotel.mdb"
Data4.DatabaseName = App.Path & "\Hotel.mdb"
Data5.DatabaseName = App.Path & "\Hotel.mdb"
Data6.DatabaseName = App.Path & "\Hotel.mdb"
Color1 = QBColor(15)
Color2 = &HE0E0E0


End Sub

Private Sub Form_Unload(Cancel As Integer)
FirstLoad = False
Data1.Recordset.Close
Data2.Recordset.Close
Data3.Recordset.Close
Data4.Recordset.Close
Data5.Recordset.Close
End Sub

Private Sub Label15_Change()
ComboBox1.Text = Label15.Caption
End Sub

Private Sub TabStrip1_Change()
If TabStrip1.SelectedItem.Index = 0 Then
    Frame1.Visible = True
    Frame2.Visible = False
    'MSHFlexGrid1.Refresh
Else
    Frame1.Visible = False
    Frame2.Visible = True
    'MSHFlexGrid1.Refresh
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_Change()
Text13.Text = Text3 & " " & Text4
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = Color1
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = Color2
'Text3.Text = Proper(Text3.Text)
End Sub

Private Sub Text4_Change()
Text13.Text = Text3 & " " & Text4
End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = Color1
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
Text4.BackColor = Color2
End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = Color1
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5.BackColor = Color2
End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = Color1
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
Text6.BackColor = Color2
End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = Color1
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7.BackColor = Color2
End Sub

Private Sub Text10_GotFocus()
Text10.BackColor = Color1
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
End Sub

Private Sub Text10_LostFocus()
Text10.BackColor = Color2
End Sub

Private Sub Text11_GotFocus()
Text11.BackColor = Color1
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)
End Sub

Private Sub Text11_LostFocus()
Text11.BackColor = Color2
End Sub

Private Sub Text12_GotFocus()
Text12.BackColor = Color1
Text12.SelStart = 0
Text12.SelLength = Len(Text12.Text)
End Sub

Private Sub Text12_LostFocus()
Text12.BackColor = Color2
End Sub

Private Sub Text13_GotFocus()
Text13.BackColor = Color1
End Sub

Private Sub Text13_LostFocus()
Text13.BackColor = Color2
End Sub

Private Sub Text14_GotFocus()
Text14.BackColor = Color1
Text14.SelStart = 0
Text14.SelLength = Len(Text4.Text)
End Sub

Private Sub Text14_LostFocus()
Text14.BackColor = Color2
End Sub


