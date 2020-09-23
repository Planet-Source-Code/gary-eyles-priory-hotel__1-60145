VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form tOldCust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Old Customer"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataField       =   "Postcode"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataField       =   "County"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataField       =   "Address 3"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataField       =   "Address 2"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataField       =   "Address 1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataField       =   "Customer Name"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   600
      Width           =   3255
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "OldCust.frx":0000
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8361
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "Last Name"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "alpha cust names"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Postcode"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "County"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Address 3"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Address 2"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Address 1"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "tOldCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Tag = Text7.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DBList1_Click()
Data1.Recordset.Bookmark = DBList1.SelectedItem
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Hotel.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Data1.Recordset.Close
End Sub

Private Sub Text7_Change()
'DBList1.MatchedWithList
End Sub
