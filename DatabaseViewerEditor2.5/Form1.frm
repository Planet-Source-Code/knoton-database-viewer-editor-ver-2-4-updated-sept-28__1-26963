VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  KNOTON´S DATABASE VIEWER/EDITOR"
   ClientHeight    =   5565
   ClientLeft      =   1635
   ClientTop       =   2400
   ClientWidth     =   9390
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   9390
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5190
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5821
            MinWidth        =   2646
            Text            =   "CLOSED"
            TextSave        =   "CLOSED"
            Object.ToolTipText     =   "Shows ConnectionState"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "READ"
            TextSave        =   "READ"
            Object.ToolTipText     =   "Shows Recordset type"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "Shows database type"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "Shows various info"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "Current Database"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9128
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Login to server"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCon(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCon(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCon(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCon(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblInfo(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblInfo(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCon(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCon(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCon(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCon(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdClose"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCon"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lstDatBas"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "View/Edit Data"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3(4)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dbgrid"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lstTables"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lstcol"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lstView"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdUpdate"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdDel"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Custom SQL Query"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3(3)"
      Tab(2).Control(1)=   "Label3(5)"
      Tab(2).Control(2)=   "txtTsql"
      Tab(2).Control(3)=   "cmdClearQuery"
      Tab(2).Control(4)=   "cmdOpen"
      Tab(2).Control(5)=   "cmdSave"
      Tab(2).Control(6)=   "cmdTSql"
      Tab(2).Control(7)=   "CDADB"
      Tab(2).Control(8)=   "chkNoRetRS"
      Tab(2).Control(9)=   "lstProc"
      Tab(2).Control(10)=   "lstSqlView"
      Tab(2).ControlCount=   11
      Begin VB.CommandButton cmdDel 
         Caption         =   "DELETE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   33
         ToolTipText     =   "Delete current record"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "UPDATE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   32
         ToolTipText     =   "Update Recordset"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ListBox lstSqlView 
         Height          =   1035
         Left            =   -71400
         TabIndex        =   30
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ListBox lstProc 
         Height          =   1035
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ListBox lstView 
         Height          =   1230
         Left            =   -69600
         TabIndex        =   27
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox chkNoRetRS 
         Caption         =   "No Return Value"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -69600
         TabIndex        =   25
         ToolTipText     =   "Check this if you dont want a return vale, for ex when you create a stored procedure"
         Top             =   600
         Width           =   1935
      End
      Begin MSComDlg.CommonDialog CDADB 
         Left            =   -66600
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Open Access Database"
         Filter          =   "Access Database (*.mdb)|*.mdb"
      End
      Begin VB.ListBox lstDatBas 
         Height          =   2595
         ItemData        =   "Form1.frx":0496
         Left            =   120
         List            =   "Form1.frx":0498
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Choose a Database"
         Top             =   720
         Width           =   4935
      End
      Begin VB.CommandButton cmdTSql 
         Caption         =   "EXECUTE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   19
         ToolTipText     =   "Execute Custom T-SQL Query"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70920
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72240
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClearQuery 
         Caption         =   "Clear Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.ListBox lstcol 
         Height          =   1230
         ItemData        =   "Form1.frx":049A
         Left            =   -72240
         List            =   "Form1.frx":049C
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "All Columns in current Table"
         Top             =   720
         Width           =   2535
      End
      Begin VB.ListBox lstTables 
         Height          =   1230
         ItemData        =   "Form1.frx":049E
         Left            =   -74880
         List            =   "Form1.frx":04A0
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Choose a table"
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdCon 
         Caption         =   "OPEN DATABASE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         ToolTipText     =   "Open Connection"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE/CLEAR"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   10
         ToolTipText     =   "Close Connection"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtCon 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   7320
         TabIndex        =   9
         ToolTipText     =   "Standard Port is 1433"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtCon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   7320
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Your Password on the server"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtCon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   5280
         TabIndex        =   5
         ToolTipText     =   "The name of the Server or if remote the IP No"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtCon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   5280
         TabIndex        =   3
         ToolTipText     =   "Your Username on the server"
         Top             =   720
         Width           =   1815
      End
      Begin RichTextLib.RichTextBox txtTsql 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   20
         Top             =   2280
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":04A2
      End
      Begin MSDataGridLib.DataGrid dbgrid 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   34
         Top             =   2460
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Views in database"
         Height          =   255
         Index           =   5
         Left            =   -71400
         TabIndex        =   31
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Stored Procedures in database"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   29
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Views in database"
         Height          =   255
         Index           =   4
         Left            =   -69600
         TabIndex        =   26
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   " Databases on Server"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblInfo 
         Caption         =   "Database Name: "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   9135
      End
      Begin VB.Label lblInfo 
         Caption         =   "Database Path: "
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   3660
         Width           =   9135
      End
      Begin VB.Label Label3 
         Caption         =   "Columns in Table"
         Height          =   255
         Index           =   2
         Left            =   -72240
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Tables on Database"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblCon 
         Caption         =   "Remote Port"
         Height          =   255
         Index           =   2
         Left            =   7320
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblCon 
         Caption         =   "PassWord"
         Height          =   255
         Index           =   1
         Left            =   7320
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCon 
         Caption         =   "ServerName/IP"
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblCon 
         Caption         =   "UserName"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CDOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Query"
      Filter          =   "Text (*.txt)|*.txt"
      InitDir         =   "app.path"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Query"
      Filter          =   "Text (*.txt)|*.txt"
      InitDir         =   "app.path"
   End
   Begin MSComDlg.CommonDialog CDSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Query"
      Filter          =   "Text (*.txt)|*.txt"
      InitDir         =   "app.path"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuNTSecurity 
         Caption         =   "NT Integrated Security"
         Begin VB.Menu mnuNTSecFalse 
            Caption         =   "False"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuNTSecTrue 
            Caption         =   "True"
         End
      End
      Begin VB.Menu mnuDBType 
         Caption         =   "Database Type"
         Begin VB.Menu mnuAccess 
            Caption         =   "MS Access"
            Begin VB.Menu mnuOpenDB 
               Caption         =   "&Open Access DB"
            End
            Begin VB.Menu mnuScanAccessDB 
               Caption         =   "&Scan Access DB"
            End
         End
         Begin VB.Menu mnuSQLServer 
            Caption         =   "MS SQL-Server"
            Begin VB.Menu mnuSQLLocalLan 
               Caption         =   "Local/Lan"
            End
            Begin VB.Menu mnuSQLRemote 
               Caption         =   "Remote/Internet"
            End
         End
      End
      Begin VB.Menu mnuRSType 
         Caption         =   "Recordset Type"
         Begin VB.Menu mnuRSRead 
            Caption         =   "Read"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRSEdit 
            Caption         =   "Edit"
         End
      End
      Begin VB.Menu mnuAssociate 
         Caption         =   "Associate"
         Begin VB.Menu mnuAssociateAccessDB 
            Caption         =   "Associate Access Databases"
         End
         Begin VB.Menu mnuUnAssociateAccessDB 
            Caption         =   "Unassociate Access Databases"
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "&Contact"
         Begin VB.Menu mnuEmail 
            Caption         =   "&Email Knoton"
         End
         Begin VB.Menu mnuWeb 
            Caption         =   "&Knoton´s Webpage"
         End
      End
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objRS As ADODB.Recordset
Dim objCon As ADODB.Connection
Dim objCom As ADODB.Command
Dim ConString As String         'Variable for the connectionstring
Dim sqlProvider As String       'Variable for the Provider
Dim sqlPws As String            'Variable for the Password
Dim sqlUid As String            'Variable for the UserName
Dim sqlNTSec As String          'Variable for the integrated security (NT-Domain login)
Dim sqlDatSource As String      'Variable for the Name/IP-number of the SQL-Server
Dim sqlDatBase As String        'Variable for the Database choosen
Dim sqlTable As String          'Variable for the Table choosen
Dim sqlColumn As String         'Variable for the Column choosen
Dim sqlNetLibPort As String     'Variable for the Port to use in Remote mode (Internet)
Dim sqlNetLib As String         'Variable to tell ADO that TCP/IP is to be used
Dim AccessDbPath As String      'Variable to tell the path to Access DB
Dim bolDBChoosen As Boolean     'Variable to tell if Database is choosen
Dim bolTableChoosen As Boolean  'Variable to tell if Table is choosen
Dim bolEditReadRS As Boolean    'Variable to tell what kind of recordset (Read or Edit mode)
Dim bolAccess As Boolean        'Variable to tell if Access DB is choosen
Dim bolNoRetRs As Boolean       'Variable to tell to get return value or not
Dim bolCommand As Boolean       'Variable to tell if Command has been used

'*** API to get a list of all drives on the system ***'
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'*** API to get the type of the drive ***'
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


'***Get all Access Databases on your computer***'
Private Sub GetAccessDatabases()
Dim i As Integer
Dim listDrive As Long
Dim strSave As String
strSave = String(255, Chr(0))
Dim strSearch As String
listDrive = GetLogicalDriveStrings(255, strSave) 'get drives
'Screen.MousePointer = vbHourglass
mnuDBType.Enabled = False
For i = 1 To 100 ' split the string with drives and list it
    If Left(strSave, InStr(1, strSave, Chr(0))) = Chr(0) Then Exit For
        strSearch = Left(strSave, InStr(1, strSave, Chr(0)) - 1)
        If GetDriveType(strSearch) = 3 Then '3 = Harddrive, dont scan floppy, cdrom, mapped drive
            Call FindFiles(strSearch, "*.mdb") 'Where and what to search for
        End If
    strSave = Right(strSave, Len(strSave) - InStr(1, strSave, Chr(0)))
Next
'Screen.MousePointer = vbDefault
mnuDBType.Enabled = True
ListFiles
End Sub

'***List all Access Databases  found on your computer***'
Private Sub ListFiles()
Dim i As Integer
lstDatBas.Clear
For i = 1 To noOfFiles
    lstDatBas.AddItem ReturnFileName(i)
Next
End Sub

'***Check the ConnectionState***'
Private Sub CheckConState()
If objRS.State = adStateOpen Then
    StatusBar1.Panels(1).Text = "OPEN"
Else
    StatusBar1.Panels(1).Text = "CLOSED"
End If
End Sub

'***Assemble the ConnectionString and return it***'
Private Function GetConString() As String
sqlDatSource = txtCon(2).Text
If mnuSQLRemote.Checked = True Then sqlNetLibPort = "," & txtCon(3).Text  'If remote server is choosen set Port
sqlUid = txtCon(0).Text
sqlPws = txtCon(1).Text



If AccessDbPath = "" Then 'SQL-Server selected
    If sqlNTSec = "True" Then
        GetConString = "Provider=" & sqlProvider & _
                       ";Integrated Security=SSPI" & _
                       ";Persist Security Info=" & sqlNTSec & _
                       ";Initial Catalog=" & sqlDatBase & _
                       ";Data Source=" & sqlDatSource
    Else
        GetConString = "Provider=" & sqlProvider & _
                        ";Persist Security Info=" & sqlNTSec & _
                        ";User ID=" & sqlUid & _
                        ";Password=" & sqlPws & _
                        ";Initial Catalog=" & sqlDatBase & _
                        ";Data Source=" & sqlDatSource & sqlNetLibPort & _
                        ";Network Library=" & sqlNetLib
    End If

Else 'MS Access Database
    GetConString = "Provider=" & sqlProvider & _
                    ";Data Source=" & AccessDbPath & _
                    ";Persist Security Info=False;User ID=" & sqlUid & _
                    ";Password=" & sqlPws
End If
End Function

'***Get either Readable or Editable Recordset***'
Private Sub GetRecordSet(strSource As String)
If objRS.State = adStateOpen Then objRS.Close 'If the database is open close before getting a new recordset
On Error GoTo Errorhandler 'I take care of eventual errors

Select Case bolEditReadRS 'Tells which kind of recordset to get
    Case False
        With objRS
            .ActiveConnection = objCon
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset 'Move the cursor in any direction and bookmarkable
            .LockType = adLockReadOnly 'Editing is not possible
            .Source = strSource 'What Recordset to get
            .Open
            .MoveFirst
        End With
    Case True
        With objRS
            .ActiveConnection = objCon
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset 'Move the cursor in any direction and bookmarkable
            .LockType = adLockOptimistic 'Editing is possible
            .Source = strSource 'What Recordset to get
            .Open
            .MoveFirst
        End With

End Select
StatusBar1.Panels(4).Text = "Found: " & objRS.RecordCount & " Records"

CheckConState

Errorhandler:
If Err.Number <> 0 Then Call CentralErrhandler("GetRecordSet")
CheckConState
End Sub

'***Create the Connection object***'
Private Sub GetCon()
On Error GoTo Errorhandler
If objCon.State = adStateOpen Then objCon.Close

With objCon
    .ConnectionString = GetConString
    .Open
End With

Errorhandler:
If Err.Number <> 0 Then Call CentralErrhandler("GetRecordSet")
End Sub

Private Sub chkNoRetRS_Click()
If chkNoRetRS.Value = 1 Then
    bolNoRetRs = True
    StatusBar1.Panels(4).Text = "No return value"
End If

If chkNoRetRS.Value = 0 Then bolNoRetRs = False
End Sub

'***Initial Connection. Get aviable Databases on Server or tables on Access DB***'
Private Sub cmdCon_Click()
On Error GoTo ErrHandler

mnuDBType.Enabled = False 'disable the possibility for the user to change dbtype
Set objCon = New ADODB.Connection
Set objRS = New ADODB.Recordset

If AccessDbPath = "" Then 'SQL-Server
    GetCon
    sqlDatBase = "Master"
    GetRecordSet ("exec sp_databases")
    
    If objRS.State = adStateOpen Then
        While Not objRS.EOF = True
            lstDatBas.AddItem objRS.Fields(0)
            objRS.MoveNext
        Wend
    End If
    
    StatusBar1.Panels(4).Text = "Found: " & objRS.RecordCount & " Databases"
Else 'MS Access Database
    sqlDatBase = AccessDbPath
    GetCon
    Set objRS = objCon.OpenSchema(adSchemaTables)
    
    While Not objRS.EOF
        If objRS!TABLE_TYPE = "TABLE" Then lstTables.AddItem objRS!TABLE_NAME
        If objRS!TABLE_TYPE = "VIEW" Then lstView.AddItem objRS!TABLE_NAME
        objRS.MoveNext
    Wend
End If

cmdCon.Enabled = False
cmdTSql.Enabled = True
cmdClose.Enabled = True
mnuDBType.Enabled = False
If bolAccess = True Then SSTab1.Tab = 1
CheckConState

ErrHandler:
If Err.Number <> 0 Then
    Call CentralErrhandler("cmdCon_Click")
    mnuDBType.Enabled = True
    cmdClose.Enabled = True
End If
End Sub

'***Delete Current Record***'
Private Sub cmdDel_Click()
On Error GoTo ErrHandler

If MsgBox("Are you sure you want to delete " & objRS.Fields(0) & " ?", vbOKCancel, "Delete") = vbOK Then
    objRS.Delete adAffectCurrent
    objRS.Update
End If
lstTables_Click

ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("cmdDel_Click")
CheckConState
End Sub

'***Open a saved query***'
Private Sub cmdOpen_Click()
CDOpen.InitDir = App.path
CDOpen.ShowOpen
txtTsql.FileName = CDOpen.FileName
End Sub

'***Clear current query***'
Private Sub cmdClearQuery_Click()
txtTsql.Text = ""
End Sub

'***Save current query***'
Private Sub cmdSave_Click()
CDSave.InitDir = App.path
CDSave.ShowSave
txtTsql.SaveFile CDSave.FileName
End Sub

'***Get the Custom SQL Recordset build by the user***'
Private Sub cmdTSql_Click()
Dim strSQL As String
strSQL = txtTsql.Text
On Error GoTo ErrHandler
If bolDBChoosen = True Then 'only if a database is choosen

Select Case bolNoRetRs
    Case False
        GetRecordSet (strSQL)
        Set dbgrid.DataSource = objRS 'Map the Recordset to the DataBaseGrid
        dbgrid.Refresh
        
        If objRS.State = adStateOpen Then SSTab1.Tab = 1 'if everything goes well show result
    Case True
        Set objCom = New ADODB.Command
        With objCom
            .ActiveConnection = objCon
            .CommandType = adCmdText
            .CommandText = strSQL
            .Execute
        End With
        dbgrid.ClearFields
        lstDatBas_Click
        SSTab1.Tab = 2
End Select

Else
    MsgBox "You must choose a database !"
End If

ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("cmdTSql_Click")
If bolNoRetRs Then CheckConState
txtTsql.SetFocus
End Sub

'***Update what has been edited in the current Recordset***'
Private Sub cmdUpdate_Click()
On Error GoTo ErrHandler

objRS.Update
dbgrid.Refresh

ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("cmdUpdate_Click")
CheckConState
End Sub

'***Close and cleanup***'
Private Sub cmdClose_Click()
Dim i As Integer
If objRS.State = adStateOpen Then
    objRS.Close
    objCon.Close
End If

    Set objRS = Nothing
    Set objCon = Nothing
    lstTables.Clear
    lstDatBas.Clear
    lstcol.Clear
    lstProc.Clear
    lstView.Clear
    lstSqlView.Clear
    lstDatBas.ToolTipText = ""
    lstDatBas.Enabled = True
    cmdCon.Enabled = False
    cmdTSql.Enabled = False
    cmdUpdate.Enabled = False
    cmdDel.Enabled = False
    cmdClose.Enabled = False
    txtCon(0).Text = ""
    txtCon(1).Text = ""
    txtCon(2).Text = ""
    txtCon(3).Text = ""
    txtCon(3).Enabled = False
    txtCon(0).Enabled = True
    txtCon(1).Enabled = True
    bolDBChoosen = False
    bolTableChoosen = False
    bolAccess = False
    bolEditReadRS = False
    StatusBar1.Panels(1).Text = "CLOSED"
    StatusBar1.Panels(2).Text = "READ"
    StatusBar1.Panels(3).Text = ""
    StatusBar1.Panels(4).Text = ""
    StatusBar1.Panels(5).Text = ""
    sqlNetLibPort = ""
    sqlDatBase = ""
    sqlUid = ""
    sqlPws = ""
    sqlNTSec = "False"
    sqlProvider = ""
    AccessDbPath = ""
    txtTsql.Text = ""
    lblInfo(0).Caption = "Database Name: "
    lblInfo(1).Caption = "Database Path: "
    mnuDBType.Enabled = True
    mnuRSRead.Checked = True
    mnuSQLLocalLan.Checked = False
    mnuSQLRemote.Checked = False
    mnuOpenDB.Checked = False
    mnuScanAccessDB.Checked = False
    mnuNTSecFalse.Checked = True
    mnuNTSecTrue.Checked = False
End Sub

Private Sub Form_Load()
sqlNTSec = "False"

If Command <> "" Then 'Access DB choosen from explorer
    bolCommand = True
    mnuOpenDB_Click
    cmdCon_Click
End If
End Sub

Private Sub Form_Resize()
If frmDB.WindowState = 0 Then
    frmDB.Height = 6255
    frmDB.Width = 9510
End If
If Not frmDB.WindowState = 1 Then
SSTab1.Width = frmDB.Width - 100
SSTab1.Height = frmDB.Height - 1100
dbgrid.Width = SSTab1.Width - 250
lstTables.Width = (SSTab1.Width - 500) / 3
lstcol.Left = lstTables.Left + lstTables.Width + 100
lstcol.Width = lstTables.Width
lstView.Left = lstcol.Left + lstcol.Width + 100
lstView.Width = lstTables.Width
Label3(1).Left = lstTables.Left
Label3(2).Left = lstcol.Left
Label3(4).Left = lstView.Left
dbgrid.Height = SSTab1.Height - dbgrid.Top - 150
txtTsql.Width = SSTab1.Width - 250
txtTsql.Height = SSTab1.Height - txtTsql.Top - 150
lblInfo(0).Width = SSTab1.Width - 250
lblInfo(1).Width = SSTab1.Width - 250
End If
End Sub

'***Set the conditions on the Recordset to get***'
Private Sub lstCol_Click()
On Error GoTo ErrHandler
If lstcol.ListIndex <> -1 Then
    lstView.ListIndex = -1
    StatusBar1.Panels(1).Text = "SEARCHING"
    Dim i As Integer
    sqlColumn = lstcol.List(lstcol.ListIndex)
    GetRecordSet ("select [" & sqlColumn & "] from " & "[" & sqlTable & "]")
    Set dbgrid.DataSource = objRS
    dbgrid.Refresh
End If
ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("lstCol_Click")
CheckConState

End Sub

'***Choose the DataBase and set the conditions on the Recordset to get***'
Private Sub lstDatBas_Click()
Dim i As Integer
Dim strProc As String
On Error GoTo ErrHandler

bolTableChoosen = False
bolDBChoosen = True
lstTables.Clear
lstcol.Clear
lstProc.Clear
lstView.Clear
lstSqlView.Clear
Select Case bolAccess
    Case True
        AccessDbPath = ReturnPath(lstDatBas.ListIndex + 1) & ReturnFileName(lstDatBas.ListIndex + 1)
        lblInfo(0).Caption = "Database Name: " & ReturnFileName(lstDatBas.ListIndex + 1)
        lblInfo(1).Caption = "Database Path: " & ReturnPath(lstDatBas.ListIndex + 1)
        lstDatBas.ToolTipText = AccessDbPath
        StatusBar1.Panels(3).Text = "MS Access Database"
        StatusBar1.Panels(5).Text = ReturnFileName(lstDatBas.ListIndex + 1)
        cmdCon_Click
    Case False
        sqlDatBase = lstDatBas.List(lstDatBas.ListIndex)
        GetCon
        
        If objRS.State = adStateOpen Then objRS.Close
        Set objRS = objCon.OpenSchema(adSchemaTables)
        
        While Not objRS.EOF
            If objRS!TABLE_TYPE = "TABLE" Then lstTables.AddItem objRS!TABLE_NAME
            If objRS!TABLE_TYPE = "VIEW" Then
                lstView.AddItem objRS!TABLE_NAME
                lstSqlView.AddItem objRS!TABLE_NAME
            End If
            objRS.MoveNext
        Wend
                
        If objRS.State = adStateOpen Then objRS.Close
        Set objRS = objCon.OpenSchema(adSchemaProcedures)
        
        While Not objRS.EOF
            strProc = objRS!PROCEDURE_NAME
            For i = 1 To Len(strProc)
                If Mid(strProc, i, 1) = ";" Then
                    strProc = Mid(strProc, 1, i - 1)
                    Exit For
                End If
            Next
            lstProc.AddItem strProc
            objRS.MoveNext
        Wend

        StatusBar1.Panels(5).Text = sqlDatBase
        CheckConState
        SSTab1.Tab = 1
End Select
StatusBar1.Panels(4).Text = "Found: " & lstTables.ListCount + lstView.ListCount + lstProc.ListCount & " Objects"

ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("lstDatBas_Click")
End Sub

Private Sub lstProc_Click()
On Error GoTo ErrHandler
Dim strHelptext As String
GetRecordSet ("Exec sp_helptext " & "[" & lstProc.List(lstProc.ListIndex) & "]")
strHelptext = ""
txtTsql.Text = ""
While Not objRS.EOF
    strHelptext = strHelptext & objRS(0)
    objRS.MoveNext
Wend

txtTsql.Text = RTrim(strHelptext)

ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("lstProc_Click")
CheckConState

End Sub

Private Sub lstSqlView_Click()
Dim strHelptext As String
On Error GoTo ErrHandler

GetRecordSet ("Exec sp_helptext " & "[" & lstSqlView.List(lstSqlView.ListIndex) & "]")
strHelptext = ""
txtTsql.Text = ""
While Not objRS.EOF
    strHelptext = strHelptext & objRS(0)
    objRS.MoveNext
Wend

txtTsql.Text = RTrim(strHelptext)
ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("lstSqlView_Click")
CheckConState
End Sub

'***Choose the Table and set the conditions on the Recordset to get***'
Private Sub lstTables_Click()
Dim i As Integer
Dim Column As Field

On Error GoTo ErrHandler
If lstTables.ListIndex <> -1 Then
    StatusBar1.Panels(1).Text = "SEARCHING"
    sqlTable = lstTables.List(lstTables.ListIndex)
    lstcol.Clear
    GetRecordSet ("Select * from " & "[" & sqlTable & "]")
    
    If objRS.State = adStateOpen Then
        For Each Column In objRS.Fields
            lstcol.AddItem Column.Name
        Next
    Set dbgrid.DataSource = objRS
    dbgrid.Refresh
    End If

    bolTableChoosen = True
    lstcol.ListIndex = -1
    lstView.ListIndex = -1
End If
ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("lstTables_Click")
CheckConState
End Sub

Private Sub lstView_Click()
On Error GoTo ErrHandler
If lstView.ListIndex <> -1 Then
    lstTables.ListIndex = -1
    lstcol.ListIndex = -1
    bolTableChoosen = True
    GetRecordSet ("Select * from " & "[" & lstView.List(lstView.ListIndex) & "]")
    Set dbgrid.DataSource = objRS
    dbgrid.Refresh
End If
ErrHandler:
If Err.Number <> 0 Then Call CentralErrhandler("lstView_Click")
CheckConState
End Sub

Private Sub mnuAssociateAccessDB_Click()
Associate "DatabaseViewerEditor", ".mdb", "Access Database", App.path & "\MSGBOX02.ICO"
End Sub

'***Send Mail to the developer of this application****
Private Sub mnuEmail_Click()
WebEmailOpen ("mailto:knoton@hotmail.com?subject=Report Database EditorViewer")
End Sub

'***Unload this application***'
Private Sub mnuExit_Click()
Unload Me
End Sub

'***Open the Help File***'
Private Sub mnuHelp_Click()
WebEmailOpen (App.path & "/help.doc")
End Sub

'***Select to open MS Access database yourself***'
Private Sub mnuOpenDB_Click()
Dim i As Integer
Dim strTemp As String
cmdCon.Enabled = True
lstDatBas.Clear
sqlProvider = "Microsoft.Jet.OLEDB.4.0"
sqlNetLib = ""
sqlNetLibPort = ""
bolAccess = True
Label3(0).Caption = "MS Access Databases"
StatusBar1.Panels(3).Text = "MS Access Database"

bolTableChoosen = False
bolDBChoosen = True
lstTables.Clear
lstcol.Clear
If bolCommand = True Then 'Access DB choosen from explorer
    AccessDbPath = Command
    bolCommand = False
Else
    CDADB.ShowOpen
    AccessDbPath = CDADB.FileName
End If
mnuSQLLocalLan.Checked = False
mnuSQLRemote.Checked = False
mnuOpenDB.Checked = True
mnuScanAccessDB.Checked = False
chkNoRetRS.Enabled = False
    
    'Split the filename and path from the dialog filename
    For i = 1 To Len(AccessDbPath) - 1
        If Mid(AccessDbPath, i, 1) = "\" Then
        strTemp = Mid(AccessDbPath, 1, i)
        End If
    Next
StatusBar1.Panels(5).Text = Mid(AccessDbPath, Len(strTemp) + 1)
lblInfo(1).Caption = lblInfo(1).Caption & strTemp
lblInfo(0).Caption = lblInfo(0).Caption & Mid(AccessDbPath, Len(strTemp) + 1)
End Sub

'***Scan for all MS Access databases***'
Private Sub mnuScanAccessDB_Click()
cmdCon.Enabled = True
lstDatBas.Clear
sqlProvider = "Microsoft.Jet.OLEDB.4.0"
sqlNetLib = ""
sqlNetLibPort = ""
bolAccess = True
Label3(0).Caption = "MS Access Databases"
StatusBar1.Panels(4).Text = "Searching for Databases please wait !"
GetAccessDatabases
StatusBar1.Panels(4).Text = "Found: " & noOfFiles & " Databases"
mnuSQLLocalLan.Checked = False
mnuSQLRemote.Checked = False
mnuOpenDB.Checked = False
mnuScanAccessDB.Checked = True
chkNoRetRS.Enabled = False
End Sub

'***Select editable recordset***'
Private Sub mnuRSEdit_Click()
bolEditReadRS = True
cmdUpdate.Enabled = True
cmdDel.Enabled = True
StatusBar1.Panels(2).Text = "EDIT"
mnuRSEdit.Checked = True
mnuRSRead.Checked = False
If bolTableChoosen = True Then lstTables_Click 'Get the New Edit Recordset
End Sub

'***Select readonly recordset***'
Private Sub mnuRSRead_Click()
bolEditReadRS = False
cmdUpdate.Enabled = False
cmdDel.Enabled = False
StatusBar1.Panels(2).Text = "READ"
mnuRSEdit.Checked = False
mnuRSRead.Checked = True
If bolTableChoosen = True Then lstTables_Click 'Get the New Read Recordset
End Sub

'***Select Local/Lan SQL-Server***'
Private Sub mnuSQLLocalLan_Click()
cmdCon.Enabled = True
lstDatBas.Clear
sqlProvider = "SQLOLEDB.1"
sqlNetLib = ""
sqlNetLibPort = ""
Label3(0).Caption = "Databases on server"
StatusBar1.Panels(3).Text = "Local/Lan SQL-Server"
bolAccess = False
mnuSQLLocalLan.Checked = True
mnuSQLRemote.Checked = False
mnuOpenDB.Checked = False
mnuScanAccessDB.Checked = False
chkNoRetRS.Enabled = True
chkNoRetRS.Value = 0
End Sub

'***Select Remote SQL-Server***'
Private Sub mnuSQLRemote_Click()
cmdCon.Enabled = True
lstDatBas.Clear
sqlProvider = "SQLOLEDB.1"
sqlNetLib = "DBMSSOCN"
txtCon(3).Enabled = True
txtCon(3).Text = "1433" 'default port for remote SQL-Server
Label3(0).Caption = "Databases on server"
StatusBar1.Panels(3).Text = "Remote SQL-Server"
bolAccess = False
mnuSQLLocalLan.Checked = False
mnuSQLRemote.Checked = True
mnuOpenDB.Checked = False
mnuScanAccessDB.Checked = False
chkNoRetRS.Enabled = True
chkNoRetRS.Value = 0
End Sub

Private Sub mnuUnAssociateAccessDB_Click()
RemoveAssociate "DatabaseViewerEditor", ".mdb"
End Sub

'***Go to the website of the developer of this application***'
Private Sub mnuWeb_Click()
WebEmailOpen ("http://www.knoton.dns2go.com")
End Sub

'***NT Security is not used (default)***'
Private Sub mnuNTSecFalse_Click()
mnuNTSecFalse.Checked = True
mnuNTSecTrue.Checked = False
sqlNTSec = "False"
txtCon(0).Enabled = True
txtCon(1).Enabled = True
End Sub

'***NT Security is used***'
Private Sub mnuNTSecTrue_Click()
mnuNTSecFalse.Checked = False
mnuNTSecTrue.Checked = True
sqlNTSec = "True"
txtCon(0).Enabled = False
txtCon(1).Enabled = False
sqlPws = ""
sqlUid = ""
End Sub

