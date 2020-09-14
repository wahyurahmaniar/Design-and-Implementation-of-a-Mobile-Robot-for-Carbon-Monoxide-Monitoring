VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Antarmuka Data CO"
   ClientHeight    =   10335
   ClientLeft      =   1935
   ClientTop       =   990
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   13710
   Begin MSCommLib.MSComm serial 
      Left            =   9960
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Batas PPM CO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   10080
      TabIndex        =   11
      Top             =   960
      Width           =   3495
      Begin VB.CommandButton cmdKirim 
         Caption         =   "Kirim Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtMax 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   15
         Text            =   "50"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtMin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   13
         Text            =   "25"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   10080
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tegangan dan CO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   10080
      TabIndex        =   1
      Top             =   3480
      Width           =   3495
      Begin VB.TextBox txtTeg 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   480
         MaxLength       =   6
         TabIndex        =   17
         Text            =   "0"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtCO 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   480
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "0"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderWidth     =   5
         Height          =   735
         Left            =   360
         Top             =   480
         Width           =   2775
      End
      Begin VB.Shape off3 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   615
      End
      Begin VB.Shape alarm3 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   615
      End
      Begin VB.Shape off2 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   615
      End
      Begin VB.Shape alarm2 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   615
      End
      Begin VB.Shape off1 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   480
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   615
      End
      Begin VB.Shape alarm1 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   480
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   615
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderWidth     =   5
         Height          =   735
         Left            =   360
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   480
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   615
      End
      Begin VB.Shape Shape12 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   615
      End
      Begin VB.Shape Shape14 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5400
      Top             =   720
   End
   Begin VB.Timer Timer3 
      Interval        =   200
      Left            =   5880
      Top             =   720
   End
   Begin VB.CommandButton keluar 
      Caption         =   "K E L U A R"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16560
      TabIndex        =   0
      Top             =   10200
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   16325
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Tabel"
      TabPicture(0)   =   "Form1.frx":1F7A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Adodc2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Data Grafik"
      TabPicture(1)   =   "Form1.frx":1F96
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text5"
      Tab(1).Control(1)=   "index2"
      Tab(1).Control(2)=   "HScroll2"
      Tab(1).Control(3)=   "HScroll1"
      Tab(1).Control(4)=   "index"
      Tab(1).Control(5)=   "Text2"
      Tab(1).Control(6)=   "MSChart1"
      Tab(1).Control(7)=   "MSChart2"
      Tab(1).ControlCount=   8
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Text            =   "INDEX"
         Top             =   8280
         Width           =   735
      End
      Begin VB.TextBox index2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   21
         Top             =   8280
         Width           =   735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   -74880
         Max             =   1000
         TabIndex        =   20
         Top             =   8880
         Value           =   100
         Width           =   9495
      End
      Begin VB.TextBox Text4 
         Height          =   435
         Left            =   5040
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   8520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   435
         Left            =   7080
         TabIndex        =   18
         Top             =   8520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   -74880
         Max             =   1000
         TabIndex        =   5
         Top             =   4560
         Value           =   100
         Width           =   9495
      End
      Begin VB.TextBox index 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   4
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   3
         Text            =   "INDEX"
         Top             =   3960
         Width           =   735
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "Form1.frx":1FB2
         TabIndex        =   6
         Top             =   480
         Width           =   9495
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   8655
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   15266
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tabel Data CO"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   7680
         Top             =   0
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\2012\Robot CO\antarmuka\data_co.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\2012\Robot CO\antarmuka\data_co.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "dataCO"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   3855
         Left            =   -74880
         OleObjectBlob   =   "Form1.frx":3A47
         TabIndex        =   23
         Top             =   4920
         Width           =   9495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11640
      Top             =   720
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\2012\Robot CO\antarmuka\data_co.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\2012\Robot CO\antarmuka\data_co.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "data"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image stop1 
      Height          =   975
      Index           =   1
      Left            =   11280
      Picture         =   "Form1.frx":54E5
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image atas 
      Height          =   975
      Index           =   1
      Left            =   11280
      Picture         =   "Form1.frx":10F32
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Image bawah 
      Height          =   975
      Index           =   1
      Left            =   11280
      Picture         =   "Form1.frx":1D8BC
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Image kiri 
      Height          =   975
      Index           =   1
      Left            =   10080
      Picture         =   "Form1.frx":2A391
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image kanan 
      Height          =   975
      Index           =   1
      Left            =   12480
      Picture         =   "Form1.frx":37412
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image kanan 
      Height          =   975
      Index           =   0
      Left            =   12480
      Picture         =   "Form1.frx":44411
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image kiri 
      Height          =   975
      Index           =   0
      Left            =   10080
      Picture         =   "Form1.frx":4505D
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image bawah 
      Height          =   975
      Index           =   0
      Left            =   11280
      Picture         =   "Form1.frx":45E0B
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Image atas 
      Height          =   975
      Index           =   0
      Left            =   11280
      Picture         =   "Form1.frx":46971
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Image stop1 
      Height          =   975
      Index           =   0
      Left            =   11280
      Picture         =   "Form1.frx":47537
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label judul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                Antarmuka Kendali Robot dan Pemantau Data CO Jarak Jauh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   180
      Width           =   13215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   13455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   10695
      Left            =   0
      Top             =   -240
      Width           =   13935
   End
   Begin VB.Menu mMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mSimpan 
         Caption         =   "Simpan Tabel"
      End
      Begin VB.Menu mHapus 
         Caption         =   "Hapus Tabel"
      End
      Begin VB.Menu mCetak 
         Caption         =   "Cetak Grafik"
      End
      Begin VB.Menu x1 
         Caption         =   "-"
      End
      Begin VB.Menu mOut 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mKoneksi 
      Caption         =   "&Koneksi"
      Begin VB.Menu mHubung 
         Caption         =   "Hubung"
      End
      Begin VB.Menu mPutus 
         Caption         =   "Putus"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection, conn2 As Connection
Dim rs As Recordset, rs2 As Recordset
Dim jumlah As Integer, jumlah2 As Integer
Dim selisih As Integer, selisih2 As Integer
Dim waktu As Double
Dim buffer As String
Dim itung As Integer

'Sub program untuk menghubungkan ADODC dengan database dan menampilkan pada tabel (datagrid)
Private Sub Koneksi()
    Set conn = New Connection
    conn.Open "PROVIDER=MSDataShape; Data PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data_co.mdb;"
    Set rs = New Recordset
    rs.Open "select * from data", conn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = rs.DataSource
    DataGrid1.Enabled = True
    DataGrid1.Columns(0).Width = "1000"
    DataGrid1.Columns(1).Width = "1700"
    DataGrid1.Columns(2).Width = "1700"
    DataGrid1.Columns(3).Width = "1500"
    DataGrid1.Columns(4).Width = "1500"
End Sub

'Sub program untuk menghubungkan ADODC dengan database
Private Sub Koneksi2()
    Set conn2 = New Connection
    conn2.Open "PROVIDER=MSDataShape; Data PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data_co.mdb;"
    Set rs2 = New Recordset
    rs2.Open "select * from dataCO", conn2, adOpenStatic, adLockOptimistic
End Sub

Private Sub atas_Click(index As Integer)
serial.Output = "A" & "A" & Chr$(13)
'mengirimkan karakter ke port serial
End Sub

Private Sub atas_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
atas(0).Visible = False
atas(1).Visible = True
StatusBar1.Panels(4).Text = "DEPAN"
End Sub

Private Sub bawah_Click(index As Integer)
serial.Output = "B" & "B" & Chr$(13)
End Sub

Private Sub bawah_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
bawah(0).Visible = False
bawah(1).Visible = True
StatusBar1.Panels(4).Text = "BELAKANG"
End Sub

Private Sub cmdKirim_Click()
serial.Output = "Y" & "Y" & " " & txtMin & " " & txtMax & Chr$(13)
End Sub

Private Sub Form_Load()
jumlah = 0
jumlah2 = 0
itung = 0

Koneksi
Koneksi2
Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = True

mHubung.Enabled = True
mPutus.Enabled = False
mOut.Enabled = True
mSimpan.Enabled = True
mHapus.Enabled = True
mCetak.Enabled = True
cmdKirim.Enabled = False

alarm1.Visible = False
alarm2.Visible = False
alarm3.Visible = False
off1.Visible = True
off2.Visible = True
off3.Visible = True

atas(1).Visible = False
bawah(1).Visible = False
kanan(1).Visible = False
kiri(1).Visible = False
stop1(1).Visible = False

atas(0).Enabled = False
bawah(0).Enabled = False
kanan(0).Enabled = False
kiri(0).Enabled = False
stop1(0).Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atas(1).Visible = False
bawah(1).Visible = False
kanan(1).Visible = False
kiri(1).Visible = False
stop1(1).Visible = False

atas(0).Visible = True
bawah(0).Visible = True
kanan(0).Visible = True
kiri(0).Visible = True
stop1(0).Visible = True
End Sub

Private Sub kanan_Click(index As Integer)
serial.Output = "D" & "D" & Chr$(13)
End Sub

Private Sub kanan_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
kanan(0).Visible = False
kanan(1).Visible = True
StatusBar1.Panels(4).Text = "KANAN"
End Sub

Private Sub kiri_Click(index As Integer)
serial.Output = "C" & "C" & Chr$(13)
End Sub

Private Sub kiri_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
kiri(0).Visible = False
kiri(1).Visible = True
StatusBar1.Panels(4).Text = "KIRI"
End Sub

Private Sub mCetak_Click()
msg = MsgBox("Cetak grafik?", vbYesNo, "Konfirmasi")
    If msg = vbNo Then Exit Sub
MSChart1.EditCopy
Printer.Print " "
Printer.PaintPicture Clipboard.GetData(), 0, 0
Printer.EndDoc
End Sub

Private Sub mHubung_Click()
Form2.Show
End Sub

Private Sub mout_Click()
Unload Form1
End
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    If serial.PortOpen = True Then
        serial.PortOpen = False
    End If
    Unload Form2
End Sub

Private Sub mHapus_Click()
msg = MsgBox("Hapus Data Tabel?", vbYesNo, "Konfirmasi")
    If msg = vbNo Then Exit Sub
    hapus_grafik
        If rs.RecordCount <> 0 Then rs.MoveFirst
        'jika database tidak kosong pergi ke baris teratas
    While rs.EOF = False 'hapus sampai akhir baris
        rs.Delete
        rs.MoveNext
    Wend
index.Text = 0
End Sub


Private Sub mPutus_Click()
If serial.PortOpen = True Then
serial.PortOpen = False
End If

Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = True
itung = 0

StatusBar1.Panels(1).Text = "Terputus..."

mHubung.Enabled = True
mPutus.Enabled = False
mOut.Enabled = True
mSimpan.Enabled = True
mHapus.Enabled = True
mCetak.Enabled = True
cmdKirim.Enabled = False

alarm1.Visible = False
alarm2.Visible = False
alarm3.Visible = False
off1.Visible = True
off2.Visible = True
off3.Visible = True

End Sub

'Fungsi untuk menyimpan pada excel
Private Sub mSimpan_Click()
Dim excel_app As Object
Dim excel_sheet As Object
Dim excel_book As Object
Dim row As Long
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim statement As String
Dim col As Integer
    
    msg = MsgBox("Simpan data tabel?", vbYesNo, "Konfirmasi")
    If msg = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    excel_app.Workbooks.Add
    Set excel_sheet = excel_app.Worksheets(1)
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & App.Path & "\data_co.mdb;" & _
        "Persist Security Info=False"
    conn.Open
        Set rs = conn.Execute("SELECT * FROM dataCO", , adCmdText)
    For col = 0 To 4
        excel_sheet.Cells(1, col + 1) = rs.Fields(col).Name
    Next col
    row = 2
    Do While Not rs.EOF
        For col = 0 To rs.Fields.Count - 1
            excel_sheet.Cells(row, col + 1) = _
                rs.Fields(col).Value
        Next col
        row = row + 1
        rs.MoveNext
    Loop
    excel_sheet.Range( _
        excel_sheet.Cells(1, 1), _
            excel_sheet.Cells(2, _
            rs.Fields.Count)).Columns.AutoFit
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    excel_sheet.Rows(1).Font.Bold = True
    excel_sheet.Rows(2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    excel_app.Workbooks(1).Close True
    excel_app.Quit
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub stop1_Click(index As Integer)
serial.Output = "E" & "E" & Chr$(13)
End Sub

Private Sub stop1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
stop1(0).Visible = False
stop1(1).Visible = True
StatusBar1.Panels(4).Text = "STOP"
End Sub

'Sub program untuk menampilkan indikator hijau, kuning, atau merah
Private Sub alarm()
If CDbl(txtCO) <= CDbl(txtMin) Then
    alarm1.Visible = True
    off1.Visible = False
    alarm2.Visible = False
    off2.Visible = True
    alarm3.Visible = False
    off3.Visible = True
ElseIf CDbl(txtCO) > CDbl(txtMin) And CDbl(txtCO) < CDbl(txtMax) Then
    alarm2.Visible = True
    off2.Visible = False
    alarm1.Visible = False
    off1.Visible = True
    alarm3.Visible = False
    off3.Visible = True
ElseIf CDbl(txtCO) >= CDbl(txtMax) Then
    alarm3.Visible = True
    off3.Visible = False
    alarm2.Visible = False
    off2.Visible = True
    alarm1.Visible = False
    off1.Visible = True
End If

End Sub

'Timer mengecek dan menerima data serial yg masuk
Private Sub Timer1_Timer()
Dim pisah() As String
Dim awal() As String
Dim i As Integer

If serial.CommEvent = comEvReceive Then 'jika ada data serial masuk
buffer = serial.Input 'simpan pada variabel buffer
    On Error Resume Next
    If buffer <> "" Then 'jika buffer tdk kosong
    pisah = Split(buffer, vbCrLf) 'pisahkan data yg masuk berdasarkan enter
    txtTeg = CDbl(pisah(0)) 'data pada enter pertama = data tegangan
    txtCO = CDbl(pisah(1)) 'data pada enter kedua = data PPM Co
    data_tambahan 'simpan pada database
    data_tambahan2 'simpan pada database
    grafik 'tampilkan grafik
    alarm 'tampilkan indikator
    itung = itung + 1
    If itung Mod 100 = 0 Then 'klo data yg masuk udah 100
    Timer1.Enabled = False
    mPutus_Click
    mHapus_Click 'hapus data
        If index.Text = "0" Then
        Timer1.Enabled = True
        mHubung_Click
        End If
    End If
    Else
    buffer = ""
    End If
End If
End Sub

Private Sub Timer2_Timer()
StatusBar1.Panels(3).Text = "Tanggal: " & Format(Date, "dd/mm/yy") & " Jam: " & Format(Time, "hh:mm:ss")
End Sub

'Untuk menggeser tulisan judul
Private Sub Timer3_Timer()
Timer3.Interval = 500
On Error Resume Next
    a = Left(judul, 1)
    b = Len(judul)
    C = Right(judul, b - 1)
    judul = C + a
End Sub

Private Sub data_tambahan()
nomor_data
rs.AddNew
rs!No = jumlah + 1
rs!Tanggal = Format(Date, "dd/mm/yy")
rs!waktu = Format(Time, "hh:mm:ss")
rs!Tegangan = CDbl(txtTeg)
rs!PPM = CDbl(txtCO)
rs.Update
End Sub

Private Sub nomor_data()
    If rs.RecordCount <> 0 Then
        rs.MoveLast
        jumlah = Val(rs!No)
    Else
    jumlah = 0
    End If
End Sub

Private Sub data_tambahan2()
nomor_data2
rs2.AddNew
rs2!No = jumlah + 1
rs2!Tanggal = Format(Date, "dd/mm/yy")
rs2!waktu = Format(Time, "hh:mm:ss")
rs2!Tegangan = CDbl(txtTeg)
rs2!PPM = CDbl(txtCO)
rs2.Update
End Sub

Private Sub nomor_data2()
    If rs2.RecordCount <> 0 Then
        rs2.MoveLast
        jumlah2 = Val(rs2!No)
    Else
    jumlah2 = 0
    End If
End Sub

Private Sub grafik()
  If rs.RecordCount = 0 Then Exit Sub
    If (HScroll1.Max - HScroll1.Value) < rs.RecordCount Then
        ReDim nilai(1 To (HScroll1.Max - HScroll1.Value), 1)
           selisih = rs.RecordCount - (HScroll1.Max - HScroll1.Value)
        rs.MoveFirst
        While selisih > 0
            selisih = selisih - 1
            rs.MoveNext
        Wend
        For X = 1 To rs.RecordCount
            nilai(X, 1) = rs!PPM
            rs.MoveNext
         Next X
        
        MSChart1.ChartData = nilai
        selisih = rs.RecordCount - (HScroll1.Max - HScroll1.Value)
        rs.MoveFirst
        While selisih > 0
            selisih = selisih - 1
            rs.MoveNext
        Wend
        For X = 1 To (HScroll1.Max - HScroll1.Value)
            MSChart1.row = X
            MSChart1.RowLabel = rs!No
            rs.MoveNext
        Next X
    End If
    
     If (HScroll2.Max - HScroll2.Value) < rs.RecordCount Then
        ReDim nilai2(1 To (HScroll2.Max - HScroll2.Value), 1)
           selisih2 = rs.RecordCount - (HScroll2.Max - HScroll2.Value)
        rs.MoveFirst
        While selisih2 > 0
            selisih2 = selisih2 - 1
            rs.MoveNext
        Wend
        For X2 = 1 To rs.RecordCount
            nilai2(X, 1) = rs!Tegangan
            rs.MoveNext
         Next X2
        
        MSChart2.ChartData = nilai
        selisih2 = rs.RecordCount - (HScroll2.Max - HScroll2.Value)
        rs.MoveFirst
        While selisih2 > 0
            selisih2 = selisih2 - 1
            rs.MoveNext
        Wend
        For X2 = 1 To (HScroll2.Max - HScroll2.Value)
            MSChart2.row = X2
            MSChart2.RowLabel = rs!No
            rs.MoveNext
        Next X2
    Else
        grafik2
    End If
End Sub

Private Sub grafik2()
ReDim nilai(1 To rs.RecordCount, 1 To 2)

ReDim nilai2(1 To rs.RecordCount, 1 To 2)
rs.MoveFirst

        For X = 1 To rs.RecordCount
            nilai(X, 1) = rs!PPM
            nilai2(X, 1) = rs!Tegangan
        rs.MoveNext
        Next X
    
        MSChart1.ChartData = nilai
        MSChart2.ChartData = nilai2
        rs.MoveFirst
    
        For X = 1 To rs.RecordCount
        MSChart1.row = X
        MSChart1.RowLabel = rs!No
        MSChart2.row = X
        MSChart2.RowLabel = rs!No
        rs.MoveNext
        Next X
        
index.Text = rs.RecordCount
HScroll1.Value = index.Text
index2.Text = rs.RecordCount
HScroll2.Value = index2.Text
End Sub

Private Sub hapus_grafik()
If rs.RecordCount = 0 Then Exit Sub
ReDim nilai(1, 1)
ReDim nilai2(1, 1)
rs.MoveFirst
       
    nilai(1, 1) = 0
    nilai2(1, 1) = 0
    rs.MoveNext
           
    MSChart1.ChartData = nilai
    MSChart2.ChartData = nilai2
    rs.MoveFirst
    
    MSChart1.row = 1
    MSChart1.RowLabel = 0
    MSChart2.row = 1
    MSChart2.RowLabel = 0
    rs.MoveNext
End Sub


