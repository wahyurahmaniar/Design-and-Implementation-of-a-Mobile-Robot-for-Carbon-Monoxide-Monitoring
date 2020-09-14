VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan"
   ClientHeight    =   2670
   ClientLeft      =   2625
   ClientTop       =   2085
   ClientWidth     =   3375
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3375
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pengaturan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdHubung 
         Caption         =   "Hubung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdPutus 
         Caption         =   "Putus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   4440
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   4
         Text            =   "COM1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   3
         Text            =   "9600"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton batal 
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton ok 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "COM PORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "BAUD RATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub batal_Click()
Unload Form2
End Sub

Private Sub Form_Load()
With Combo1
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .AddItem "COM5"
        .AddItem "COM6"
        .AddItem "COM7"
        .AddItem "COM8"
        .AddItem "COM9"
        .AddItem "COM10"
End With
With Combo2
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "19200"
        .AddItem "38400"
        .AddItem "56600"
    End With
End Sub


Private Sub ok_Click()
Dim Port As Integer
    On Error GoTo errcode
    Select Case Combo1.ListIndex
    Case -1
        Port = 1
    Case 0
        Port = 1
    Case 1
        Port = 2
    Case 2
        Port = 3
    Case 3
        Port = 4
    Case 4
        Port = 5
    Case 5
        Port = 6
    Case 6
        Port = 7
    Case 7
        Port = 8
    Case 8
        Port = 9
    Case 9
        Port = 10
    End Select
    If Form1.serial.PortOpen = False Then
        Form1.serial.CommPort = Port
        Form1.serial.RThreshold = 1
        Form1.serial.InputLen = 40
        Form1.serial.Settings = Combo2.List(Combo2.ListIndex) & ",N,8,1"
        Form1.serial.PortOpen = True
    End If
    Form1.StatusBar1.Panels(2).Text = "COM" & Port & " " & Combo2.List(Combo2.ListIndex) & ",N,8,1"
    Form1.StatusBar1.Panels(1).Text = "Terhubung..."
    Form1.Timer1.Enabled = True
    'Form1.Timer1.Interval = Time.Text
    Form1.mPutus.Enabled = True
    Form1.mHubung.Enabled = False
    Form1.mSimpan.Enabled = False
    Form1.mHapus.Enabled = False
    Form1.mOut.Enabled = False
    Form1.mCetak.Enabled = False
    Form1.cmdKirim.Enabled = True
    Form1.atas(0).Enabled = True
    Form1.bawah(0).Enabled = True
    Form1.kanan(0).Enabled = True
    Form1.kiri(0).Enabled = True
    Form1.stop1(0).Enabled = True
    Me.Hide
       Exit Sub
errcode:
    MsgBox "Portnya salah!", vbInformation, "Informasi"
    Combo1.SetFocus
End Sub





