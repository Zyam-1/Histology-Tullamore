VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmWorkSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   10665
   ClientLeft      =   165
   ClientTop       =   150
   ClientWidth     =   14265
   Icon            =   "frmWorkSheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   14265
   Begin VB.CommandButton cmdVantage 
      Caption         =   "Vantage"
      Height          =   615
      Left            =   3720
      TabIndex        =   111
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDart 
      Height          =   315
      Left            =   -120
      Picture         =   "frmWorkSheet.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   3120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   615
      Left            =   12600
      Picture         =   "frmWorkSheet.frx":15A4
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   9960
      Width           =   690
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   13320
      Picture         =   "frmWorkSheet.frx":196B
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   9960
      Width           =   735
   End
   Begin VB.CommandButton cmdDiscrepancyLog 
      Caption         =   "&Discrepancy Log"
      Height          =   615
      Left            =   10800
      Picture         =   "frmWorkSheet.frx":1CAD
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   9960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrnPreview 
      Caption         =   "Print Preview"
      Height          =   615
      Left            =   8160
      Picture         =   "frmWorkSheet.frx":1F18
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   9960
      Width           =   1275
   End
   Begin VB.CommandButton cmdViewReports 
      Caption         =   "View Reports"
      Height          =   615
      Left            =   6960
      Picture         =   "frmWorkSheet.frx":1108A
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   9960
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox PreviewPrint 
      Height          =   375
      Left            =   14280
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   62
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdPrnReport 
      Caption         =   "&Print Report"
      Height          =   615
      Left            =   9480
      Picture         =   "frmWorkSheet.frx":1123C
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9960
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdViewScans 
      Caption         =   "View Scans"
      Height          =   615
      Left            =   5880
      Picture         =   "frmWorkSheet.frx":11657
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdScanOrder 
      Caption         =   "Scan Order"
      Height          =   675
      Left            =   4920
      Picture         =   "frmWorkSheet.frx":11B89
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraWorkSheet 
      BorderStyle     =   0  'None
      Height          =   10515
      Left            =   3720
      TabIndex        =   36
      Top             =   240
      Width           =   10455
      Begin TabDlg.SSTab SSTabMovement 
         Height          =   1455
         Left            =   5160
         TabIndex        =   99
         Top             =   7500
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   2566
         _Version        =   393216
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Specimen"
         TabPicture(0)   =   "frmWorkSheet.frx":11DE4
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Picture1"
         Tab(0).Control(1)=   "grdTracker(0)"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Stain"
         TabPicture(1)   =   "frmWorkSheet.frx":11E00
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdTracker(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Case"
         TabPicture(2)   =   "frmWorkSheet.frx":11E1C
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "grdTracker(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Block/Slide"
         TabPicture(3)   =   "frmWorkSheet.frx":11E38
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "grdTracker(3)"
         Tab(3).ControlCount=   1
         Begin VB.PictureBox Picture1 
            Height          =   255
            Left            =   -75000
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   103
            Top             =   300
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid grdTracker 
            Height          =   975
            Index           =   0
            Left            =   -74880
            TabIndex        =   100
            Top             =   360
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1720
            _Version        =   393216
            RowHeightMin    =   315
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid grdTracker 
            Height          =   975
            Index           =   1
            Left            =   -74880
            TabIndex        =   101
            Top             =   360
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1720
            _Version        =   393216
            RowHeightMin    =   315
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid grdTracker 
            Height          =   975
            Index           =   2
            Left            =   120
            TabIndex        =   102
            Top             =   360
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1720
            _Version        =   393216
            RowHeightMin    =   315
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid grdTracker 
            Height          =   975
            Index           =   3
            Left            =   -74880
            TabIndex        =   104
            Top             =   360
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1720
            _Version        =   393216
            RowHeightMin    =   315
            Appearance      =   0
         End
      End
      Begin VB.CommandButton cmdComments 
         Caption         =   "&Comments"
         Height          =   615
         Left            =   2840
         Picture         =   "frmWorkSheet.frx":11E54
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1560
         Width           =   840
      End
      Begin VB.CommandButton cmdAudit 
         Caption         =   "&Audit"
         Height          =   615
         Left            =   2160
         Picture         =   "frmWorkSheet.frx":11FC3
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox txtContainerLabel 
         Height          =   855
         Left            =   7800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtNOS 
         Height          =   855
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton cmdCytoHist 
         Caption         =   "Histo Link"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1040
         Picture         =   "frmWorkSheet.frx":12026
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1560
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdEditDemo 
         Caption         =   "Edit Demo"
         Height          =   615
         Left            =   0
         Picture         =   "frmWorkSheet.frx":1211B
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1560
         Width           =   915
      End
      Begin VB.Frame fraDemographics 
         Height          =   1575
         Left            =   0
         TabIndex        =   65
         Top             =   -30
         Width           =   5895
         Begin VB.Label lblDOD 
            AutoSize        =   -1  'True
            Caption         =   "DateOfDeath"
            Height          =   195
            Left            =   720
            TabIndex        =   98
            Top             =   1440
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
            Height          =   195
            Left            =   2280
            TabIndex        =   75
            Top             =   1440
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblSex 
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            Height          =   195
            Left            =   1800
            TabIndex        =   74
            Top             =   1440
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label lblPatientAddress3 
            AutoSize        =   -1  'True
            Caption         =   "Address3"
            Height          =   195
            Left            =   2760
            TabIndex        =   73
            Top             =   1440
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lblPatientAddress2 
            AutoSize        =   -1  'True
            Caption         =   "Address2"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   840
            Width           =   2580
         End
         Begin VB.Label lblPatientDoctor 
            Alignment       =   1  'Right Justify
            Caption         =   "Doctor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   71
            Top             =   510
            Width           =   3150
         End
         Begin VB.Label lblPatientAddress1 
            AutoSize        =   -1  'True
            Caption         =   "Address1"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   600
            Width           =   2775
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblPatientWard 
            Alignment       =   1  'Right Justify
            Caption         =   "Ward"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   69
            Top             =   240
            Width           =   3030
         End
         Begin VB.Label lblPatientName 
            Caption         =   "Patient Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   68
            Top             =   270
            Width           =   2895
         End
         Begin VB.Label lblPatientGP 
            Alignment       =   1  'Right Justify
            Caption         =   "GP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            TabIndex        =   67
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label lblPatientBorn 
            Caption         =   "Born"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdClinicalHist 
         Caption         =   "Clinical Details"
         Height          =   615
         Left            =   3800
         Picture         =   "frmWorkSheet.frx":122DC
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1560
         Width           =   1260
      End
      Begin VB.CommandButton cmdQCode 
         Height          =   255
         Left            =   4560
         Picture         =   "frmWorkSheet.frx":2144E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7500
         Width           =   375
      End
      Begin VB.CommandButton cmdMCode 
         Enabled         =   0   'False
         Height          =   255
         Left            =   9840
         Picture         =   "frmWorkSheet.frx":215B3
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3120
         Width           =   375
      End
      Begin VB.Frame fraCaseState 
         Caption         =   "Case State"
         Enabled         =   0   'False
         Height          =   675
         Left            =   0
         TabIndex        =   41
         Top             =   9000
         Width           =   5055
         Begin VB.OptionButton optState 
            Caption         =   "With Pathologist"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   24
            Top             =   265
            Width           =   1560
         End
         Begin VB.OptionButton optState 
            Caption         =   "Awaiting Authorisation"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   25
            Top             =   265
            Width           =   2115
         End
         Begin VB.OptionButton optState 
            Caption         =   "In Histology"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   265
            Value           =   -1  'True
            Width           =   4635
         End
      End
      Begin VB.Frame fraReport 
         Height          =   675
         Left            =   5400
         TabIndex        =   40
         Top             =   9000
         Width           =   3795
         Begin VB.OptionButton optReport 
            Caption         =   "Preliminary Report"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Authorised Report"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox txtPCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   11
         Top             =   2520
         Width           =   1200
      End
      Begin VB.TextBox txtQDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   7500
         Width           =   3255
      End
      Begin VB.TextBox txtQCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   19
         Top             =   7500
         Width           =   1200
      End
      Begin VB.TextBox txtPDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   2520
         Width           =   3840
      End
      Begin VB.TextBox txtMCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   14
         Top             =   3120
         Width           =   1200
      End
      Begin VB.TextBox txtMDescription 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         TabIndex        =   15
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtFindCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8160
         TabIndex        =   38
         Top             =   2400
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtFindDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         TabIndex        =   37
         Top             =   2400
         Visible         =   0   'False
         Width           =   1635
      End
      Begin RichTextLib.RichTextBox txtMicro 
         Height          =   1935
         Left            =   0
         TabIndex        =   17
         Top             =   5280
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmWorkSheet.frx":21718
      End
      Begin RichTextLib.RichTextBox txtGross 
         Height          =   1815
         Left            =   0
         TabIndex        =   13
         Top             =   3120
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3201
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmWorkSheet.frx":2179A
      End
      Begin MSComCtl2.DTPicker DTSampleTaken 
         Height          =   285
         Left            =   7740
         TabIndex        =   7
         Top             =   75
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   140574723
         CurrentDate     =   40207
      End
      Begin MSMask.MaskEdBox txtSampleRecTime 
         Height          =   285
         Left            =   9420
         TabIndex        =   10
         Top             =   435
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSampleTakenTime 
         Height          =   285
         Left            =   9420
         TabIndex        =   8
         Top             =   75
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid grdMCodes 
         Height          =   1455
         Left            =   5160
         TabIndex        =   39
         Top             =   3420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2566
         _Version        =   393216
         RowHeightMin    =   315
         BackColorBkg    =   -2147483648
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grdQCodes 
         Height          =   1155
         Left            =   0
         TabIndex        =   22
         Top             =   7800
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   3
         RowHeightMin    =   315
         BackColorBkg    =   -2147483648
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker DTSampleRec 
         Height          =   285
         Left            =   7740
         TabIndex        =   9
         Top             =   435
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "  "
         Format          =   140574723
         CurrentDate     =   40207
      End
      Begin MSFlexGridLib.MSFlexGrid grdAmendments 
         Height          =   1875
         Left            =   5160
         TabIndex        =   18
         Top             =   5280
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3307
         _Version        =   393216
         Cols            =   3
         RowHeightMin    =   315
         BackColorBkg    =   -2147483648
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grdDelete 
         Height          =   1155
         Left            =   10260
         TabIndex        =   42
         Top             =   1980
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   3
         RowHeightMin    =   315
         BackColorBkg    =   -2147483648
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grdTempMCode 
         Height          =   1155
         Left            =   8040
         TabIndex        =   82
         Top             =   2760
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   3
         RowHeightMin    =   315
         BackColorBkg    =   -2147483648
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin VB.Label lblRecordSaved 
         Caption         =   "Record Saved"
         Height          =   255
         Left            =   9240
         TabIndex        =   112
         Top             =   9360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblGeneralComments 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3000
         TabIndex        =   97
         Top             =   2280
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         Height          =   75
         Left            =   9240
         TabIndex        =   94
         Top             =   10020
         Width           =   285
      End
      Begin VB.Label lblAddendumAdded 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2160
         TabIndex        =   93
         Top             =   7740
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblWithPathologist 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   90
         Top             =   10020
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Container Labelled"
         Height          =   195
         Left            =   7800
         TabIndex        =   89
         Top             =   1680
         Width           =   1320
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Nature Of Specimen"
         Height          =   195
         Left            =   5160
         TabIndex        =   88
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   9240
         TabIndex        =   81
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   9240
         TabIndex        =   80
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   10200
         TabIndex        =   79
         Top             =   480
         Width           =   75
      End
      Begin VB.Label lblDemographics 
         AutoSize        =   -1  'True
         Caption         =   "lblDemographics"
         Height          =   195
         Left            =   5160
         TabIndex        =   76
         Top             =   600
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblSurname 
         AutoSize        =   -1  'True
         Caption         =   "lblSurname"
         Height          =   195
         Left            =   6000
         TabIndex        =   64
         Top             =   1920
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblFirstName 
         AutoSize        =   -1  'True
         Caption         =   "lblFirstname"
         Height          =   195
         Left            =   5040
         TabIndex        =   63
         Top             =   1920
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblClinicalHist 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3540
         TabIndex        =   61
         Top             =   2220
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Image imgSquare 
         Height          =   225
         Left            =   720
         Picture         =   "frmWorkSheet.frx":2181C
         Top             =   0
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgSquareCross 
         Height          =   225
         Left            =   360
         Picture         =   "frmWorkSheet.frx":21B2A
         Tag             =   "0"
         Top             =   0
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgSquareTick 
         Height          =   225
         Left            =   0
         Picture         =   "frmWorkSheet.frx":21E00
         Tag             =   "1"
         Top             =   0
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gross"
         Height          =   315
         Left            =   0
         TabIndex        =   58
         Top             =   2880
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Micro"
         Height          =   195
         Left            =   0
         TabIndex        =   57
         Top             =   5040
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "M Codes"
         Height          =   195
         Left            =   5160
         TabIndex        =   56
         Top             =   2880
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Addendum / Amendments"
         Height          =   195
         Left            =   5160
         TabIndex        =   55
         Top             =   5040
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Q Codes"
         Height          =   195
         Left            =   0
         TabIndex        =   54
         Top             =   7260
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Sample Taken"
         Height          =   195
         Left            =   6120
         TabIndex        =   53
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Sample Received"
         Height          =   195
         Left            =   6120
         TabIndex        =   52
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Preliminary Report Date:"
         Height          =   195
         Left            =   6120
         TabIndex        =   51
         Top             =   840
         Width           =   1710
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Authorised Report Date:"
         Height          =   195
         Left            =   6120
         TabIndex        =   50
         Top             =   1200
         Width           =   1710
      End
      Begin VB.Label lblPreReportDate 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   8040
         TabIndex        =   49
         Top             =   840
         Width           =   45
      End
      Begin VB.Label lblValReportDate 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   8040
         TabIndex        =   48
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Movement Tracker"
         Height          =   195
         Left            =   5160
         TabIndex        =   47
         Top             =   7260
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "P Code"
         Height          =   195
         Left            =   0
         TabIndex        =   46
         Top             =   2270
         Width           =   525
      End
      Begin VB.Label lblNopas 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5160
         TabIndex        =   45
         Top             =   120
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblMrn 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5880
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblAandE 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6480
         TabIndex        =   43
         Top             =   120
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.Frame fraSearch 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   360
      TabIndex        =   31
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton cmdDartViewer 
         Height          =   390
         Left            =   1440
         Picture         =   "frmWorkSheet.frx":220D6
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame fraLinkedCase 
         Caption         =   "Cell Block On"
         Height          =   735
         Left            =   1800
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
         Begin VB.CommandButton cmdLinkedCaseId 
            Caption         =   "CaseId"
            Height          =   375
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCaseId 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         MaxLength       =   12
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPatientId 
         Height          =   285
         Left            =   60
         TabIndex        =   2
         Top             =   1260
         Width           =   3135
      End
      Begin VB.ComboBox cmbPatientId 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   900
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Patient Identifier"
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   660
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Case Number"
         Height          =   195
         Left            =   60
         TabIndex        =   33
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.ComboBox cmbState 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   9000
      Width           =   3255
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   315
      Left            =   2100
      TabIndex        =   4
      Top             =   2100
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   420
      TabIndex        =   3
      Top             =   2100
      Width           =   1455
   End
   Begin MSComctlLib.ImageList lstImages 
      Left            =   1200
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWorkSheet.frx":229A0
            Key             =   "two"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWorkSheet.frx":22D27
            Key             =   "one"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvCaseDetails 
      Height          =   6375
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   11245
      _Version        =   393217
      Indentation     =   88
      LabelEdit       =   1
      Style           =   5
      ImageList       =   "lstImages"
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblCaseLocked 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   360
      TabIndex        =   106
      Top             =   9360
      Width           =   3285
   End
   Begin VB.Label lblCheckedBy 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   105
      Top             =   10140
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgRedInfo 
      Height          =   240
      Left            =   14100
      Picture         =   "frmWorkSheet.frx":230AE
      Top             =   7920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBlueInfo 
      Height          =   240
      Left            =   14040
      Picture         =   "frmWorkSheet.frx":234D0
      Stretch         =   -1  'True
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblWithPathologistName 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5520
      TabIndex        =   92
      Top             =   9720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      TabIndex        =   35
      Top             =   10080
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   360
      TabIndex        =   34
      Top             =   10080
      Width           =   1335
   End
   Begin VB.Menu mnuMCodesMenu 
      Caption         =   "MCodesMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMCodesDel 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuQCodesMenu 
      Caption         =   "QCodesMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuQCodesDel 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuAmendMenu 
      Caption         =   "AmendmentMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAmendDel 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuMoveSpecMenu 
      Caption         =   "MoveSpecMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMoveSpecDel 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuPopUpLevel1 
      Caption         =   "PopUpLevel1"
      Visible         =   0   'False
      Begin VB.Menu mnuAddTissueType 
         Caption         =   "Add Tissue Type"
      End
      Begin VB.Menu mnuAddCutUp 
         Caption         =   "Add Cut-Up Details"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuDisposeCase 
         Caption         =   "Dispose Case"
      End
   End
   Begin VB.Menu mnuPopupLevel2 
      Caption         =   "PopUpLevel2"
      Visible         =   0   'False
      Begin VB.Menu mnuEditTissueType 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuAllEmbedded 
         Caption         =   "All Embedded"
      End
      Begin VB.Menu mnuReferral 
         Caption         =   "Referral"
      End
      Begin VB.Menu mnuSeparator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFrozenSection 
         Caption         =   "Add Frozen Section"
      End
      Begin VB.Menu mnuTouchPrep 
         Caption         =   "Add Touch Prep"
      End
      Begin VB.Menu mnuSingleBlock 
         Caption         =   "Add Single Block"
      End
      Begin VB.Menu mnuMultipleBlocks 
         Caption         =   "Add Multiple Blocks"
      End
      Begin VB.Menu mnuSingleSlideLevel2 
         Caption         =   "Add Single Slide"
      End
      Begin VB.Menu mnuMultipleSlidesLevel2 
         Caption         =   "Add Multiple Slides"
      End
      Begin VB.Menu mnuSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelTissueType 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuPopupLevel3 
      Caption         =   "PopupLevel3"
      Visible         =   0   'False
      Begin VB.Menu mnuBlockReferral 
         Caption         =   "Referral"
      End
      Begin VB.Menu mnuSeperator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSingleSlideLevel3 
         Caption         =   "Add Single Slide"
      End
      Begin VB.Menu mnuMultipleSlidesLevel3 
         Caption         =   "Add Multiple Slides"
      End
      Begin VB.Menu mnuNoOfLevelsLevel3 
         Caption         =   "Add No. Of Levels"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddControlLevel3 
         Caption         =   "Add Control"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRoutineStainLevel3 
         Caption         =   "Add Routine Stain"
      End
      Begin VB.Menu mnuSpecialStainLevel3 
         Caption         =   "Add Special Stain"
      End
      Begin VB.Menu mnuImmunoStainLevel3 
         Caption         =   "Add Immunohistochemical Stain"
      End
      Begin VB.Menu mnuAddExtraLevels 
         Caption         =   "Add Extra Levels"
      End
      Begin VB.Menu mnuPrnBlockNumber 
         Caption         =   "Print to Block Number"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSeperator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelBlock 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuPopupLevel4 
      Caption         =   "PopupLevel4"
      Visible         =   0   'False
      Begin VB.Menu mnuAddControlLevel4 
         Caption         =   "Add Control"
      End
      Begin VB.Menu mnuRoutineStainLevel4 
         Caption         =   "Add Routine Stain"
      End
      Begin VB.Menu mnuSpecialStainLevel4 
         Caption         =   "Add Special Stain"
      End
      Begin VB.Menu mnuImmunoStainLevel4 
         Caption         =   "Add Immunohistochemical Stain"
      End
      Begin VB.Menu mnuNoOfLevelsLevel4 
         Caption         =   "Add No. Of Levels"
      End
      Begin VB.Menu mnuSeperator8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelSlide 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuPopupLevel5 
      Caption         =   "PopupLevel5"
      Visible         =   0   'False
      Begin VB.Menu mnuDelStain 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuPopupFormatGrossText 
      Caption         =   "PopupFormatGrossText"
      Visible         =   0   'False
      Begin VB.Menu mnuGrossBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuGrossItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuGrossUnderline 
         Caption         =   "Underline"
      End
   End
   Begin VB.Menu mnuPopupFormatMicroText 
      Caption         =   "PopupFormatMicroText"
      Visible         =   0   'False
      Begin VB.Menu mnuMicroBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuMicroItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuMicroUnderline 
         Caption         =   "Underline"
      End
   End
End
Attribute VB_Name = "frmWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit

Private Grid As MSFlexGrid

Dim Cada As Integer
Dim gridId As String
Dim PatientIdCombo As String
Dim PatientIdText As String
Private PrevNode As MSComctlLib.Node
Private Validated As Boolean
Private CaseWasOriginalValidated As Boolean
Private mCaseId As String
Private Search As Boolean
Private mclsToolTip As New clsToolTip
Private CodeEventLog As Boolean
Dim TreePositionX As Single
Dim TreePositionY As Single
Dim Activated As Boolean
Dim bWithPathologist As Boolean




Private Sub cmdAudit_Click()
10    With frmCaseEventLog
20        .txtCaseId = txtCaseId
30        .SID = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
40        .RunReport
50        .Show 1
60    End With
End Sub

Private Sub cmdClinicalHist_Click()
10    With frmClinicalHist
20        .Description = lblClinicalHist
30        .Move frmWorkSheet.Left + fraWorkSheet.Left + cmdClinicalHist.Left - .Width, frmWorkSheet.Top + cmdClinicalHist.Top
40        .Show vbModal
50    End With
60    CheckClinicalHist

70    mclsToolTip.ToolText(cmdClinicalHist) = lblClinicalHist
End Sub

Private Sub cmdClinicalHist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'10      mclsToolTip.ToolText(cmdClinicalHist) = lblClinicalHist
End Sub

Private Sub CheckGeneralComments()

10    On Error GoTo CheckGeneralComments_Error

20    If lblGeneralComments <> "" Then

30        cmdComments.BackColor = &H8080FF
40    Else
50        cmdComments.BackColor = &H8000000F
60    End If

70    Exit Sub

CheckGeneralComments_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmWorkSheet", "CheckGeneralComments", intEL, strES

End Sub


Private Sub CheckClinicalHist()


10    On Error GoTo CheckClinicalHist_Error

20    If lblClinicalHist <> "" Then
30        cmdClinicalHist.Caption = "Clinician Details"
40        cmdClinicalHist.BackColor = &H8080FF
50    Else
60        cmdClinicalHist.Caption = "Clinician Details"
70        cmdClinicalHist.BackColor = &H8000000F
80    End If



90    Exit Sub

CheckClinicalHist_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmWorkSheet", "CheckClinicalHist", intEL, strES


End Sub

Private Sub CheckDiscrepancyLog()
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CheckDiscrepancyLog_Error

20    sql = "SELECT * FROM Discrepancy WHERE CaseId = N'" & CaseNo & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        cmdDiscrepancyLog.BackColor = &H8080FF
70    Else
80        cmdDiscrepancyLog.BackColor = &H8000000F
90    End If


100   Exit Sub

CheckDiscrepancyLog_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmWorkSheet", "CheckDiscrepancyLog", intEL, strES, sql


End Sub




Private Sub cmdComments_Click()
10    With frmComment
20        .CommentType = "GENERAL"
30        .GeneralComment = lblGeneralComments
40        .cmdExit.Visible = True
50        .Move frmWorkSheet.Left + fraWorkSheet.Left + cmdComments.Left - .Width, frmWorkSheet.Top + cmdComments.Top
60        .Show 1
70    End With
80    CheckGeneralComments

90    mclsToolTip.ToolText(cmdComments) = lblGeneralComments
End Sub

Private Sub cmdCytoHist_Click()
10    With frmDemographics
20        .AddNew = False
30        .Link = True
40        .Show 1
50    End With
End Sub

Private Sub cmdDartViewer_Click()
10  On Error GoTo cmdDart_Click_Error
    
    Dim CaseId As String
    Dim formatedCaseID As String

20  If Dir("\\tdws08fs01.mhb.health.gov.ie\MRHP_WardEnquiry\MRHT\Netaquire\The Plumtree Group\DartViewer\DartViewer.exe") = "" Then
30      iMsg "Dart client not installed on this machine. Please contact you system administrator", vbInformation
40      Exit Sub
50  End If
    DoEvents
    DoEvents
    Sleep (55)
    
410 CaseId = Trim(txtCaseId.Text)
430 formatedCaseID = Replace(CaseId, " ", "")
    
    DoEvents
    DoEvents
    Sleep (55)
100 Shell "\\tdws08fs01.mhb.health.gov.ie\MRHP_WardEnquiry\MRHT\Netaquire\The Plumtree Group\DartViewer\DartViewer.exe " & formatedCaseID, vbNormalFocus
    DoEvents
    DoEvents
    Sleep (55)
110 Exit Sub

cmdDart_Click_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmWorkSheet", "cmdDart_Click", intEL, strES
End Sub

Private Sub cmdDiscrepancyLog_Click()
10    With frmAddDiscrepancy
20        .CaseId = txtCaseId
30        .PatientName = lblPatientName
40        .Show 1
50    End With
60    CheckDiscrepancyLog
End Sub

Private Sub cmdEditDemo_Click()
10    With frmDemographics
20        .AddNew = False
30        .Link = False
40        .Show 1
50    End With
End Sub

Private Sub cmdLinkedCaseId_Click()
      Dim LinkedCaseId As String

10    LinkedCaseId = cmdLinkedCaseId.Caption

20    CaseNo = Replace(LinkedCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")

30    sCaseLockedBy = CaseLockedBy(CaseNo)

40    If sCaseLockedBy <> UserName And sCaseLockedBy <> "" Then
50        lblCaseLocked = "RECORD BEING EDITED BY" & " " & sCaseLockedBy
60        bLocked = True
70        lblCaseLocked.BackColor = &H8080FF
80    ElseIf sCaseLockedBy = "" Then
90        LockCase CaseNo
100       bLocked = False
110       UnLockControl
120       lblCaseLocked.BackColor = &H80FF80
130       lblCaseLocked = "CSRECORD BEING EDITED BY YOU"
140   Else
150       bLocked = False
160       UnLockControl
170       lblCaseLocked.BackColor = &H80FF80
180       lblCaseLocked = "RECORD BEING EDITED BY YOU"
190   End If

200   ResetWorkSheet
210   txtCaseId = LinkedCaseId
220   If FillTree Then
230       tvCaseDetails.Nodes(1).Selected = True
240       LoadDemographics CaseNo
250       FillWorkSheet CaseNo
260   End If
270   ExpandAll tvCaseDetails

280   Select Case UCase$(UserMemberOf)
      Case "CLERICAL"
290       DisableClerical
300   Case "SCIENTIST"
310       DisableScientist
320   Case "MANAGER"
330       DisableManager
340   Case "CONSULTANT"
350       DisableConsultant
360   Case "SPECIALIST REGISTRAR"
370       DisableConsultant
380   End Select

390   If bLocked Then
400       LockControl
410   End If

'420   cmdMCode.Enabled = False
'430   txtMCode.Enabled = False
'440   txtMDescription.Enabled = False

450   Set PrevNode = Nothing
End Sub

Private Sub cmdPrnPreview_Click()
10    PrintHistology "", True
20    With frmRichText
          'Ibrahim 27-07-24
30        '.cmdPrint.Visible = False
40        '.cmdExit.Left = 0
          'Ibrahim 27-07-24
50        .rtb.SelStart = 0
60        .Show 1
70    End With
End Sub

Private Sub cmdVantage_Click()
10    With FrmCaseStatus
        If Trim(txtCaseId.Text) <> "" Then
            .txtCaseId = Trim(txtCaseId.Text)
        End If
       
       '.cmdSearch
20    .Show 1
30    End With
End Sub

Private Sub cmdViewReports_Click()
10    With frmReportViewer
20        .SampleID = CaseNo
30        .Year = 2000 + Val(Right(CaseNo, 2))
          If UCase(UserMemberOf) = "MANAGERS" Or UCase(UserMemberOf) = "MANAGER" Then
            .Width = 17385
          End If
40        .Show 1
50    End With
End Sub



Private Sub DTSampleRec_GotFocus()
10    DTSampleRec.CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub DTSampleRec_Validate(Cancel As Boolean)
10    If DTSampleRec < DTSampleTaken Then
20        frmMsgBox.Msg "SampleDate After RecDate Please Amend", , "Histology", mbExclamation
30        Cancel = True
40    ElseIf DTSampleRec = DTSampleTaken Then
50        If txtSampleTakenTime <> "" Then
60            If Format(txtSampleRecTime.FormattedText, "HH:mm") < Format(txtSampleTakenTime.FormattedText, "HH:mm") Then
70                frmMsgBox.Msg "SampleDate After RecDate Please Amend", , "Histology", mbExclamation
80                Cancel = True
90            End If
100       End If
110   End If
End Sub


Private Sub DTSampleTaken_GotFocus()
10    DTSampleTaken.CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub DTSampleTaken_Validate(Cancel As Boolean)
10    If DTSampleTaken > DTSampleRec Then
20        frmMsgBox.Msg "SampleDate After RecDate Please Amend", , "Histology", mbExclamation
30        Cancel = True
40    ElseIf DTSampleRec = DTSampleTaken Then
50        If txtSampleTakenTime <> "" Then
60            If Format(txtSampleTakenTime.FormattedText, "HH:mm") > Format(txtSampleRecTime.FormattedText, "HH:mm") Then
70                frmMsgBox.Msg "SampleDate After RecDate Please Amend", , , mbExclamation
80                Cancel = True
90            End If
100       End If
110   End If
End Sub


Private Sub grdTracker_Click(Index As Integer)

10    If Rada <> 0 Then
20        If grdTracker(Index).Rows > 1 Then
30            Select Case Cada
              Case 0, 1, 2, 3:
40                With frmMovement
50                    .Update = True
60                    .MovementId = grdTracker(Index).TextMatrix(Rada, 5)
70                    .Description = grdTracker(Index).TextMatrix(Rada, 0)
80                    .Code = grdTracker(Index).TextMatrix(Rada, 4)
90                    .RefType = SSTabMovement.Tab
100                   .Move frmWorkSheet.Left + fraWorkSheet.Left + SSTabMovement.Left - .Width, frmWorkSheet.Top + SSTabMovement.Top - SSTabMovement.Height
110                   .Show vbModal
120               End With

130           Case 6:
140               If grdTracker(Index).TextMatrix(Rada, 3) <> "" Then
150                   grdTracker(Index).row = Rada
160                   grdTracker(Index).col = Cada
170                   If grdTracker(Index).CellPicture = imgSquare.Picture Then
180                       Set grdTracker(Index).CellPicture = imgSquareTick.Picture
190                   ElseIf grdTracker(Index).CellPicture = imgSquareTick.Picture Then
200                       Set grdTracker(Index).CellPicture = imgSquareCross.Picture
210                   Else
220                       Set grdTracker(Index).CellPicture = imgSquare.Picture
230                   End If
240               End If
250           Case 10:
260               With frmReferralDetails
270                   .Move frmWorkSheet.Left + fraWorkSheet.Left + SSTabMovement.Left - .Width, frmWorkSheet.Top + SSTabMovement.Height
280                   .Show 1
290               End With
300           End Select
310       End If
320   End If
End Sub

Private Sub grdTracker_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
10    GridToolTip grdTracker(Index), X, Y
End Sub

Private Sub grdTracker_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
10    If grdTracker(Index).Rows > 1 Then
20        Rada = grdTracker(Index).MouseRow
30        Cada = grdTracker(Index).MouseCol
40        If Rada <> 0 Then
50            gridId = grdTracker(Index).TextMatrix(Rada, 5)
60            HighlightRow (gridId)
70        End If
80    End If
End Sub



Private Sub mnuAddControlLevel3_Click()

      Dim tnode As MSComctlLib.Node
      Dim UniqueId As String

      'Add Control to the tree
10    With tvCaseDetails.SelectedItem

20        UniqueId = GetUniqueID

30        Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L4" & UniqueId, "Control", 1, 2)
40        .Expanded = True
50        tnode.Selected = True
60    End With
70    DataChanged = True
80    TreeChanged = True

End Sub

Private Sub mnuAddControlLevel4_Click()

      Dim tnode As MSComctlLib.Node
      Dim UniqueId As String

      'Add Control to the tree
10    With tvCaseDetails.SelectedItem

20        UniqueId = GetUniqueID

30        Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L4" & UniqueId, "Control", 1, 2)
40        .Expanded = True
50        tnode.Selected = True
60    End With
70    DataChanged = True
80    TreeChanged = True

End Sub

Private Sub mnuAddCutUp_Click()
10    With frmCutUpEmbed
20        .SingleEdit = True
30        .Phase = "Cut Up"
40        .Show 1
50    End With
End Sub

Private Sub mnuAddExtraLevels_Click()

10    With frmInputNo
20        .InputType = "L"
30        .Label = "Please Enter Number of Levels"
40        .Move ScaleX(TreePositionX, vbPixels, vbTwips), ScaleY(TreePositionY, vbPixels, vbTwips)
50        .Show 1
60    End With
End Sub

Private Sub mnuAllEmbedded_Click()

      Dim sAllEmbedded As String
      Dim oNode As MSComctlLib.Node

10    Set oNode = tvCaseDetails.SelectedItem

20    If mnuAllEmbedded.Checked = True Then
30        mnuAllEmbedded.Checked = False
40        sAllEmbedded = Mid(oNode.Text, 1, Len(oNode.Text) - 5)
50    Else
60        mnuAllEmbedded.Checked = True
70        sAllEmbedded = oNode.Text & " (AE)"
80    End If

90    oNode.Text = sAllEmbedded

100   TreeChanged = True
110   DataChanged = True
End Sub

Private Sub mnuBlockReferral_Click()
      Dim UniqueId As String

10    UniqueId = GetUniqueID

20    With frmMovement
30        .Update = False
40        .MovementId = UniqueId
50        .Description = Trim(tvCaseDetails.SelectedItem)
60        .RefType = 3
70        .Move frmWorkSheet.Left + fraWorkSheet.Left + SSTabMovement.Left - .Width, frmWorkSheet.Top + SSTabMovement.Top - SSTabMovement.Height
80        .Show vbModal
90    End With
End Sub

Private Sub mnuDelBlock_Click()

10    DeleteNode
End Sub

Private Sub DeleteNode()
      Dim CaseNode As MSComctlLib.Node
      Dim sql As String
      Dim tb As Recordset
      Dim tempNode As MSComctlLib.Node


10    With tvCaseDetails

20        If frmMsgBox.Msg("Are You Sure You Want To Delete" & " " & .SelectedItem.Text & "?", mbYesNo, "Histology", mbQuestion) = 1 Then

30            Set CaseNode = .SelectedItem
40            While Not CaseNode.Parent Is Nothing
50                Set CaseNode = CaseNode.Parent
60            Wend

70            If Not PrevNode Is Nothing Then
80                Set PrevNode = PrevNode.Parent
90            Else
100               Set PrevNode = CaseNode
110           End If

120           sql = "SELECT * FROM CaseTree WHERE LocationID = N'" & Right(.SelectedItem.Key, Len(.SelectedItem.Key) - 2) & "'"
130           Set tb = New Recordset
140           RecOpenServer 0, tb, sql

150           Set tempNode = .SelectedItem
160           If Not tb.EOF Then


170               Do Until .SelectedItem Is Nothing
180                   DeleteChildren .SelectedItem
190                   .SelectedItem = Nothing
200               Loop

210               With frmComment
220                   .Node = tempNode
230                   .CommentType = "DELTREE"
240                   .cmdExit.Visible = False
250                   .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top
260                   .Show 1
270               End With

280               If InStr(1, tempNode.Text, "Block") Then
290                   DeleteBlockDetails tempNode
300               End If

310               sql = "DELETE FROM CaseTree WHERE LocationID = N'" & Right(tempNode.Key, Len(tempNode.Key) - 2) & "'"
320               Cnxn(0).Execute sql

330           End If

340           .Nodes.Remove tempNode.Index

350           DataChanged = True
360           TreeChanged = True

370       End If
380   End With
End Sub

Private Sub DeleteBlockDetails(blockNode As MSComctlLib.Node)
      Dim sql As String

10    On Error GoTo DeleteBlockDetails_Error

20    sql = "DELETE FROM BlockDetails WHERE CaseId = N'" & CaseNo & "' " & _
            "AND Block = N'" & blockNode.Text & "' "

30    If InStr(1, blockNode.Parent.Text, "Frozen Section") Then
40        sql = sql & "AND TissueListId = N'" & Right(blockNode.Parent.Parent.Key, Len(blockNode.Parent.Parent.Key) - 2) & "' "
50    Else
60        sql = sql & "AND TissueListId = N'" & Right(blockNode.Parent.Key, Len(blockNode.Parent.Key) - 2) & "' "
70    End If

80    Cnxn(0).Execute sql


90    Exit Sub

DeleteBlockDetails_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmWorkSheet", "DeleteBlockDetails", intEL, strES, sql


End Sub



Private Sub DeleteChildren(ParentNode As MSComctlLib.Node)

      Dim objNode As MSComctlLib.Node
      Dim sql As String

10    Set objNode = ParentNode.Child

20    Do Until objNode Is Nothing

30        If InStr(1, objNode.Text, "Block") Then
40            DeleteBlockDetails objNode
50        End If

60        sql = "DELETE FROM CaseTree WHERE LocationID = N'" & Right(objNode.Key, Len(objNode.Key) - 2) & "'"
70        Cnxn(0).Execute sql

80        DeleteChildren objNode

90        Set objNode = objNode.Next

100   Loop
End Sub

Private Sub mnuDelSlide_Click()

10    DeleteNode
End Sub

Private Sub mnuDelTissueType_Click()

10    DeleteNode
End Sub

Private Sub mnuDisposeCase_Click()
      Dim sql As String
      Dim tb As Recordset
      Dim Sql2 As String
      Dim strSID As String
      Dim blnDisposed As Boolean
      Dim strReason As String

10    On Error GoTo mnuDisposeCase_Click_Error

20    strSID = Replace(txtCaseId, " ", "")
30    strSID = Replace(strSID, "/", "")
40    blnDisposed = False

50    strReason = iBOX("Enter reason for disposal:", , "", False)    'Enter Reason

60    sql = "SELECT LocationName FROM CaseTree CT INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
            "WHERE C.State = '" & "Authorised" & "' AND (CT.Disposal IS NULL OR CT.Disposal = '') " & _
            "AND TissueTypeListId IS NOT NULL " & _
            "AND SUBSTRING(CT.CaseId,2,1) = 'A' " & _
            "and c.CaseId = N'" & strSID & "' "

70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql

90    Do While Not tb.EOF
100       Sql2 = "Update CaseTree SET Disposal = 'D', DisposalComment = N'" & strReason & "', DisposedBy = N'" & AddTicks(UserName) & "', " & _
                 "DisposalDate = '" & Format(Now, "dd/mmm/yyyy") & "'" & _
                 "WHERE CaseId = N'" & strSID & "' and LocationName = N'" & tb!LocationName & "" & "' "
110       Cnxn(0).Execute (Sql2)

120       tb.MoveNext
130       blnDisposed = True
140   Loop

150   CaseUpdateLogEvent strSID, Disposal, " Disposed (Comments: " & strReason & ")"

160   If blnDisposed Then
170       iMsg "Case id: " & txtCaseId & " disposed."
180   End If


190   Exit Sub

mnuDisposeCase_Click_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmWorkSheet", "mnuDisposeCase_Click", intEL, strES, sql

End Sub
'ZyamTissue
Private Sub mnuEditTissueType_Click()
10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .Update = True
50            .ListType = "T"
60            .ListTypeName = "Tissue Type"
70            .ListTypeNames = "Tissue Types"
80            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top
90            .Show
100       End With
110   End With
120   frmWorkSheet.Enabled = False
End Sub

Private Sub mnuFrozenSection_Click()
      Dim tnode As MSComctlLib.Node
      Dim NoOfFrozenSections As Integer
      Dim iFrozenSection As Integer
      Dim UniqueId As String
      Dim obj As MSComctlLib.Node
      Dim tempNode As MSComctlLib.Node
      Dim i As Integer

10    NoOfFrozenSections = GetFrozenSectionNumber(tvCaseDetails.SelectedItem)
20    With tvCaseDetails.SelectedItem
30        UniqueId = GetUniqueID

40        iFrozenSection = NoOfFrozenSections + 1


50        Set obj = .Child
60        For i = 1 To .Children
70            If UCase(Left(obj.Text, 14)) = "Frozen Section" Then
80                Set tempNode = obj
90            End If
100           Set obj = obj.Next
110       Next

120       If tempNode Is Nothing Then
130           Set tempNode = .Child
140           If Not tempNode Is Nothing Then
150               Set tnode = tvCaseDetails.Nodes.Add(tempNode.Key, tvwFirst, "L2" & UniqueId, "Frozen Section" & " " & iFrozenSection, 1, 2)
160           Else
170               Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L2" & UniqueId, "Frozen Section" & " " & iFrozenSection, 1, 2)
180           End If
190       Else
200           Set tnode = tvCaseDetails.Nodes.Add(tempNode.Key, tvwNext, "L2" & UniqueId, "Frozen Section" & " " & iFrozenSection, 1, 2)
210       End If

220       .Expanded = True
230       tnode.Selected = True
240   End With
250   DataChanged = True
260   TreeChanged = True
End Sub

Private Sub mnuGrossBold_Click()
10    If IsNull(txtGross.SelBold) Then
20        txtGross.SelBold = True
30    ElseIf txtGross.SelBold = False Then
40        txtGross.SelBold = True
50    Else
60        txtGross.SelBold = False
70    End If
End Sub

Private Sub mnuGrossItalic_Click()
10    If IsNull(txtGross.SelItalic) Then
20        txtGross.SelItalic = True
30    ElseIf txtGross.SelItalic = False Then
40        txtGross.SelItalic = True
50    Else
60        txtGross.SelItalic = False
70    End If
End Sub

Private Sub mnuGrossUnderline_Click()
10    If IsNull(txtGross.SelUnderline) Then
20        txtGross.SelUnderline = True
30    ElseIf txtGross.SelUnderline = False Then
40        txtGross.SelUnderline = True
50    Else
60        txtGross.SelUnderline = False
70    End If
End Sub

Private Sub mnuImmunoStainLevel3_Click()

10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .ListType = "IS"
50            .Level = "L2"
60            .ListTypeName = "Immunohistochemical Stain"
70            .ListTypeNames = "Immunohistochemical Stains"
80            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top
90            .Show
100       End With
110   End With
120   frmWorkSheet.Enabled = False
End Sub



Private Sub mnuMicroBold_Click()
10    If IsNull(txtMicro.SelBold) Then
20        txtMicro.SelBold = True
30    ElseIf txtMicro.SelBold = False Then
40        txtMicro.SelBold = True
50    Else
60        txtMicro.SelBold = False
70    End If
End Sub

Private Sub mnuMicroItalic_Click()
10    If IsNull(txtMicro.SelItalic) Then
20        txtMicro.SelItalic = True
30    ElseIf txtMicro.SelItalic = False Then
40        txtMicro.SelItalic = True
50    Else
60        txtMicro.SelItalic = False
70    End If
End Sub

Private Sub mnuMicroUnderline_Click()
10    If IsNull(txtMicro.SelUnderline) Then
20        txtMicro.SelUnderline = True
30    ElseIf txtMicro.SelUnderline = False Then
40        txtMicro.SelUnderline = True
50    Else
60        txtMicro.SelUnderline = False
70    End If
End Sub

Private Sub mnuMultipleSlidesLevel3_Click()
10    AddMultipleSlides
End Sub


Private Sub mnuNoOfLevelsLevel3_Click()
10    With frmInputNo
20        .InputType = "C"
30        .Label = "Please Enter Number of Levels"
40        .Move ScaleX(TreePositionX, vbPixels, vbTwips), ScaleY(TreePositionY, vbPixels, vbTwips)
50        .Show 1
60    End With
End Sub

Private Sub mnuNoOfLevelsLevel4_Click()
10    With frmInputNo
20        .InputType = "C"
30        .Label = "Please Enter Number of Levels"
40        .Move ScaleX(TreePositionX, vbPixels, vbTwips), ScaleY(TreePositionY, vbPixels, vbTwips)
50        .Show 1
60    End With
End Sub

Private Sub mnuOpen_Click()
10    txtCaseId = tvCaseDetails.SelectedItem
20    CaseNo = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
30    If FillTree Then
40        Set PrevNode = Nothing
50        tvCaseDetails.Nodes(1).Selected = True
60        ResetWorkSheet
70        txtPatientId = ""
80        LoadDemographics CaseNo
90        FillWorkSheet CaseNo
100       fraWorkSheet.Enabled = True
110       ExpandAll tvCaseDetails
120       Search = False
130   End If
End Sub





Private Sub mnuReferral_Click()
      Dim UniqueId As String
      Dim arr() As String

      'ITS 819013

10    arr = Split(tvCaseDetails.SelectedItem, ":")

20    UniqueId = GetUniqueID

30    With frmMovement
40        .Update = False
50        .MovementId = UniqueId
60        .Description = Trim(arr(2))
70        .Code = Trim(arr(1))
80        .SpecId = Trim(arr(0))
90        .RefType = 0
100       .Move frmWorkSheet.Left + fraWorkSheet.Left + SSTabMovement.Left - .Width, frmWorkSheet.Top + SSTabMovement.Top - SSTabMovement.Height
110       .Show vbModal
120   End With
End Sub

Private Sub mnuRoutineStainLevel3_Click()

10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .ListType = "RS"
50            .Level = "L2"
60            .ListTypeName = "Routine Stain"
70            .ListTypeNames = "Routine Stains"
80            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top
90            .Show
100       End With
110   End With
120   frmWorkSheet.Enabled = False
End Sub

Private Sub mnuRoutineStainLevel4_Click()

10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .ListType = "RS"
50            .Level = "L3"
60            .ListTypeName = "Routine Stain"
70            .ListTypeNames = "Routine Stains"


80            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top
90            .Show
100       End With
110   End With
120   frmWorkSheet.Enabled = False
End Sub



Private Sub mnuSingleSlideLevel3_Click()
10    AddSingleSlide
End Sub

Private Sub mnuSpecialStainLevel3_Click()

10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .ListType = "SS"
50            .Level = "L2"
60            .ListTypeName = "Special Stain"
70            .ListTypeNames = "Special Stains"
80            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top
90            .Show
100       End With
110   End With
120   frmWorkSheet.Enabled = False
End Sub

Private Sub mnuTouchPrep_Click()
      Dim tnode As MSComctlLib.Node
      Dim i As Integer
      Dim UniqueId As String
      Dim TempId As String
      Dim obj As MSComctlLib.Node
      Dim tempNode As MSComctlLib.Node

10    With tvCaseDetails.SelectedItem
20        UniqueId = GetUniqueID

30        Set obj = .Child
40        For i = 1 To .Children
50            If InStr(1, obj.Text, "Frozen Section") Then
60                Set tempNode = obj
70            End If
80            Set obj = obj.Next
90        Next

100       If tempNode Is Nothing Then
110           Set tempNode = .Child
120           If Not tempNode Is Nothing Then
130               Set tnode = tvCaseDetails.Nodes.Add(tempNode.Key, tvwFirst, "L2" & UniqueId, "Touch Prep", 1, 2)
140           Else
150               Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L2" & UniqueId, "Touch Prep", 1, 2)
160           End If
170       Else
180           Set tnode = tvCaseDetails.Nodes.Add(tempNode.Key, tvwNext, "L2" & UniqueId, "Touch Prep", 1, 2)
190       End If

          'Add 6 slides automatically to a touch prep
200       For i = 1 To 6
210           TempId = GetUniqueID
220           Set tnode = tvCaseDetails.Nodes.Add("L2" & UniqueId, tvwChild, "L3" & TempId, "Slide" & " " & i, 1, 2)
230       Next

240       .Expanded = True
250       tnode.Selected = True
260   End With
270   DataChanged = True
280   TreeChanged = True
End Sub

Private Sub optState_Click(Index As Integer)

10    optReport(1).Value = False
End Sub


Private Sub optState_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
10    If Index = 1 Then
20        With frmWithPathologist
30            .PathologistName = lblWithPathologistName
40            .CheckedBy = lblCheckedBy
50            .Move frmWorkSheet.Left + fraWorkSheet.Left + fraCaseState.Left, frmWorkSheet.Top + fraWorkSheet.Top + fraCaseState.Top - fraCaseState.Height
60            .Show 1
70        End With
80    Else
90        lblWithPathologist = ""
100       lblWithPathologistName = ""
110       lblCheckedBy = ""
120   End If
End Sub





Private Sub tvCaseDetails_KeyUp(KeyCode As Integer, Shift As Integer)

10    On Error GoTo tvCaseDetails_KeyUp_Error

20    Select Case Left(tvCaseDetails.SelectedItem.Key, 2)
      Case "L1"
          'if click on a T Code on the tree then enable the M Code textboxes and button
          'M Code is associated with T Code
30        If (Not Validated Or lblAddendumAdded = "TRUE") And UCase$(UserMemberOf) <> "CLERICAL" Then
40            cmdMCode.Enabled = True
50            txtMCode.Enabled = True
60            txtMDescription.Enabled = True
70            grdMCodes.Enabled = True
80        End If

90        InitializeGridCodes grdMCodes
100       FillMCodes Right(tvCaseDetails.SelectedItem.Key, Len(tvCaseDetails.SelectedItem.Key) - 2), Left(tvCaseDetails.SelectedItem.Text, 1)
110   Case Else
120       cmdMCode.Enabled = False
130       txtMCode.Enabled = False
140       txtMDescription.Enabled = False
150       InitializeGridCodes grdMCodes
160   End Select

170   Exit Sub

tvCaseDetails_KeyUp_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmWorkSheet", "tvCaseDetails_KeyUp", intEL, strES

End Sub

Private Sub tvCaseDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Dim oNode As MSComctlLib.Node
      Dim sql As String

10    On Error GoTo tvCaseDetails_MouseMove_Error

20    Set oNode = tvCaseDetails.HitTest(X, Y)

30    If Not oNode Is Nothing Then
40        If InStr(1, oNode.Text, "Block") Then
              'if Levels requested on block add tooltip to show how many
50            If oNode.Tag <> "" And oNode.Tag <> 0 Then
60                tvCaseDetails.ToolTipText = oNode.Tag & " Levels Requested"
70            End If
80        ElseIf InStr(1, oNode.Text, "Slide") Then
              'show how many levels have been done on a slide in a tooltip
90            If oNode.Tag <> "" And oNode.Tag <> 0 Then
100               tvCaseDetails.ToolTipText = oNode.Tag & " Levels"
110           End If
120       End If
130   Else
140       tvCaseDetails.ToolTipText = ""
150   End If

160   Exit Sub

tvCaseDetails_MouseMove_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmWorkSheet", "tvCaseDetails_MouseMove", intEL, strES, sql

End Sub

Private Sub tvCaseDetails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    On Error GoTo tvCaseDetails_MouseUp_Error
20    If Not Activated Then Exit Sub

30    Set tvCaseDetails.SelectedItem = tvCaseDetails.HitTest(X, Y)
40    If tvCaseDetails.SelectedItem Is Nothing Then
50        Exit Sub
60    End If
70    Select Case Left(tvCaseDetails.SelectedItem.Key, 2)
      Case "L1"
          'if click on a T Code on the tree then enable the M Code textboxes and button
          'M Code is associated with T Code
80        If (Not Validated Or lblAddendumAdded = "TRUE") And UCase$(UserMemberOf) <> "CLERICAL" Then
90            If Not bLocked Then
100               cmdMCode.Enabled = True
110               txtMCode.Enabled = True
120               txtMDescription.Enabled = True
130               grdMCodes.Enabled = True
140           End If
150       End If

160       InitializeGridCodes grdMCodes
170       FillMCodes Right(tvCaseDetails.SelectedItem.Key, Len(tvCaseDetails.SelectedItem.Key) - 2), Left(tvCaseDetails.SelectedItem.Text, 1)
180   Case Else
190       cmdMCode.Enabled = False
200       txtMCode.Enabled = False
210       txtMDescription.Enabled = False
220       InitializeGridCodes grdMCodes
230   End Select

240   Exit Sub

tvCaseDetails_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmWorkSheet", "tvCaseDetails_MouseUp", intEL, strES


End Sub

Private Sub txtCaseId_GotFocus()
10    mCaseId = txtCaseId
End Sub

Private Sub txtCaseId_KeyPress(KeyAscii As Integer)
      Dim lngSel As Long, lngLen As Long


10    On Error GoTo txtCaseId_KeyPress_Error


20    If DataChanged = False Then
30        If UCase(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
40            Call ValidateTullCaseId(KeyAscii, Me)
50        Else
60           Call ValidateLimCaseId(KeyAscii, Me)
70        End If


80    Else
90        If frmMsgBox.Msg("Save Changes", mbYesNo, "Histology", mbQuestion) = 1 Then
100           cmdSave_Click
110       End If

120       If UCase(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
130           Call ValidateTullCaseId(KeyAscii, Me)
140       Else
150           Call ValidateLimCaseId(KeyAscii, Me)
160       End If


170       ResetWorkSheet
180       tvCaseDetails.Nodes.Clear
190       'txtPatientId = ""
         
200   End If
      KeyAscii = Asc(UCase(Chr(KeyAscii)))


210   Exit Sub

txtCaseId_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmWorkSheet", "txtCaseId_KeyPress", intEL, strES

End Sub


Private Sub cmbPatientId_Click()
10    If cmbPatientId = "NOPAS" Then
20        txtPatientId.Text = lblNopas
30    ElseIf cmbPatientId = "MRN" Then
40        txtPatientId.Text = lblMrn
50    ElseIf cmbPatientId = "A&E No" Then
60        txtPatientId.Text = lblAandE
70    End If
End Sub

Private Sub cmdClear_Click()
10    If DataChanged = False Then
20        ResetSearch
30        ResetWorkSheet

40        DataMode = DataModeNew
50        cmdSearch.Enabled = True
60    Else
70        If frmMsgBox.Msg("Save Changes", mbYesNo, "Histology", mbQuestion) = 1 Then
80            cmdSave_Click
90        End If

100       ResetSearch
110       ResetWorkSheet
120       DataMode = DataModeNew
130       cmdSearch.Enabled = True

140   End If
End Sub

Private Sub cmdExit_Click()
10    pBar = 0
20    Unload Me
End Sub

Private Sub cmdMCode_Click()
      Dim i As Integer
      Dim UniqueId As String
      Dim TissueTypeId As String
      Dim TissueTypeLetter As String
      Dim TissuePath As String
      Dim TissueTypeListId As String
      Dim TissueTypeCode As String
      Dim X As Integer

10    On Error GoTo cmdMCode_Click_Error

20    For i = 1 To grdMCodes.Rows - 1
30        If UCase(grdMCodes.TextMatrix(i, 0)) = UCase(txtMCode) Then
40            MsgBox "Item already exists in the list, Please choose different item", vbInformation
50            Exit Sub
60        End If
70    Next i
80    If txtMCode <> "" Then
90        X = InStr(5, tvCaseDetails.SelectedItem.Text, " : ") - 5
100       If X > 0 Then
110           UniqueId = GetUniqueID
120           TissueTypeId = Right(tvCaseDetails.SelectedItem.Key, Len(tvCaseDetails.SelectedItem.Key) - 2)
130           TissueTypeLetter = Left(tvCaseDetails.SelectedItem.Text, 1)
140           TissuePath = tvCaseDetails.SelectedItem.Text

150           TissueTypeCode = Mid(tvCaseDetails.SelectedItem.Text, 5, X)
160           TissueTypeListId = GetListID(TissueTypeCode, "T")

170           grdMCodes.AddItem txtMCode & vbTab & txtMDescription & vbTab & UniqueId _
                                & vbTab & TissueTypeId & vbTab & TissueTypeLetter & vbTab & TissuePath _
                                & vbTab & TissueTypeListId, grdMCodes.Rows
              'since M code is per specimen need to store all the M Codes in seperate temp grid
              'so that when click on another specimen the M Code is not lost for the previous one
180           grdTempMCode.AddItem txtMCode & vbTab & txtMDescription & vbTab & UniqueId _
                                   & vbTab & TissueTypeId & vbTab & TissueTypeLetter & vbTab & TissuePath _
                                   & vbTab & TissueTypeListId, grdTempMCode.Rows
190           DataChanged = True
200       End If
210   End If
220   txtMCode = ""
230   txtMDescription = ""

240   Exit Sub

cmdMCode_Click_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmWorkSheet", "cmdMCode_Click", intEL, strES

End Sub

Private Sub cmdPrnReport_Click()
      Dim tb As New Recordset
      Dim cc As Recordset
      Dim sql As String
      Dim NoOfCopies As String
      Dim i As Integer
      Dim lc As Recordset

10    On Error GoTo cmdPrnReport_Click_Error

      'Find out the number of copies needed to be printed
      '_________________________________________________________________________________________
20    sql = "SELECT * FROM Cases C " & _
            "INNER JOIN Demographics D ON D.CaseId = C.CaseId " & _
            "WHERE C.CaseID = N'" & CaseNo & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If Not tb.EOF Then
60        If Not tb!DuplicatesPrinted Then
70            If Mid(CaseNo, 2, 1) = "P" Or Mid(CaseNo, 2, 1) = "A" Then
80                NoOfCopies = GetOptionSetting("AutopsyPrintouts", "2")
90            Else
100               If UCase(Left(tb!Source & "", 1)) = "T" Then
110                   NoOfCopies = GetOptionSetting("SourceTullamorePrintouts", "2")
120               ElseIf UCase(Left(tb!Source & "", 1)) = "M" Then
130                   NoOfCopies = GetOptionSetting("SourceMullingarPrintouts", "3")
140               ElseIf UCase(Left(tb!Source & "", 1)) = "P" Then
150                   NoOfCopies = GetOptionSetting("SourcePortlaoisePrintouts", "3")
160               Else
170                   NoOfCopies = GetOptionSetting("SourceGPPrintouts", "2")
180               End If
190           End If
200       Else
210           NoOfCopies = "1"
220       End If
          '_________________________________________________________________________________________

          'Print copies of report
230       For i = 1 To Val(NoOfCopies)
240           PreviewPrint.Cls
250           PrintHistology , , CStr(i)
260       Next i
          '_________________________________________________________________________________________

          'If copies of reports are required for a consultant print here
270       If Not tb!DuplicatesPrinted Then

280           sql = "SELECT Consultant FROM SendCopyTo WHERE CaseID = N'" & CaseNo & "'"
290           Set cc = New Recordset
300           RecOpenServer 0, cc, sql

310           Do While Not cc.EOF
320               PreviewPrint.Cls
330               PrintHistology cc!Consultant
340               cc.MoveNext
350           Loop
360       End If
          '_________________________________________________________________________________________

          'If Linked case set case as printed because don't want it appearing in not printed list
370       If tb!LinkedCaseId & "" <> "" Then
380           sql = "SELECT * FROM Cases WHERE CaseId = '" & tb!LinkedCaseId & "'"
390           Set lc = New Recordset
400           RecOpenServer 0, lc, sql

410           If Not lc.EOF Then
420               lc!DuplicatesPrinted = 1
430               lc!PrintedVal = 1
440               lc.Update
450           End If
460       End If

470       tb!DuplicatesPrinted = 1
480       tb!PrintedVal = 1
490       tb.Update
500   End If
      '_________________________________________________________________________________________

      'If report printed and it's saved in database then display "View Reports" button
510   sql = "Select * From Reports Where Sampleid = '" & CaseNo & "'"
520   Set tb = New Recordset
530   RecOpenServer 0, tb, sql

540   If Not tb.EOF Then
550       cmdViewReports.Visible = True
560   End If
      '_________________________________________________________________________________________


570   Exit Sub

cmdPrnReport_Click_Error:

      Dim strES As String
      Dim intEL As Integer

580   intEL = Erl
590   strES = Err.Description
600   LogError "frmWorkSheet", "cmdPrnReport_Click", intEL, strES, sql

End Sub

Private Sub cmdQCode_Click()
      Dim UniqueId As String
      Dim i As Integer
      Dim blnOpenMovementTracker As Boolean

10    On Error GoTo cmdQCode_Click_Error

20    If txtQCode <> "" Then
30        UniqueId = GetUniqueID

40        grdQCodes.AddItem txtQCode & vbTab & txtQDescription & vbTab & UniqueId, grdQCodes.Rows
50        For i = 1 To grdQCodes.Rows - 1
60            If grdQCodes.TextMatrix(i, 0) = "Q021" Then
70                With grdQCodes
80                    .row = i
90                    .col = 0
100                   .CellForeColor = vbRed
110                   .col = 1
120                   .CellForeColor = vbRed
130               End With
140           End If
150       Next

160       If txtQCode = "Q020" Or txtQCode = "Q021" Or txtQCode = "Q022" Then
              'if Q code any of above then it is an amendment
170           With frmAmendments
180               .Update = False
190               .AmendId = UniqueId
200               .Code = txtQCode
210               .Description = txtQDescription
220               .Move frmWorkSheet.Left + fraWorkSheet.Left + grdAmendments.Left - .Width, frmWorkSheet.Top + grdAmendments.Top
230               .Show vbModal
240           End With
              'if it was validated and the Q Code is added then it is moved back to awaiting authorisation
250           If Validated Then
260               optState(2).Value = True    'Awaiting Authorisation
270               optState(2).Tag = userCode    'Remember User Code who added Q020,Q021,Q022 because when saved the case should be assigned to them. ITS 819005
280               optReport(1).Value = False
290           End If
300       End If


310       blnOpenMovementTracker = False
320       If frmList.External Then
330           If Trim$(txtQCode) = "Q017" Then    'Case subject to MDT review
340               If UCase$(UserMemberOf) = "MANAGER" Or UCase$(UserMemberOf) = "SCIENTIST" Then
350                   blnOpenMovementTracker = True    'Show Movement Tracker window for these users ONLY
360               End If
370           Else
380               blnOpenMovementTracker = True    'Show Movement Tracker window for all Q codes marked as External
390           End If
400           If blnOpenMovementTracker Then
410               With frmMovement
420                   .Update = False
430                   .MovementId = UniqueId
440                   .Description = txtQDescription
450                   .Code = txtQCode
460                   .RefType = 2
470                   .Move frmWorkSheet.Left + fraWorkSheet.Left + SSTabMovement.Left - .Width, frmWorkSheet.Top + SSTabMovement.Top - SSTabMovement.Height
480                   .Show vbModal
490               End With
500           End If
510       End If
520       txtQCode = ""
530       txtQDescription = ""

540   End If

550   Exit Sub

cmdQCode_Click_Error:

      Dim strES As String
      Dim intEL As Integer

560   intEL = Erl
570   strES = Err.Description
580   LogError "frmWorkSheet", "cmdQCode_Click", intEL, strES


End Sub


Private Sub FillTracker()
      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim Index As Integer

10    On Error GoTo FillTracker_Error

20    sql = "Select * From CaseMovements C " & _
            "Where C.CaseID = N'" & CaseNo & "'"


30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        While Not tb.EOF
70            s = tb!Description & "" & vbTab & tb!SpecId & "" & _
                  vbTab & tb!DateSent & "" & vbTab & tb!DateReceived & "" & vbTab & _
                  tb!Code & "" & vbTab & tb!CaseListId & vbTab & vbTab & tb!Type & "" & _
                  vbTab & tb!Destination & vbTab & tb!ReferralReason & ""

80            Select Case UCase(tb!Type & "")
              Case "SPECIMEN"
90                Index = 0
100           Case "STAIN"
110               Index = 1
120           Case "CASE"
130               Index = 2
140           Case "BLOCK/SLIDE"
150               Index = 3
160           End Select

170           With grdTracker(Index)
180               .AddItem s, .Rows
190               .row = .Rows - 1
200               .col = 6
210               .CellPictureAlignment = flexAlignCenterCenter

220               If Not IsNull(tb!Agreed) Then
230                   If tb!Agreed = "1" Then
240                       Set .CellPicture = imgSquareTick.Picture
250                   ElseIf tb!Agreed = "0" Then
260                       Set .CellPicture = imgSquareCross.Picture
270                   Else
280                       Set .CellPicture = imgSquare.Picture
290                   End If
300               Else
310                   Set .CellPicture = imgSquare.Picture
320               End If

330               .col = 10
340               .CellPictureAlignment = flexAlignCenterCenter

                  'if there is a discrepency with what was sent out and what was returned then set the image to red else blue
350               If CheckReferralDiscrep(tb!CaseListId & "") Then
360                   Set .CellPicture = imgRedInfo.Picture
370               Else
380                   Set .CellPicture = imgBlueInfo.Picture
390               End If

400               ChangeTabCaptionColour SSTabMovement, Picture1, vbRed, tb!Type & "", Index
410           End With

420           tb.MoveNext
430       Wend
440   End If

450   Exit Sub

FillTracker_Error:

      Dim strES As String
      Dim intEL As Integer

460   intEL = Erl
470   strES = Err.Description
480   LogError "frmWorkSheet", "FillTracker", intEL, strES, sql


End Sub


Private Sub ArchiveCases(Id As String, Optional sType As String = "")
      Dim tb As Recordset
      Dim tbArc As Recordset
      Dim sql As String
      Dim State As String
      Dim i As Integer
      Dim SampleTaken As String
      Dim SampleReceived As String
      Dim PreReportDate As String
      Dim ValReportDate As String


      'Audit Cases (no triggers used as the Gross and Micro fields use ntext)
10    On Error GoTo ArchiveCases_Error

20    sql = "SELECT * FROM Cases WHERE " & _
            "CaseID = N'" & Id & "'"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    sql = "SELECT * FROM CasesAudit WHERE 0 = 1"
60    Set tbArc = New Recordset
70    RecOpenServer 0, tbArc, sql

80    Do While Not tb.EOF


90        For i = 0 To 2
100           If optState(i).Value = True Then
110               State = optState(i).Caption
120           End If
130       Next

140       If State = "" Then
150           State = "Authorised"
160       End If

170       If txtSampleTakenTime.Text = "" Then
180           SampleTaken = Format(DTSampleTaken, "dd/MM/yyyy")
190       Else
200           SampleTaken = Format(DTSampleTaken & " " & txtSampleTakenTime.FormattedText, "dd/MM/yyyy hh:mm:ss")
210       End If

220       SampleReceived = Format(DTSampleRec & " " & txtSampleRecTime.FormattedText, "dd/MM/yyyy hh:mm:ss")

230       PreReportDate = Format(lblPreReportDate, "dd/MM/yyyy hh:mm:ss")
240       ValReportDate = Format(lblValReportDate, "dd/MM/yyyy hh:mm:ss")

250       If (tb!Gross <> txtGross) Or _
             (tb!Micro <> txtMicro) Or _
             (tb!State <> State) Or _
             (tb!WithPathologist <> lblWithPathologist) Or _
             (tb!SampleTaken <> SampleTaken) Or _
             (tb!SampleReceived <> SampleReceived) Or _
             (tb!PreReportDate & "" <> PreReportDate) Or _
             (tb!ValReportDate & "" <> ValReportDate) Or _
             (tb!Preliminary <> optReport(0).Value) Or _
             (tb!Validated <> optReport(1).Value) Or _
             (tb!GeneralComments <> lblGeneralComments) Then

260           tbArc.AddNew
270           tbArc!CaseId = tb!CaseId
280           If tb!Gross <> txtGross Then
290               tbArc!Gross = tb!Gross
300           End If
310           If tb!Micro <> txtMicro Then
320               tbArc!Micro = tb!Micro
330           End If
340           tbArc!LinkedCaseId = tb!LinkedCaseId
350           tbArc!State = tb!State
360           tbArc!Phase = tb!Phase
370           tbArc!WithPathologist = tb!WithPathologist
380           tbArc!CheckedBy = tb!CheckedBy & ""
390           tbArc!SampleTaken = tb!SampleTaken
400           tbArc!SampleReceived = tb!SampleReceived
410           tbArc!PreReportDate = tb!PreReportDate
420           tbArc!ValReportDate = tb!ValReportDate
430           tbArc!Preliminary = tb!Preliminary
440           tbArc!Validated = tb!Validated
450           tbArc!PrintedVal = tb!PrintedVal
460           tbArc!OrigValDate = tb!OrigValDate
470           tbArc!OrigValBy = tb!OrigValBy
480           tbArc!UserName = tb!UserName
490           tbArc!ValidatedBy = tb!ValidatedBy
500           tbArc!GeneralComments = tb!GeneralComments
510           tbArc!Year = tb!Year
520           tbArc!UserName = tb!UserName
530           tbArc!DateTimeOfRecord = tb!DateTimeOfRecord
540           tbArc!ArchivedBy = UserName
550           tbArc.Update
560       End If

570       tb.MoveNext
580   Loop

590   Exit Sub

ArchiveCases_Error:

      Dim strES As String
      Dim intEL As Integer

600   intEL = Erl
610   strES = Err.Description
620   LogError "frmWorkSheet", "ArchiveCases", intEL, strES, sql

End Sub



Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim i As Integer
      Dim j As Integer
      Dim iIndex As Integer
      Dim tempState As String
      Dim lc As Recordset
      Dim cl As Recordset
      Dim bAddendumAdded As Boolean
      Dim strCaseId As String
      Dim blnOriginalAuthState As Boolean

10    On Error GoTo cmdSave_Click_Error

20    ClearExtraRequests = False

30    If txtCaseId = "" Then
40        frmMsgBox.Msg "Please Enter Case Number", , "Histology", mbExclamation
50        Exit Sub
60    End If
70    If DataMode = DataModeEdit And tvCaseDetails.SelectedItem Is Nothing Then
80        frmMsgBox.Msg "Please Select A Node First", , "Histology", mbExclamation
90        Exit Sub
100   End If

110   If DTSampleTaken.CustomFormat = " " Or DTSampleRec.CustomFormat = " " Or txtSampleRecTime = "" Then
120       frmMsgBox.Msg "PleaseFill In Mandatory Fields", mbOKOnly, "Histology", mbExclamation
130       Exit Sub
140   End If

150   If optState(1).Value = True And lblWithPathologist = "" Then
160       frmMsgBox.Msg "Please Select A Pathologist", mbOKOnly, "Histology", mbExclamation
170       Exit Sub
180   End If

190   cmdSave.Caption = "Saving..."

200   If lblPatientName.Caption <> "" Then

210       strCaseId = UCase(Replace(Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", ""), " ", ""))

220       If VerifyCaseIdFormat(strCaseId) Then

              'Save Demographics
230           sql = "Select * From Demographics Where CaseID = N'" & strCaseId & "'"
240           Set tb = New Recordset
250           RecOpenServer 0, tb, sql

              'Check to see if demographic entry or taken from PAS system
260           If Val(GetOptionSetting("DemographicEntry", "0")) <> 0 Then
270               If Not tb.EOF Then
280                   tb!ClinicalHistory = lblClinicalHist.Caption
290                   tb!NatureOfSpecimen = txtNOS
300                   tb!SpecimenLabelled = txtContainerLabel
310                   tb.Update
320               End If
330           Else
340               If tb.EOF Then tb.AddNew

350               tb!CaseId = strCaseId
360               tb!FirstName = lblFirstName.Caption
370               tb!Surname = lblSurname.Caption
380               tb!PatientName = lblPatientName.Caption
390               tb!Sex = Mid$(lblSex.Caption, 2, 1)
400               tb!Address1 = lblPatientAddress1.Caption
410               tb!Address2 = lblPatientAddress2.Caption
420               tb!Address3 = lblPatientAddress3.Caption
430               If Trim(Mid$(lblPatientBorn.Caption, 6)) <> "" Then
440                   tb!DateOfBirth = Mid$(lblPatientBorn.Caption, 6)
450               End If
460               tb!Age = lblAge.Caption
470               tb!Ward = lblPatientWard.Caption
480               tb!Clinician = lblPatientDoctor.Caption
490               tb!GP = Mid$(lblPatientGP.Caption, 5)
500               tb!Nopas = lblNopas.Caption
510               tb!MRN = lblMrn.Caption
520               tb!AandENo = lblAandE.Caption
530               tb!ClinicalHistory = lblClinicalHist.Caption
540               tb!NatureOfSpecimen = txtNOS
550               tb!SpecimenLabelled = txtContainerLabel
560               tb!UserName = UserName
570               tb.Update
580           End If


              'Save all data in the Cases table
590           sql = "Select * From Cases Where CaseID = N'" & strCaseId & "'"
600           Set tb = New Recordset
610           RecOpenServer 0, tb, sql

620           If tb.EOF Then
630               tb.AddNew
640           Else
650               ArchiveCases strCaseId
660           End If

670           tb!CaseId = strCaseId
680           If tb!Gross & "" = "" And txtGross.Text = "" Then
690               tb!Gross = ""
700           Else
710               If Replace(Replace(tb!Gross & "", vbLf, ""), vbCr, "") <> Replace(Replace(Replace(txtGross.TextRTF, "\par ", " \par "), vbLf, ""), vbCr, "") Then
720                   CaseAddLogEvent strCaseId, GrossEdited
730               End If
740               tb!Gross = Replace(txtGross.TextRTF, "\par ", " \par ")
750           End If


760           If Mid(strCaseId, 2, 1) <> "P" And Mid(CaseNo & "", 2, 1) <> "A" Then
770               If tb!Micro & "" = "" And txtMicro.Text = "" Then
780                   tb!Micro = ""
790               Else
800                   If Replace(Replace(tb!Micro & "", vbLf, ""), vbCr, "") <> Replace(Replace(Replace(txtMicro.TextRTF, "\par ", " \par "), vbLf, ""), vbCr, "") Then
810                       CaseAddLogEvent strCaseId, MicroEdited
820                   End If
830                   tb!Micro = Replace(txtMicro.TextRTF, "\par ", " \par ")
840               End If
850           End If

860           If optState(0).Value = True Then
870               tb!State = optState(0).Caption
880               If tb!Phase & "" = "" Then
890                   tb!Phase = "Cut-Up"
900               End If
910               bWithPathologist = False

920           ElseIf optState(1).Value = True Then
930               tempState = tb!State & ""
940               tb!State = optState(1).Caption
950               If tempState <> "With Pathologist" Then
960                   CaseAddLogEvent strCaseId, WithPathologist, lblWithPathologist & " (Checked By - " & lblCheckedBy & ")"
970                   ClearExtraRequests = True
980               End If
990               bWithPathologist = True
1000          ElseIf optState(2).Value = True Then
1010              tb!State = optState(2).Caption
1020              bWithPathologist = False

1030          Else
1040              tb!State = "Authorised"
1050              tb!AddendumAdded = False
1060              bWithPathologist = False
1070          End If
1080          tb!WithPathologist = lblWithPathologist
1090          If lblWithPathologist <> "" Then
1100              tb!AAPathologist = lblWithPathologist
1110          End If
              'if Q020, Q021, Q022 added then Awaiting Authoristion Tag will be set
1120          If optState(2).Tag <> "" Then
1130              tb!AAPathologist = optState(2).Tag    'Assign case back to user who added Q code
1140              optState(2).Tag = ""    'Clear tag
1150          End If
1160          tb!CheckedBy = lblCheckedBy

1170          For i = 1 To grdQCodes.Rows - 1
1180              If grdQCodes.TextMatrix(i, 0) = "Q020" _
                     Or grdQCodes.TextMatrix(i, 0) = "Q021" _
                     Or grdQCodes.TextMatrix(i, 0) = "Q022" Then

1190                  sql = "SELECT * FROM CaseListLink WHERE CaseId = N'" & strCaseId & "' " & _
                            "AND CaseListId = " & grdQCodes.TextMatrix(i, 2)
1200                  Set cl = New Recordset
1210                  RecOpenServer 0, cl, sql

1220                  If cl.EOF Then
1230                      bAddendumAdded = True
1240                      tb!AddendumAdded = bAddendumAdded
1250                      tb!DuplicatesPrinted = 0
1260                      Exit For
1270                  Else
1280                      bAddendumAdded = False
1290                  End If

1300              End If

1310          Next

1320          If (optReport(0).Value = True) And (tb!PreReportDate & "" = "") Then
1330              tb!PreReportDate = Format(Now, "yyyy-MM-dd hh:mm")
1340          ElseIf (optReport(1).Value = True) And (tb!OrigValDate & "" = "") Then
1350              tb!OrigValDate = Format(Now, "yyyy-MM-dd hh:mm")
1360              CaseWasOriginalValidated = True
1370              tb!OrigValBy = UserName
1380              tb!ValReportDate = Format(Now, "yyyy-MM-dd hh:mm")
1390              tb!ValidatedBy = UserName
1400          ElseIf (tb!Validated = 0) And (optReport(1).Value = True) Then
1410              tb!ValReportDate = Format(Now, "yyyy-MM-dd hh:mm")
1420              tb!ValidatedBy = UserName
1430          End If

1440          tb!Preliminary = optReport(0).Value

1450          If tb!Validated <> optReport(1).Value Then
1460              If optReport(1).Value = True Then
1470                  CaseAddLogEvent strCaseId, Authorised
1480              Else
1490                  CaseAddLogEvent strCaseId, UnAuthorised
1500              End If
1510          End If
1520          blnOriginalAuthState = IIf(tb!Validated, True, False)
1530          tb!Validated = optReport(1).Value

1540          tb!GeneralComments = lblGeneralComments.Caption

              'If there's a linked Histology Case Id to this Cytology Case
              'AND IF this Cytology Case id is been Authorised and was not previouslt Authorised THEN
              'Set the linked Histology case id to "Waiting Authorisation"
1550          If tb!LinkedCaseId & "" <> "" And UCase(Left(strCaseId, 1)) = "C" And UCase(Left(tb!LinkedCaseId & "", 1)) = "H" And tb!Validated And Not blnOriginalAuthState Then

1560              sql = "SELECT * FROM Cases WHERE CaseId = N'" & tb!LinkedCaseId & "'"
1570              Set lc = New Recordset
1580              RecOpenServer 0, lc, sql

1590              If Not lc.EOF Then
1600                  lc!State = "Awaiting Authorisation"
1610                  lc!AAPathologist = getUserCode(tb!ValidatedBy)
1620                  lc.Update
1630              End If
1640          End If

1650          If (optReport(1) = True) Or ((tb!OrigValDate & "" <> "") And tb!AddendumAdded = False) Then
1660              Validated = True
1670              DisableCase
1680              If UCase$(UserMemberOf) = "CONSULTANT" Or _
                     UCase$(UserMemberOf) = "SPECIALIST REGISTRAR" Then

1690                  txtQCode.Enabled = True
1700                  txtQDescription.Enabled = True
1710                  fraCaseState.Enabled = True
1720                  fraReport.Enabled = True
1730                  cmdSave.Enabled = True
1740              End If

1750              If tb!AddendumAdded = False _
                     And UCase$(UserMemberOf) <> "CLERICAL" Then
                      '1750                  fraCaseState.Enabled = True
1760                  cmdSave.Enabled = True
1770              End If
1780          Else
1790              Validated = False
1800              EnableCase
1810          End If


1820          If txtSampleTakenTime <> "" Then
1830              tb!SampleTaken = Format(DTSampleTaken & " " & txtSampleTakenTime.FormattedText, "dd/MMM/yyyy hh:mm")
1840          Else
1850              tb!SampleTaken = Format(DTSampleTaken, "dd/MMM/yyyy")
1860          End If
1870          tb!SampleReceived = Format(DTSampleRec & " " & txtSampleRecTime.FormattedText, "dd/MMM/yyyy hh:mm")

1880          tb!UserName = UserName
1890          tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm")
1900          tb.Update

              '******ADD CASE TREE DETAILS HERE
1910          If TreeChanged = True Or ClearExtraRequests = True Then

1920              i = 1
                  Dim xnode As MSComctlLib.Node
1930              Set xnode = tvCaseDetails.SelectedItem
1940              Do While (tvCaseDetails.Nodes(xnode.Index).Parent Is Nothing) = False
1950                  Set xnode = tvCaseDetails.Nodes(xnode.Index).Parent
1960              Loop
1970              iIndex = xnode.Index

1980              SaveTree tvCaseDetails.Nodes(iIndex), i

1990              If tvCaseDetails.Nodes(iIndex).Children > 0 Then
2000                  LocateNode iIndex, i
2010              End If

2020              While iIndex <> tvCaseDetails.Nodes(iIndex).LastSibling.Index
2030                  If tvCaseDetails.Nodes(iIndex).Next.Children > 0 Then
2040                      LocateNode tvCaseDetails.Nodes(iIndex).Next.Index, i
2050                  End If
2060                  iIndex = tvCaseDetails.Nodes(iIndex).Next.Index
2070              Wend
2080          End If

              '**********************

              '******ADD CODES HERE


2090          If txtPCode <> "" Then
2100              AddNewCaseList "P", GetListID(txtPCode, "P"), ""
2110              If CodeEventLog Then
2120                  CaseAddLogEvent strCaseId, PCodeEdited, "(" & txtPCode & " : " & txtPDescription & ")"
2130              End If
2140          End If
2150          For i = 1 To grdTempMCode.Rows - 1
2160              AddNewCaseList "M", GetListID(grdTempMCode.TextMatrix(i, 0), "M"), grdTempMCode.TextMatrix(i, 2), grdTempMCode.TextMatrix(i, 3), grdTempMCode.TextMatrix(i, 4), grdTempMCode.TextMatrix(i, 6)
2170              If CodeEventLog Then
2180                  CaseAddLogEvent strCaseId, MCodeAdded, "to " & grdTempMCode.TextMatrix(i, 5) & " (" & grdTempMCode.TextMatrix(i, 0) & " : " & grdTempMCode.TextMatrix(i, 1) & ")"
2190              End If

2200          Next i
2210          For i = 1 To grdQCodes.Rows - 1
2220              AddNewCaseList "Q", GetListID(grdQCodes.TextMatrix(i, 0), "Q"), grdQCodes.TextMatrix(i, 2)
2230              If CodeEventLog Then
2240                  CaseAddLogEvent strCaseId, QCodeAdded, "(" & grdQCodes.TextMatrix(i, 0) & " : " & grdQCodes.TextMatrix(i, 1) & ")"
2250              End If
2260          Next i

              '**********************

              '******ADD AMENDMENTS HERE


2270          For i = 1 To grdAmendments.Rows - 1
2280              AddNewCaseAmendment Replace(grdAmendments.TextMatrix(i, 1), vbCrLf, Chr(13)), grdAmendments.TextMatrix(i, 3), grdAmendments.TextMatrix(i, 2), grdAmendments.TextMatrix(i, 0)

2290          Next i

              '**********************

              '******ADD CASE MOVEMENTS HERE
2300          For j = 0 To 3
2310              For i = 1 To grdTracker(j).Rows - 1
2320                  With grdTracker(j)
2330                      .row = i
2340                      .col = 6
2350                      If .CellPicture = imgSquareCross.Picture Then
2360                          AddNewCaseMovement .TextMatrix(i, 7), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 8), .TextMatrix(i, 0), .TextMatrix(i, 9), .TextMatrix(i, 1), "0"
2370                      ElseIf .CellPicture = imgSquareTick.Picture Then
2380                          AddNewCaseMovement .TextMatrix(i, 7), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 8), .TextMatrix(i, 0), .TextMatrix(i, 9), .TextMatrix(i, 1), "1"
2390                      Else
2400                          AddNewCaseMovement .TextMatrix(i, 7), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 8), .TextMatrix(i, 0), .TextMatrix(i, 9), .TextMatrix(i, 1), ""
2410                      End If

2420                  End With
2430              Next i
2440          Next j


              '**********************

              '******DELETE RECORDS HERE

2450          For i = 1 To grdDelete.Rows - 1
2460              With grdDelete
2470                  DeleteRecords .TextMatrix(i, 0), .TextMatrix(i, 2)

2480              End With
2490          Next i
2500      Else
2510          frmMsgBox.Msg "CaseID Format Incorrect", , , mbExclamation
2520          cmdSave.Caption = "Save"
2530          Exit Sub
2540      End If
2550  Else
2560      frmMsgBox.Msg "No Demographic Available", , "Histology", mbExclamation
2570      cmdSave.Caption = "Save"
2580      Exit Sub
2590  End If

2600  For j = 0 To 3
2610      For i = 1 To grdTracker(j).Rows - 1
2620          If grdTracker(j).TextMatrix(i, 3) = "" Then
2630              Exit For
2640          Else
2650          End If
2660      Next i
2670  Next j

2680  If (optReport(0) = True Or optReport(1) = True) Then
2690      cmdPrnReport.Visible = True
2700  Else
2710      cmdPrnReport.Visible = False
2720  End If


2730  DataMode = DataModeEdit
2740  cmdSave.Caption = "Save"
2750  lblMessage.ZOrder 0
2760  lblMessage = "Record Saved"
2770  DataChanged = False
2780  TreeChanged = False

2790  If ClearExtraRequests = True Then
2800      FillTree
2810      ExpandAll tvCaseDetails
2820  End If
2830  lblRecordSaved.Visible = True

2840  Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer
2850  cmdSave.Caption = "Save"
2860  intEL = Erl
2870  strES = Err.Description
2880  LogError "frmWorkSheet", "cmdSave_Click", intEL, strES, sql

End Sub



Private Sub LocateNode(iNodeIndex As Integer, ByRef i As Integer)

      Dim n As Integer
      Dim iTempIndex As Integer

10    iTempIndex = tvCaseDetails.Nodes(iNodeIndex).Child.FirstSibling.Index

      'Loop through all a Parents Child Nodes
20    For n = 1 To tvCaseDetails.Nodes(iNodeIndex).Children

30        SaveTree tvCaseDetails.Nodes(iTempIndex), i

40        If tvCaseDetails.Nodes(iTempIndex).Children > 0 Then
50            LocateNode iTempIndex, i
60        End If

70        If n <> tvCaseDetails.Nodes(iNodeIndex).Children Then
80            iTempIndex = tvCaseDetails.Nodes(iTempIndex).Next.Index
90        End If
100   Next n

End Sub

Private Sub SaveTree(aNode As MSComctlLib.Node, ByRef i As Integer)
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveTree_Error

20    Set tb = New Recordset

30    sql = "SELECT * FROM CaseTree CT WHERE LocationID = N'" & Right(aNode.Key, Len(aNode.Key) - 2) & "' "

40    RecOpenServer 0, tb, sql

50    If tb.EOF Then
60        tb.AddNew

70        CaseUpdateLogEvent CaseNo, TreeNodeAdded, , aNode.FullPath
80    End If

90    If tb!LocationName <> aNode.Text Then
100       tb!UserName = UserName
110       tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")
120       CaseAddLogEvent CaseNo, TreeNodeEdited, " (CHANGED TO - " & aNode.Text & ")", tb!LocationName & ""
130   End If

140   tb!CaseId = CaseNo
150   tb!LocationID = Right(aNode.Key, Len(aNode.Key) - 2)
160   tb!LocationName = aNode.Text
170   If Left(aNode.Key, 2) = "L0" Then
180       tb!LocationParentID = 0
190   Else
200       tb!LocationParentID = Right(aNode.Parent.Key, Len(aNode.Parent.Key) - 2)
210   End If
220   If Left(aNode.Key, 2) = "L0" Then
230       tb!LocationLevel = 0
240   Else
          tb!LocationLevel = GetNodeLevel(Right(aNode.Parent.Key, Len(aNode.Parent.Key) - 2)) + 1

270   End If

280   If Left(aNode.Key, 2) = "L1" Then
290       tb!TissueTypeListId = aNode.Tag
300       If Right(aNode.Text, 4) = "(AE)" Then
310           tb!AllEmbedded = 1
320       Else
330           tb!AllEmbedded = 0
340       End If
350   End If


360   If InStr(1, aNode.Text, "Block") Then
370       tb!Type = "B"
380       tb!NodeNumber = IIf(Trim(Mid(aNode.Text, InStr(aNode.Text, " "))) = "", "A", Trim(Mid(aNode.Text, InStr(aNode.Text, " "))))
390       If InStr(1, aNode.Parent.Text, "Frozen Section") Then
400           AddBlockDetails aNode, aNode.Parent.Parent
410           tb!LocationSpecimenID = Right(aNode.Parent.Parent.Key, Len(aNode.Parent.Parent.Key) - 2)
420       Else
430           AddBlockDetails aNode, aNode.Parent
440           tb!LocationSpecimenID = Right(aNode.Parent.Key, Len(aNode.Parent.Key) - 2)
450       End If
460       If ClearExtraRequests = True Then
470           tb!ExtraRequests = "0"
480       Else
490           tb!ExtraRequests = aNode.Tag
500       End If
510       tb!TotalLevelRequests = Val(tb!TotalLevelRequests & "") + Val(aNode.Tag)
520   ElseIf InStr(1, aNode.Text, "Slide") Then
530       tb!Type = "S"
540       tb!NodeNumber = Trim(Mid(aNode.Text, InStr(aNode.Text, " ")))
550       tb!NoOfSections = Val(aNode.Tag)
560   Else
570       If aNode.ForeColor = vbBlue Then
580           If ClearExtraRequests = True Then
590               tb!ExtraRequests = "0"
600           Else
610               tb!ExtraRequests = "S"
620           End If
630       End If
640   End If


650   tb!TreeOrder = i
660   i = i + 1
670   tb!LocationPath = aNode.FullPath
      If GetNodeLevel(Right(aNode.Parent.Key, Len(aNode.Parent.Key) - 2)) + 1 = 1 Then
        tb!Type = "TS"
      End If
680   tb.Update

690   Exit Sub

SaveTree_Error:

      Dim strES As String
      Dim intEL As Integer

700   intEL = Erl
710   strES = Err.Description
720   LogError "frmWorkSheet", "SaveTree", intEL, strES, sql

End Sub



Private Sub AddNewCaseMovement(sType As String, Code As String, _
                               UniqueId As String, _
                               DateSent As String, _
                               DateReceived As String, _
                               Destination As String, _
                               Description As String, _
                               ReferralReason As String, _
                               SpecId As String, _
                               Agreed As String)
      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo AddNewCaseMovement_Error

20    sql = "Select * From CaseMovements Where CaseListId = N'" & UniqueId & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then tb.AddNew

60    tb!CaseId = CaseNo
70    tb!CaseListId = UniqueId
80    tb!Code = Code
90    tb!Type = sType
100   tb!SpecId = SpecId
110   tb!Destination = Destination
120   If DateSent <> "" Then
130       tb!DateSent = DateSent
140   End If
150   If DateReceived <> "" Then
160       tb!DateReceived = DateReceived
170   End If
180   tb!Description = Description
190   tb!ReferralReason = ReferralReason
200   tb!Agreed = Agreed
210   tb!UserName = UserName
220   tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")

230   tb.Update

240   Exit Sub

AddNewCaseMovement_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmWorkSheet", "AddNewCaseMovement", intEL, strES, sql


End Sub

Private Sub AddNewCaseAmendment(Comment As String, Code As String, UniqueId As String, DateTime As String)
      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo AddNewCaseAmendment_Error

20    sql = "Select * From CaseAmendments Where CaseListId = N'" & UniqueId & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then tb.AddNew

60    If tb!Comment & "" <> Comment Then
70        ArchiveAmendment UniqueId
80        tb!UserName = UserName
90    End If
100   tb!CaseId = CaseNo
110   tb!CaseListId = UniqueId
120   tb!Code = Code
130   tb!Comment = Comment
140   If Not tb!Valid Then
150       tb!Valid = optReport(1).Value
160   End If
170   If DateTime <> "" Then
180       tb!DateTimeOfRecord = DateTime
190   End If
200   tb.Update

210   Exit Sub

AddNewCaseAmendment_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmWorkSheet", "AddNewCaseAmendment", intEL, strES, sql


End Sub

Private Sub ArchiveAmendment(ByVal Id As String)

      Dim tb As Recordset
      Dim tbArc As Recordset
      Dim f As Field
      Dim sql As String

10    On Error GoTo ArchiveAmendment_Error

20    sql = "SELECT * FROM CaseAmendments WHERE " & _
            "CaseListID = N'" & Id & "'"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    sql = "SELECT * FROM CaseAmendmentsAudit WHERE 0 = 1"
60    Set tbArc = New Recordset
70    RecOpenServer 0, tbArc, sql

80    Do While Not tb.EOF

90        tbArc.AddNew
100       For Each f In tb.Fields
110           tbArc(f.name) = tb(f.name)
120       Next
130       tbArc!ArchivedBy = UserName
140       tbArc.Update

150       tb.MoveNext
160   Loop

170   Exit Sub

ArchiveAmendment_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmWorkSheet", "ArchiveAmendment", intEL, strES, sql

End Sub

Private Sub AddNewCaseList(ListType As String, ListId As Integer, _
                           UniqueId As String, _
                           Optional TissueTypeId As String, _
                           Optional TissueTypeLetter As String, _
                           Optional TissueTypeListId As String)
      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo AddNewCaseList_Error

20    If ListType = "P" Then
30        sql = "SELECT * From CaseListLink WHERE CaseId = N'" & CaseNo & "' AND Type  = N'" & ListType & "'"
40    Else
50        sql = "Select * From CaseListLink Where CaseListId = N'" & UniqueId & "'"
60    End If
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If tb.EOF Then
100       tb.AddNew
110       tb!UserName = UserName
120       tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")
130       UpdateListRank ListId
140   End If
150   If ListType = "P" Then
160       If tb!ListId <> ListId Then
170           CodeEventLog = True
180       Else
190           CodeEventLog = False
200       End If
210   Else
220       If tb!CaseListId <> UniqueId Then
230           CodeEventLog = True
240       Else
250           CodeEventLog = False
260       End If
270   End If

280   tb!ListId = ListId
290   tb!CaseListId = UniqueId
300   tb!CaseId = CaseNo
310   tb!Type = ListType
320   If TissueTypeId <> "" Then
330       tb!TissueTypeId = TissueTypeId
340   End If
350   If TissueTypeLetter <> "" Then
360       tb!TissueTypeLetter = TissueTypeLetter
370   End If
380   If TissueTypeListId <> "" Then
390       tb!TissueTypeListId = TissueTypeListId
400   End If


      '390     tb!UserName = UserName
      '400   tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")

410   tb.Update

420   Exit Sub

AddNewCaseList_Error:

      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "frmWorkSheet", "AddNewCaseList", intEL, strES, sql


End Sub

Private Sub DeleteRecords(UniqueId As String, Table As String)
      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo DeleteRecords_Error

20    sql = "Select * From " & Table & " Where CaseListId = N'" & UniqueId & "'"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then tb.Delete

60    Exit Sub

DeleteRecords_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "frmWorkSheet", "DeleteRecords", intEL, strES, sql


End Sub


Private Sub AddBlockDetails(blockNode As MSComctlLib.Node, TissueNode As MSComctlLib.Node)
      Dim sql As String
      Dim tb As Recordset
      Dim BlockExists As Boolean

10    On Error GoTo AddBlockDetails_Error

20    BlockExists = True
30    sql = "SELECT * FROM BlockDetails " & _
            "WHERE CaseId = N'" & CaseNo & "' " & _
            "AND TissueListId = N'" & Right(TissueNode.Key, Len(TissueNode.Key) - 2) & "' " & _
            "AND UniqueValue = N'" & Left(TissueNode.Text, 1) & "' "

40    If InStr(blockNode.Text, "Block") Then
50        sql = sql & "AND BlockNumber = N'" & IIf(Trim(Mid(blockNode.Text, InStr(blockNode.Text, " "))) = "", "A", Trim(Mid(blockNode.Text, InStr(blockNode.Text, " ")))) & "'"
60    Else
70        BlockExists = False
80    End If

90    If BlockExists Then
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql


120       If tb.EOF Then
130           tb.AddNew
140       End If
150       tb!CaseId = CaseNo
160       tb!Tissuelistid = Right(TissueNode.Key, Len(TissueNode.Key) - 2)
170       tb!Block = blockNode.Text
180       tb!BlockNumber = IIf(Trim(Mid(blockNode.Text, InStr(blockNode.Text, " "))) = "", "A", Trim(Mid(blockNode.Text, InStr(blockNode.Text, " "))))
190       tb!UniqueValue = Left(TissueNode.Text, 1)
200       tb!LocationPath = blockNode.FullPath
210       tb.Update
220   End If



230   Exit Sub

AddBlockDetails_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmWorkSheet", "AddBlockDetails", intEL, strES, sql


End Sub

Private Sub cmdSearch_Click()
10    If DataChanged = False Then
20        If txtCaseId = "" Then
30            ResetWorkSheet


40            If FillTree Then
50                DataMode = DataModeEdit
60                Search = True
70            End If
80        End If
90    Else

100       If frmMsgBox.Msg("Save Changes", mbYesNo, , mbQuestion) = 1 Then
110           cmdSave_Click
120       End If

130       If txtCaseId = "" Then
140           ResetWorkSheet


150           If FillTree Then
160               DataMode = DataModeEdit
170               DataChanged = False
180               TreeChanged = False
190               Search = True
200           End If
210       End If
220   End If
230   PatientIdCombo = cmbPatientId
240   PatientIdText = txtPatientId
End Sub

Private Sub Form_Activate()
CheckIfCaseIdExist
'10    frmWorkSheet_ChangeLanguage
20    If txtCaseId <> "" Then
30        CaseNo = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")

40        sCaseLockedBy = CaseLockedBy(CaseNo)

50        If sCaseLockedBy <> UserName And sCaseLockedBy <> "" Then
60            lblCaseLocked = "RECORD BEING EDITED BY" & " " & sCaseLockedBy
70            bLocked = True
80            lblCaseLocked.BackColor = &H8080FF
90        ElseIf sCaseLockedBy = "" Then
100           LockCase CaseNo
110           bLocked = False
120           UnLockControl
130           lblCaseLocked.BackColor = &H80FF80
140           lblCaseLocked = "RECORD BEING EDITED BY YOU!"
150       Else
160           bLocked = False
170           UnLockControl
180           lblCaseLocked.BackColor = &H80FF80
190           lblCaseLocked = "RECORD BEING EDITED BY YOU!"
200       End If

210       If Mid(txtCaseId, 2, 1) = "P" Or Mid(txtCaseId, 2, 1) = "A" Then
220           lngMaxDigits = 12
230           txtMicro.Visible = False
240           txtGross.Height = 4000
250       Else
260           lngMaxDigits = 11
270           txtMicro.Visible = True
280           txtGross.Height = 1815
290       End If
300       If DataMode = 0 Then
310           If FillTree Then
320               DataMode = DataModeEdit
330               tvCaseDetails.Nodes(1).Selected = True
340               LoadDemographics CaseNo
350               FillWorkSheet CaseNo
360           End If
370       End If


380   Else
390       cmdClinicalHist.Visible = False
400       cmdComments.Visible = False
410       cmdDiscrepancyLog.Visible = False
420       cmdAudit.Visible = False
430       fraDemographics.Visible = False
440       cmdEditDemo.Visible = False
450       cmdCytoHist.Visible = False
460       fraWorkSheet.Enabled = False
470       txtCaseId.SetFocus
480   End If


490   If Left(txtCaseId, 1) = "C" Or _
         Mid(txtCaseId, 2, 1) = "A" Then
500       cmdCytoHist.Enabled = True
510   Else
520       cmdCytoHist.Enabled = False
530   End If
540   pBar = 0

550   Select Case UCase$(UserMemberOf)
      Case "CLERICAL"
560       DisableClerical
570   Case "SCIENTIST"
580       DisableScientist
590   Case "MANAGER"
600       DisableManager
610   Case "CONSULTANT"
620       DisableConsultant
630   Case "SPECIALIST REGISTRAR"
640       DisableConsultant
650   End Select

660   If bLocked Then
670       LockControl
680   End If
'ALI
 SSTabMovement.Tab = 2
'------

690   Activated = True
End Sub



Private Sub DisableClerical()
      Dim i As Integer

10    cmdCytoHist.Enabled = False
20    txtNOS.Locked = True
30    txtContainerLabel.Locked = True
40    txtPCode.Enabled = False
50    txtPDescription.Enabled = False
60    txtMCode.Enabled = False
70    txtMDescription.Enabled = False
80    txtQCode.Enabled = False
90    txtQDescription.Enabled = False
100   grdMCodes.Enabled = False
110   grdQCodes.Enabled = False
120   cmdMCode.Enabled = False
130   cmdQCode.Enabled = False
140   For i = 0 To 3
150       grdTracker(i).Enabled = False
160   Next
170   fraCaseState.Enabled = False
180   fraReport.Enabled = False
190   optReport(0).Enabled = False
200   optReport(1).Enabled = False
210   txtCaseId.Enabled = False
220   mnuDelTissueType.Visible = False
230   mnuDelBlock.Visible = False
240   mnuDelSlide.Visible = False
250   mnuSeperator8.Visible = False
260   mnuSeperator7.Visible = False
270   mnuSeperator3.Visible = False

End Sub

Private Sub DisableScientist()
10    fraReport.Enabled = False
20    grdQCodes.Enabled = False
30    mnuDelTissueType.Visible = False
40    mnuDelBlock.Visible = False
50    mnuDelSlide.Visible = False
60    mnuSeperator8.Visible = False
70    mnuSeperator7.Visible = False
80    mnuSeperator3.Visible = False
End Sub

Private Sub DisableManager()
10    fraReport.Enabled = False
End Sub

Private Sub DisableConsultant()
10    mnuSingleSlideLevel3.Visible = False
20    mnuMultipleSlidesLevel3.Visible = False
30    mnuPrnBlockNumber.Visible = False
40    mnuDelTissueType.Visible = False
50    mnuDelBlock.Visible = False
60    mnuDelSlide.Visible = False
70    mnuSeperator8.Visible = False
80    mnuSeperator7.Visible = False
90    mnuSeperator3.Visible = False
'
'100   txtMCode.Enabled = True
'110   txtMDescription.Enabled = True
'120   cmdMCode.Enabled = True
      

End Sub


Private Sub Form_Resize()
10    If Me.WindowState <> vbMinimized Then

20        Me.Top = 0
30        Me.Left = Screen.Width / 2 - Me.Width / 2
40    End If
End Sub

Private Sub Form_Load()


10    On Error GoTo Form_Load_Error
20
30    Activated = False

40    ChangeFont Me, "Arial"
50    cmbPatientId.AddItem "NOPAS"
60    cmbPatientId.AddItem "MRN"
70    cmbPatientId.AddItem "A&E No"
80    PopulateGenericList "Status", cmbState
90    cmbPatientId.Text = "MRN"

100   lblLoggedIn = UserName

110   ResetWorkSheet

120   lngMaxDigits = 11
130   loadtooltip
140   strReportPath = App.Path & "\"


150   Me.Caption = "NetAcquire - Cellular Pathology. Version " & strVersion

160   If blnIsTestMode Then EnableTestMode Me
      'Zyam 24-07-24
'170   If UCase(UserMemberOf) = "CONSULTANT" Then
'180     txtMCode.Enabled = True
'190     txtMDescription.Enabled = True
'200     cmdMCode.Enabled = True
'210   End If
      'Zyam 24-07-24

220   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmWorkSheet", "Form_Load", intEL, strES


End Sub
Private Sub CheckIfCaseIdExist()
          Dim checkcaseId As Integer
          Dim sql As String
          Dim tb As Recordset
          Dim Count As Integer
          Dim CaseId As String
          
10        CaseId = Replace(txtCaseId.Text, "/", "")
20        CaseId = Replace(CaseId, " ", "")

30        On Error GoTo CheckIfCaseIdExist_Error

40

50        sql = "SELECT COUNT(*) AS RecordCount FROM VentanaStatusUpdate WHERE CaseID = '" & Trim(CaseId) & "'"

60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql

80        If Not tb.EOF Then
90            Count = tb!RecordCount
100           If Count = 1 Then
110               cmdVantage.BackColor = vbGreen
120           Else
130               cmdVantage.BackColor = vbRed
140           End If
150       End If

160       tb.Close
170       Set tb = Nothing

180       Exit Sub

CheckIfCaseIdExist_Error:
          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmCheckCaseId", "CheckIfCaseIdExist", intEL, strES, sql
End Sub


Private Sub loadtooltip()

      Dim i As Integer

10    With mclsToolTip
          '
          ' Create the tooltip window.
          '
20        Call .Create(Me)
          '
          ' Set the tooltip's width so that it displays
          ' multiline text and no tool's line length exceeds
          ' roughly 240 pixels.
          '
30        .MaxTipWidth = 240
          '
          ' Show the tooltip for 20 seconds.
          '
          '.DelayTime(ttDelayShow) = 20000
          '
          'MODIFICATION
          'Create a standard header for all the tooltips (note the class adds a vbNewLine to end so you don't have to
          '.ToolTipHeader = "CLSTOOLTIP DEMO" & vbNewLine & "_______________"
          '
          ' Add a tooltip tool to each control on the Form.
          '

40        Call .AddTool(grdAmendments)
50        For i = 0 To 3
60            Call .AddTool(grdTracker(i))
70        Next
80        Call .AddTool(cmdClinicalHist)
90        Call .AddTool(cmdComments)
100       Call .AddTool(fraDemographics)
110   End With
End Sub


Public Sub LoadDemographics(CaseId As String)
      Dim sql As String
      Dim tb As New Recordset
      Dim s As String

10    On Error GoTo LoadDemographics_Error

20    sql = "SELECT * FROM Demographics d " & _
            "LEFT JOIN Cases c ON d.Caseid = c.caseid " & _
            "LEFT JOIN CaseListLink cl ON c.caseid = cl.caseid " & _
            "LEFT JOIN Lists l ON cl.ListId = l.ListId " & _
            "WHERE d.CaseId = N'" & CaseId & "'"
30    If txtPatientId <> "" Then
40        If cmbPatientId = "NOPAS" Then
50            sql = sql & "AND Nopas = N'" & txtPatientId & "' "
60        ElseIf cmbPatientId = "MRN" Then
70            sql = sql & "AND Mrn = N'" & txtPatientId & "' "
80        ElseIf cmbPatientId = "A&E No" Then
90            sql = sql & "AND AandENo = N'" & txtPatientId & "' "
100       End If
110   End If
120   Set tb = New Recordset
130   RecOpenClient 0, tb, sql

140   If Not tb.EOF Then
150       fraDemographics.Visible = True
160       lblFirstName.Caption = tb!FirstName & ""
170       lblSurname.Caption = tb!Surname & ""
180       lblPatientName.Caption = tb!PatientName & ""
190       If tb!Sex <> "" Then
200           lblSex.Caption = "(" & tb!Sex & ")"
210       End If
220       lblPatientAddress1.Caption = tb!Address1 & ""
230       lblPatientAddress2.Caption = tb!Address2 & ""
240       lblPatientAddress3.Caption = tb!Address3 & ""
250       If tb!DateOfBirth <> "" Then
260           lblPatientBorn.Caption = "Born " & tb!DateOfBirth & " "
270       End If
280       If tb!Age <> "" Then
290           lblAge.Caption = tb!Age
300       End If
310       lblPatientWard.Caption = tb!Ward & ""
320       If tb!AutopsyFor & "" <> "" Then
330           lblPatientDoctor.Caption = tb!AutopsyRequestedBy & ""
340           lblDOD = tb!DateOfDeath & ""
350       Else
360           lblPatientDoctor.Caption = tb!Clinician & ""
370       End If
380       If tb!GP <> "" Then
390           lblPatientGP.Caption = "GP: " & tb!GP & ""
400       End If
410       lblNopas.Caption = tb!Nopas & ""
420       lblMrn.Caption = tb!MRN & ""
430       lblAandE.Caption = tb!AandENo & ""
440       txtNOS = tb!NatureOfSpecimen & ""
450       txtContainerLabel = tb!SpecimenLabelled & ""
460       lblClinicalHist = tb!ClinicalHistory & ""
470       CheckClinicalHist
480       CheckDiscrepancyLog
490       mclsToolTip.ToolText(cmdClinicalHist) = lblClinicalHist

500       If cmbPatientId = "NOPAS" Then
510           txtPatientId.Text = lblNopas
520       ElseIf cmbPatientId = "MRN" Then
530           txtPatientId.Text = lblMrn
540       ElseIf cmbPatientId = "A&E No" Then
550           txtPatientId.Text = lblAandE
560       End If
570       cmdClinicalHist.Visible = True

580       cmdDiscrepancyLog.Visible = True

590       If Val(GetOptionSetting("DemographicEntry", "0")) <> 0 Then
600           cmdEditDemo.Visible = True
610       End If
620       cmdCytoHist.Visible = True
630       cmdAudit.Visible = True

640       s = "Chart No:  " & lblMrn & vbCrLf & _
              "Name:      " & lblPatientName & vbCrLf & _
              "Sex:       " & lblSex & vbCrLf & _
              "Address1:  " & lblPatientAddress1 & vbCrLf & _
              "Address2:  " & lblPatientAddress2 & vbCrLf & _
              "Address3:  " & lblPatientAddress3 & vbCrLf & _
              "DOB:       " & lblPatientBorn & vbCrLf & _
              "Ward:      " & lblPatientWard & vbCrLf
650       If tb!AutopsyFor & "" <> "" Then
660           s = s & tb!AutopsyFor & ": " & lblPatientDoctor & vbCrLf
670           s = s & "Date Of Death: " & lblDOD & vbCrLf
680       Else
690           s = s & "Clinician: " & lblPatientDoctor & vbCrLf
700       End If
710       s = s & "GP:        " & lblPatientGP
720       mclsToolTip.ToolText(fraDemographics) = s
730   Else
740       txtPatientId.Text = ""

750   End If
760   cmdSearch.Enabled = False

770   Exit Sub

LoadDemographics_Error:

      Dim strES As String
      Dim intEL As Integer

780   intEL = Erl
790   strES = Err.Description
800   LogError "frmWorkSheet", "LoadDemographics", intEL, strES, sql

End Sub


Private Sub FillAllCodes(CaseId As String)

      Dim tb As Recordset
      Dim sql As String
      Dim i As Integer

10    On Error GoTo FillAllCodes_Error

20    sql = "Select L.Code, L.Description, L.ListType, C.CaseListId, " & _
            "C.TissueTypeId, C.TissueTypeLetter, C.TissueTypeListId " & _
            "From CaseListLink C " & _
            "Inner Join Lists L " & _
            "On C.ListID = L.ListID " & _
            "Where C.CaseID = N'" & CaseId & "'"


30    Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        While Not tb.EOF
70            Select Case tb!ListType & ""
              Case "P":
80                txtPCode = tb!Code & ""
90                txtPDescription = tb!Description & ""
100           Case "M":
110               grdTempMCode.AddItem tb!Code & "" _
                                       & vbTab & tb!Description & "" _
                                       & vbTab & tb!CaseListId & "" _
                                       & vbTab & tb!TissueTypeId & "" _
                                       & vbTab & tb!TissueTypeLetter & "" _
                                       & vbTab & "" _
                                       & vbTab & tb!TissueTypeListId & "", _
                                       grdTempMCode.Rows
120           Case "Q"
130               grdQCodes.AddItem tb!Code & "" _
                                    & vbTab & tb!Description & "" _
                                    & vbTab & tb!CaseListId & "", _
                                    grdQCodes.Rows
140           End Select
150           tb.MoveNext
160       Wend
170   End If

180   For i = 1 To grdQCodes.Rows - 1
190       If grdQCodes.TextMatrix(i, 0) = "Q021" Then
200           With grdQCodes
210               .row = i
220               .col = 0
230               .CellForeColor = vbRed
240               .col = 1
250               .CellForeColor = vbRed
260           End With
270       End If
280   Next

290   Exit Sub

FillAllCodes_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "frmWorkSheet", "FillAllCodes", intEL, strES, sql

End Sub

Private Sub FillAmendments(CaseId As String)

      Dim tb As Recordset
      Dim sql As String
      Dim i As Integer


10    On Error GoTo FillAmendments_Error

20    sql = "Select * " & _
            "From CaseAmendments C " & _
            "Where C.CaseID = N'" & CaseId & "' " & _
            "order by c.datetimeofrecord"  'see ITS 818948 (1)

30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        While Not tb.EOF
70            grdAmendments.AddItem Format(tb!DateTimeOfRecord & "", "dd/mm/yy hh:mm") & vbTab & tb!Comment & "" & vbTab & tb!CaseListId & "" & vbTab & tb!Code & "", grdAmendments.Rows
80            tb.MoveNext
90        Wend
100   End If

110   For i = 1 To grdAmendments.Rows - 1
120       If grdAmendments.TextMatrix(i, 3) = "Q021" Then
130           With grdAmendments
140               .row = i
150               .col = 0
160               .CellForeColor = vbRed
170               .col = 1
180               .CellForeColor = vbRed
190           End With
200       End If
210   Next


220   Exit Sub

FillAmendments_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmWorkSheet", "FillAmendments", intEL, strES, sql


End Sub

Private Sub FillWorkSheet(CaseId As String)

      Dim tb As Recordset
      Dim sn As Recordset
      Dim rsRec As Recordset
      Dim sql As String
      Dim i As Integer
      Dim j As Integer

10    On Error GoTo FillWorkSheet_Error


20    sql = "Select * From Cases Where CaseID = N'" & CaseId & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    If Not tb.EOF Then
          '60     txtGross.TextRTF = tb!Gross & ""
          '70      txtMicro.TextRTF = tb!Micro & ""
60        txtGross.TextRTF = Replace(tb!Gross & "", " \par ", "\par ")
70        txtMicro.TextRTF = Replace(tb!Micro & "", " \par ", "\par ")
80        If Not IsNull(tb!SampleTaken) Then
90            DTSampleTaken.CustomFormat = "dd/MM/yyyy"
100           DTSampleTaken.Value = Format(tb!SampleTaken, "dd/MMM/yyyy")
110           If Format$(tb!SampleTaken, "hh:mm") <> "00:00" Then
120               txtSampleTakenTime.SelText = Format(tb!SampleTaken, "hh:mm")
130           End If
140       End If
150       If Not IsNull(tb!SampleReceived) Then
160           DTSampleRec.CustomFormat = "dd/MM/yyyy"
170           DTSampleRec.Value = Format(tb!SampleReceived, "dd/MMM/yyyy")
180           txtSampleRecTime.SelText = Format(tb!SampleReceived & "", "hh:mm")
190       End If

200       If tb!LinkedCaseId & "" <> "" Then
210           fraLinkedCase.Visible = True
220           If Len(tb!LinkedCaseId) = 9 Then
230               cmdLinkedCaseId.Caption = Left(tb!LinkedCaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!LinkedCaseId, 2)
240           Else
250               cmdLinkedCaseId.Caption = Left(tb!LinkedCaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!LinkedCaseId, 2)
260           End If
270           cmdCytoHist.Enabled = False
280           End If

290       If UCase(Trim(tb!State) & "") = UCase("In Histology") Then
300           optState(0).Value = True
310           bWithPathologist = False
320       ElseIf UCase(Trim(tb!State) & "") = UCase("With Pathologist") Then
330           optState(1).Value = True
340           bWithPathologist = True
350       ElseIf UCase(Trim(tb!State) & "") = UCase("Awaiting Authorisation") Then
360           optState(2).Value = True
370           bWithPathologist = False
380       End If

390       sql = "Select * From Users Where Code = N'" & tb!WithPathologist & "'"
400       Set rsRec = New Recordset
410       RecOpenServer 0, rsRec, sql

420       If Not rsRec.EOF Then
430           lblWithPathologistName = rsRec!UserName & ""
440       End If

450       lblWithPathologist = tb!WithPathologist & ""
460       lblCheckedBy = tb!CheckedBy & ""

470       If (tb!Validated Or tb!Preliminary) Then
480           txtPatientId.Enabled = False
490       End If

500       If Not IsNull(tb!PreReportDate) Then
510           lblPreReportDate = Format(tb!PreReportDate & "", "dd/MM/yyyy hh:mm")
520       End If

530       If Not IsNull(tb!ValReportDate) Then
540           lblValReportDate = Format(tb!ValReportDate & "", "dd/MM/yyyy hh:mm")
550       End If
560       lblGeneralComments = tb!GeneralComments & ""
570       CheckGeneralComments


580       mclsToolTip.ToolText(cmdComments) = lblGeneralComments

590       cmdComments.Visible = True

600       FillAllCodes CaseId
610       FillTracker
620       FillAmendments CaseId
630       For j = 0 To 3
640           If grdTracker(j).Rows > 1 Then
650               For i = 1 To grdTracker(j).Rows - 1
660                   If UCase(grdTracker(j).TextMatrix(i, 3)) = "" Then
670                       Exit For
680                   End If
690               Next i
700           Else
710           End If
720       Next
730   End If

740   If tb!OrigValDate & "" <> "" Then
750       Validated = True
760       CaseWasOriginalValidated = True
770   Else
780       Validated = False
790       CaseWasOriginalValidated = False
800   End If
810   If Validated Then

820       DisableCase
830       If UCase$(UserMemberOf) = "CONSULTANT" Or _
             UCase$(UserMemberOf) = "SPECIALIST REGISTRAR" Then

840           txtQCode.Enabled = True
850           txtQDescription.Enabled = True
860           fraCaseState.Enabled = True
870           fraReport.Enabled = True
880           cmdSave.Enabled = True
890       ElseIf UCase$(UserMemberOf) = "MANAGER" Then
900           cmdEditDemo.Enabled = True
910           txtQCode.Enabled = True
920           txtQDescription.Enabled = True
930       ElseIf UCase$(UserMemberOf) = "SCIENTIST" Then
940           txtQCode.Enabled = True
950           txtQDescription.Enabled = True
960       ElseIf UCase$(UserMemberOf) = "IT MANAGER" Then
970           txtQCode.Enabled = True
980           txtQDescription.Enabled = True
990           fraCaseState.Enabled = True
1000          fraReport.Enabled = True
1010          cmdSave.Enabled = True
1020          cmdEditDemo.Enabled = True
1030      End If
1040  Else
1050      EnableCase
1060  End If

1070  If tb!Validated = False Then
1080      If tb!AddendumAdded = True Then
1090          If UCase$(UserMemberOf) <> "CLERICAL" Then
1100              tvCaseDetails.Enabled = True

1110          End If
1120          fraCaseState.Enabled = True
1130          cmdSave.Enabled = True
1140          lblAddendumAdded = "TRUE"
1150      Else
1160          fraCaseState.Enabled = True
1170          cmdSave.Enabled = True
1180          lblAddendumAdded = "FALSE"
1190      End If
1200  End If



1210  optReport(1).Value = IIf(IsNull(tb!Validated), 0, tb!Validated)

1220  optReport(0).Value = IIf(IsNull(tb!Preliminary), 0, tb!Preliminary)

1230  If (optReport(0) = True Or optReport(1) = True) Then
1240      cmdPrnReport.Visible = True
1250  End If

1260  sql = "Select * From Reports Where Sampleid = N'" & CaseId & "'"
1270  Set sn = New Recordset
1280  RecOpenServer 0, sn, sql

1290  If Not sn.EOF Then
1300      cmdViewReports.Visible = True
1310  End If


1320  DataChanged = False


1330  Exit Sub

FillWorkSheet_Error:

      Dim strES As String
      Dim intEL As Integer

1340  intEL = Erl
1350  strES = Err.Description
1360  LogError "frmWorkSheet", "FillWorkSheet", intEL, strES, sql

End Sub

Public Sub CheckLinkedCase(CaseId As String)

10    fraLinkedCase.Visible = True
20    If Len(CaseId) = 9 Then
30        cmdLinkedCaseId.Caption = Left(CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(CaseId, 2)
40    Else
50        cmdLinkedCaseId.Caption = Left(CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(CaseId, 2)
60    End If
70    cmdCytoHist.Enabled = False

End Sub

Private Sub DisableCase()
10    txtGross.Locked = True
20    txtMicro.Locked = True
30    DTSampleTaken.Enabled = False
40    txtSampleTakenTime.Enabled = False
50    DTSampleRec.Enabled = False
60    txtSampleRecTime.Enabled = False
70    txtPCode.Enabled = False
80    txtPDescription.Enabled = False
90    txtMCode.Enabled = False
100   txtMDescription.Enabled = False
110   txtQCode.Enabled = False
120   txtQDescription.Enabled = False
130   fraCaseState.Enabled = False
140   fraReport.Enabled = False
150   grdMCodes.Enabled = False
160   cmdEditDemo.Enabled = False
End Sub

Private Sub EnableCase()
10    txtGross.Locked = False
20    txtMicro.Locked = False
30    DTSampleTaken.Enabled = True
40    txtSampleTakenTime.Enabled = True
50    DTSampleRec.Enabled = True
60    txtSampleRecTime.Enabled = True

70    If UCase$(UserMemberOf) <> "CLERICAL" Then
80        txtPCode.Enabled = True
90        txtPDescription.Enabled = True
100       fraCaseState.Enabled = True
110       grdMCodes.Enabled = True
120   End If

130   If UCase$(UserMemberOf) = "CONSULTANT" Or _
         UCase$(UserMemberOf) = "SPECIALIST REGISTRAR" Then
140       fraReport.Enabled = True
150   End If

160   If UCase$(UserMemberOf) = "MANAGER" Or _
         UCase$(UserMemberOf) = "CONSULTANT" Or _
         UCase$(UserMemberOf) = "SPECIALIST REGISTRAR" Or _
         UCase$(UserMemberOf) = "IT MANAGER" Then

170       txtQCode.Enabled = True
180       txtQDescription.Enabled = True
190   End If
200   cmdEditDemo.Enabled = True

210   cmdSave.Enabled = True

End Sub

Private Function FillTree() As Boolean


      Dim tb As New Recordset
      Dim sql As String
      Dim nod As MSComctlLib.Node

10    On Error GoTo FillTree_Error

20    If txtCaseId = "" Then
30        sql = "Select * From CaseTree CT " & _
                "LEFT Join Cases C On CT.CaseID = C.CaseID " & _
                "LEFT JOIN CaseMovements CM ON CT.CaseId = CM.CaseId " & _
                "LEFT JOIN Demographics D On D.CaseId = CT.CaseId "



40        If UCase(cmbState) = "Authorised Not Printed" Then
50            sql = sql & "WHERE C.Validated = 1 AND (C.ValReportDate = '' OR C.ValReportDate IS NULL) "
         
60        ElseIf UCase(cmbState) = "EXTERNAL EVENTS - OUTSTANDING" Then
70            sql = sql & "WHERE (CM.CaseId <> '') "
80        ElseIf cmbState <> "" Then
90            sql = sql & "WHERE C.State = N'" & cmbState & "' "
100       Else
110           sql = sql & "WHERE 1=1 "
120       End If

130       If txtPatientId <> "" Then
140           If cmbPatientId = "NOPAS" Then
150               sql = sql & "AND D.Nopas = N'" & txtPatientId & "' "
160           ElseIf cmbPatientId = "MRN" Then
170               sql = sql & "AND D.Mrn = N'" & txtPatientId & "' "
180           ElseIf cmbPatientId = "A&E No" Then
190               sql = sql & "AND D.AandENo = N'" & txtPatientId & "' "
200           End If
210       End If

220       sql = sql & "Order By TreeOrder"
230   Else
240       sql = "Select * From CaseTree Where CaseId = N'" & CaseNo & "' Order By TreeOrder"
250   End If

260   RecOpenClient 0, tb, sql

270   If Not tb.EOF Then
280       With tvCaseDetails
290           .Nodes.Clear
300           While Not tb.EOF
310               If tb!LocationParentID = 0 Then
320                   Set nod = .Nodes.Add(, , "L" & tb!LocationLevel & tb!LocationID, tb!LocationName, 1, 2)
330                   nod.Bold = True

340               Else
350                   .SingleSel = True
                     
360                   Set nod = .Nodes.Add("L" & Val(tb!LocationLevel) - 1 & tb!LocationParentID, tvwChild, "L" & tb!LocationLevel & tb!LocationID, tb!LocationName, 1, 2)
370               End If

380               If tb!LocationLevel = 1 Then
390                   nod.Tag = tb!TissueTypeListId
400               End If

410               If tb!ExtraRequests & "" <> "" Then
420                   nod.Tag = tb!ExtraRequests
430                   If tb!ExtraRequests <> "0" Then
440                       nod.ForeColor = vbBlue
450                   End If
460               End If

470               If tb!NoOfSections & "" <> "" Then
480                   nod.Tag = tb!NoOfSections
490               End If

500               tb.MoveNext
510           Wend
520           .SingleSel = False
530       End With
540       FillTree = True
550   Else
560       FillTree = False
570   End If


580   Exit Function

FillTree_Error:

      Dim strES As String
      Dim intEL As Integer

590   intEL = Erl
600   strES = Err.Description
610   LogError "frmWorkSheet", "FillTree", intEL, strES, sql


End Function


Private Sub Form_Unload(Cancel As Integer)
      Dim i As Integer

10    If TimedOut Then
20        sCaseLockedBy = CaseLockedBy(CaseNo)
30        If sCaseLockedBy = UserName Then
40            UnlockCase
50        End If

60        Exit Sub
70    End If
80    If DataChanged = False Then
90        With frmWorklist
100           .Enabled = True
110           .tmrRefresh.Enabled = True

120           .sCaseId = ""
130       End With
140       DataMode = 0
150   Else

160       If frmMsgBox.Msg("Alert!! Do you want to save your changes?", mbYesNo, , mbQuestion) = 1 Then
170           cmdSave_Click
180       End If
190       With frmWorklist
200           .Enabled = True
210           .tmrRefresh.Enabled = True
220           .sCaseId = ""

230       End With
240       DataMode = 0

250   End If
260   Unload frmList
270   mclsToolTip.RemoveTool grdAmendments
280   For i = 0 To 3
290       mclsToolTip.RemoveTool grdTracker(i)
300   Next

310   sCaseLockedBy = CaseLockedBy(CaseNo)
320   If sCaseLockedBy = UserName Then
330       UnlockCase
340   End If

350   mclsToolTip.RemoveTool cmdClinicalHist
360   mclsToolTip.RemoveTool cmdComments
370   mclsToolTip.RemoveTool fraDemographics
380   Set mclsToolTip = Nothing

End Sub

Private Sub grdAmendments_DblClick()
10    If grdAmendments.Rows > 1 Then
20        With frmAmendments
30            .Update = True
40            .AmendId = grdAmendments.TextMatrix(Rada, 2)
50            .Code = grdAmendments.TextMatrix(Rada, 3)
60            .Move frmWorkSheet.Left + fraWorkSheet.Left + grdAmendments.Left - .Width, frmWorkSheet.Top + grdAmendments.Top
70            .Show vbModal

80        End With
90    End If
End Sub


Private Sub grdAmendments_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    If grdAmendments.Rows > 1 Then
20        Rada = grdAmendments.MouseRow
30        If Rada <> 0 Then
40            gridId = grdAmendments.TextMatrix(Rada, 2)
50            HighlightRow (gridId)

60        End If
70    End If
End Sub

Private Sub grdAmendments_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    GridToolTip grdAmendments, X, Y

End Sub

Private Sub GridToolTip(Grid As MSFlexGrid, X As Single, Y As Single)

      Dim lngRow As Long
      Dim lngCol As Long
      Dim lngStartPos As Long

10    On Error GoTo GridToolTip_Error

20    lngCol = 0
30    lngRow = 0

      '
      ' Find the column
      '

40    lngCol = Grid.MouseCol

      '
      ' Find the row
      '

50    lngRow = Grid.MouseRow

60    If lngRow = Grid.Rows Or lngCol = Grid.Cols Then

          ' Off the grid just blank the tooltip
70        Grid.ToolTipText = vbNullString

80    Else

          '
          ' Set the tool tip here. I'm just showing Row & Col as example

90        mclsToolTip.ToolText(Grid) = Replace(Grid.TextMatrix(lngRow, lngCol), "<<tab>>", " ")

100   End If


110   Exit Sub

GridToolTip_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmWorkSheet", "GridToolTip", intEL, strES

End Sub

Private Sub grdMCodes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    If grdMCodes.Rows > 1 Then
20        txtFindCode.Visible = False
30        txtFindDescription.Visible = False
40        Rada = grdMCodes.MouseRow
50        If Rada <> 0 Then
60            With grdMCodes
70                gridId = .TextMatrix(Rada, 2)
80                HighlightRow (gridId)

90            End With
100           If Button = 2 Then    ' Check if right mouse button was clicked.
110               PopupMenu mnuMCodesMenu   ' Display the menu as a pop-up menu.
120           End If
130       End If
140   End If
End Sub




Private Sub grdQCodes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    txtFindCode.Visible = False
20    txtFindDescription.Visible = False
30    If grdQCodes.Rows > 1 Then
40        Rada = grdQCodes.MouseRow
50        If Rada <> 0 Then
60            gridId = grdQCodes.TextMatrix(Rada, 2)
70            HighlightRow (gridId)
80            If Not CaseWasOriginalValidated Then    'ITS #819017 to fix
90                If Button = 2 Then   ' Check if right mouse button was clicked.
100                   PopupMenu mnuQCodesMenu   ' Display the menu as a pop-up menu.
110               End If
120           Else
130               If AllowedRemoveQcode(txtCaseId, gridId) Then
140                   PopupMenu mnuQCodesMenu   ' Display the menu as a pop-up menu.
150               End If
160           End If
170       End If
180   End If
End Sub

Private Function AllowedRemoveQcode(ByVal strCaseIdentification As String, ByVal strCaseListId As String) As Boolean
      Dim sql As String
      Dim tb As Recordset
      Dim dateQcodeAdded As Date
      Dim dateLatestAuthorisation As Date
      Dim strCaseId As String

      'here need to test this new function

10    On Error GoTo AllowedRemoveQcode_Error

20    strCaseId = Replace(strCaseIdentification, " ", "")
30    strCaseId = Replace(strCaseId, "/", "")

40    sql = "SELECT CL.CaseListId,CL.Username,CL.DateTimeCreated,C.ValReportDate from CaseListLink as CL, Cases as C " & _
            "WHERE c.CaseId = cl.CaseId " & _
            "AND c.CaseId = '" & strCaseId & "' " & _
            "AND CL.CaseListId = N'" & strCaseListId & "'"

50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql

70    If Not tb.EOF Then
80        dateQcodeAdded = Format(tb!DateTimeCreated, "dd/mmm/yyyy hh:mm:ss")
90        dateLatestAuthorisation = Format(tb!ValReportDate, "dd/mmm/yyyy hh:mm:ss")
100       If dateQcodeAdded > dateLatestAuthorisation Then
110           If UCase(UserName) = UCase(Trim$(tb!UserName & "")) Or UserMemberOf = "Manager" Then
120               AllowedRemoveQcode = True
130           Else
140               AllowedRemoveQcode = False
150           End If
160       Else
170           AllowedRemoveQcode = False
180       End If
190   Else
200       AllowedRemoveQcode = False
210   End If

220   Exit Function

AllowedRemoveQcode_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmWorkSheet", "AllowedRemoveQcode", intEL, strES, sql

End Function

Private Sub HighlightRow(Id As String)
      Dim j As Integer
      Dim r As Integer
      Dim i As Integer
10    r = 1
20    With grdAmendments
30        .BackColor = vbWhite
40        .ForeColor = vbBlack
50        Do Until r = .Rows
60            If .TextMatrix(r, 2) = gridId Then
70                .row = r
80                .RowSel = r
90                .TopRow = r
100               For j = 0 To .Cols - 1
110                   .col = j
120                   .CellBackColor = &H80000015
130                   .CellForeColor = &H80000005
140               Next
150           Else
160               .row = r
170               .RowSel = r
180               For j = 0 To .Cols - 1
190                   .col = j
200                   .CellBackColor = vbWhite
210                   If .TextMatrix(r, 3) = "Q021" Then
220                       .CellForeColor = vbRed
230                   Else
240                       .CellForeColor = vbBlack
250                   End If
260               Next

270           End If
280           r = r + 1
290       Loop
300   End With

310   r = 1
320   With grdQCodes
330       .BackColor = vbWhite
340       .ForeColor = vbBlack
350       Do Until r = .Rows
360           If .TextMatrix(r, 2) = gridId Then
370               .row = r
380               .RowSel = r
390               .TopRow = r
400               For j = 0 To .Cols - 1
410                   .col = j
420                   .CellBackColor = &H80000015
430                   .CellForeColor = &H80000005
440               Next
450           Else
460               .row = r
470               .RowSel = r
480               For j = 0 To .Cols - 1
490                   .col = j
500                   .CellBackColor = vbWhite
510                   If .TextMatrix(r, 0) = "Q021" Then
520                       .CellForeColor = vbRed
530                   Else
540                       .CellForeColor = vbBlack
550                   End If
560               Next
570           End If
580           r = r + 1
590       Loop
600   End With

610   For i = 0 To 3
620       r = 1
630       With grdTracker(i)
640           .BackColor = vbWhite
650           .ForeColor = vbBlack
660           Do Until r = .Rows
670               If .TextMatrix(r, 5) = gridId Then
680                   .row = r
690                   .RowSel = r
700                   .TopRow = r
710                   For j = 0 To .Cols - 1
720                       .col = j
730                       .CellBackColor = &H80000015
740                       .CellForeColor = &H80000005

750                   Next
760               Else
770                   .row = r
780                   .RowSel = r
790                   For j = 0 To .Cols - 1
800                       .col = j
810                       .CellBackColor = vbWhite
820                       .CellForeColor = vbBlack

830                   Next
840               End If
850               r = r + 1
860           Loop
870       End With
880   Next i

890   r = 1
900   With grdMCodes
910       .BackColor = vbWhite
920       .ForeColor = vbBlack
930       Do Until r = .Rows
940           If .TextMatrix(r, 2) = gridId Then
950               .row = r
960               .RowSel = r

970               For j = 0 To .Cols - 1
980                   .col = j
990                   .CellBackColor = &H80000015
1000                  .CellForeColor = &H80000005
1010              Next
1020          Else
1030              .row = r
1040              .RowSel = r

1050              For j = 0 To .Cols - 1
1060                  .col = j
1070                  .CellBackColor = vbWhite
1080                  .CellForeColor = vbBlack
1090              Next
1100          End If
1110          r = r + 1
1120      Loop
1130  End With
End Sub

Private Sub mnuAddTissueType_Click()

10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .Update = False
50            .ListType = "T"
60            .ListTypeName = "Tissue Type"
70            .ListTypeNames = "Tissue Types"
80            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top - 750
90            .Show
100       End With
110       frmWorkSheet.Enabled = False
120   End With
End Sub



Private Sub mnuAmendDel_Click()
      Dim r As Integer
      Dim s As String

10    s = grdAmendments.TextMatrix(Rada, 2) & vbTab & grdAmendments.TextMatrix(Rada, 3) & vbTab & "CaseAmendments"
20    grdDelete.AddItem s

30    If grdAmendments.Rows - grdAmendments.FixedRows = 1 Then
40        grdAmendments.Rows = grdAmendments.Rows - 1
50    Else
60        grdAmendments.RemoveItem Rada
70    End If

80    r = 1
90    Do Until r = grdQCodes.Rows
100       If grdQCodes.TextMatrix(r, 2) = gridId Then
110           s = grdQCodes.TextMatrix(r, 2) & vbTab & grdQCodes.TextMatrix(r, 0) & vbTab & "CaseListLink"
120           grdDelete.AddItem s
130           If grdQCodes.Rows - grdQCodes.FixedRows = 1 Then
140               grdQCodes.Rows = grdQCodes.Rows - 1
150           Else
160               grdQCodes.RemoveItem r
170           End If
180           Exit Do
190       End If
200       r = r + 1
210   Loop

220   DataChanged = True
End Sub




Private Sub mnuDelStain_Click()
      Dim r As Integer
      Dim s As String

10    With tvCaseDetails

20        r = 1
30        Do Until r = grdTracker(1).Rows
40            If grdTracker(1).TextMatrix(r, 5) = .SelectedItem.Key Then
50                s = grdTracker(1).TextMatrix(r, 5) & vbTab & grdTracker(1).TextMatrix(r, 4) & vbTab & "CaseMovements"
60                grdDelete.AddItem s
70                If grdTracker(1).Rows - grdTracker(1).FixedRows = 1 Then
80                    grdTracker(1).Rows = grdTracker(1).Rows - 1
90                Else
100                   grdTracker(1).RemoveItem r
110               End If
120               Exit Do
130           End If
140           r = r + 1
150       Loop


160       DeleteNode

170   End With
180   DataChanged = True
190   TreeChanged = True

End Sub



Private Sub mnuImmunoStainLevel4_Click()

10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .ListType = "IS"
50            .Level = "L3"
60            .ListTypeName = "Immunohistochemical Stain"
70            .ListTypeNames = "Immunohistochemical Stains"
80            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top
90            .Show
100       End With
110   End With
120   frmWorkSheet.Enabled = False
End Sub





Private Sub mnuMCodesDel_Click()
      Dim s As String
      Dim i As Integer

10    s = grdMCodes.TextMatrix(Rada, 2) & vbTab & grdMCodes.TextMatrix(Rada, 0) & vbTab & "CaseListLink" & _
          vbTab & grdMCodes.TextMatrix(Rada, 5)
20    grdDelete.AddItem s
30    For i = 0 To grdTempMCode.Rows - 1
40        If grdTempMCode.TextMatrix(i, 2) = grdMCodes.TextMatrix(Rada, 2) _
             And grdTempMCode.TextMatrix(i, 3) = grdMCodes.TextMatrix(Rada, 3) _
             And grdTempMCode.TextMatrix(i, 4) = grdMCodes.TextMatrix(Rada, 4) Then
50            If grdTempMCode.Rows - grdTempMCode.FixedRows = 1 Then
60                grdTempMCode.Rows = grdTempMCode.Rows - 1
70                Exit For
80            Else
90                grdTempMCode.RemoveItem i
100               Exit For
110           End If
120       End If
130   Next i
140   If grdMCodes.Rows - grdMCodes.FixedRows = 1 Then
150       grdMCodes.Rows = grdMCodes.Rows - 1
160   Else
170       grdMCodes.RemoveItem Rada
180   End If


190   DataChanged = True
End Sub


'********* Delete row from movement tracker *************
Private Sub mnuMoveSpecDel_Click()

      Dim s As String

10    s = grdTracker(0).TextMatrix(Rada, 5) & vbTab & grdTracker(0).TextMatrix(Rada, 4) & vbTab & "CaseMovements"
20    grdDelete.AddItem s
30    If grdTracker(0).Rows - grdTracker(0).FixedRows = 1 Then
40        grdTracker(0).Rows = grdTracker(0).Rows - 1
50    Else
60        grdTracker(0).RemoveItem Rada
70    End If



End Sub

Private Sub mnuMultipleBlocks_Click()

10    With frmInputNo

20        .InputType = "B"
30        .Label = "Please Enter Number of Blocks"
40        .Move ScaleX(TreePositionX, vbPixels, vbTwips), ScaleY(TreePositionY, vbPixels, vbTwips)
50        .Show 1
60    End With

End Sub

Private Sub mnuMultipleSlidesLevel2_Click()
10    AddMultipleSlides

End Sub

Private Sub AddMultipleSlides()
10    With frmInputNo

20        .InputType = "S"
30        .Label = "Please Enter Number of Slides"
40        .Move ScaleX(TreePositionX, vbPixels, vbTwips), ScaleY(TreePositionY, vbPixels, vbTwips)
50        .Show 1
60    End With
End Sub


Private Sub mnuQCodesDel_Click()
      Dim r As Integer
      Dim s As String

10    s = grdQCodes.TextMatrix(Rada, 2) & vbTab & grdQCodes.TextMatrix(Rada, 0) & vbTab & "CaseListLink"
20    grdDelete.AddItem s
30    If grdQCodes.Rows - grdQCodes.FixedRows = 1 Then
40        grdQCodes.Rows = grdQCodes.Rows - 1
50    Else
60        grdQCodes.RemoveItem Rada
70    End If


80    r = 1
90    Do Until r = grdTracker(2).Rows
100       If grdTracker(2).TextMatrix(r, 5) = gridId Then
110           s = grdTracker(2).TextMatrix(r, 5) & vbTab & grdTracker(2).TextMatrix(r, 4) & vbTab & "CaseMovements"
120           grdDelete.AddItem s
130           If grdTracker(2).Rows - grdTracker(2).FixedRows = 1 Then
140               grdTracker(2).Rows = grdTracker(2).Rows - 1
150           Else
160               grdTracker(2).RemoveItem r
170           End If
180           Exit Do
190       End If
200       r = r + 1
210   Loop

220   r = 1
230   Do Until r = grdAmendments.Rows
240       If grdAmendments.TextMatrix(r, 2) = gridId Then
250           s = grdAmendments.TextMatrix(r, 2) & vbTab & grdAmendments.TextMatrix(r, 3) & vbTab & "CaseAmendments"
260           grdDelete.AddItem s
270           If grdAmendments.Rows - grdAmendments.FixedRows = 1 Then
280               grdAmendments.Rows = grdAmendments.Rows - 1
290           Else
300               grdAmendments.RemoveItem r
310           End If
320           Exit Do
330       End If
340       r = r + 1
350   Loop
360   DataChanged = True
End Sub



Private Sub mnuSpecialStainLevel4_Click()

10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .ListType = "SS"
50            .Level = "L3"
60            .ListTypeName = "Special Stain"
70            .ListTypeNames = "Special Stains"
80            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top
90            .Show
100       End With
110   End With
120   frmWorkSheet.Enabled = False
End Sub





Private Sub mnuSingleBlock_Click()
      Dim iBlock As Integer
      Dim tnode As MSComctlLib.Node
      Dim NoOfBlocks As Integer
      Dim UniqueId As String


10    On Error GoTo mnuSingleBlock_Click_Error

20    If UCase(Left(Trim(tvCaseDetails.SelectedItem.Text), 14)) = "Frozen Section" Then
30        NoOfBlocks = GetBlockNumber(tvCaseDetails.SelectedItem.Parent)
40    Else
50        NoOfBlocks = GetBlockNumber(tvCaseDetails.SelectedItem)
60    End If
70    With tvCaseDetails.SelectedItem
80        UniqueId = GetUniqueID

90        If sysOptBlockNumberingFormat(0) = "1" Then
100           iBlock = NoOfBlocks + 1
110           Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L2" & UniqueId, "Block" & " " & iBlock, 1, 2)
120           If InStr(1, tvCaseDetails.SelectedItem.Text, "Frozen Section") Then
130               AddDefaultStains "L2" & UniqueId, tvCaseDetails.SelectedItem
140           End If
150       Else
160           If NoOfBlocks = 1 Then
170               .Child.Text = "Block" & " A"
180           End If
190           iBlock = NoOfBlocks + 65
200           Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L2" & UniqueId, "Block" & " " & IIf(Chr(iBlock) = "A", "", Chr(iBlock)), 1, 2)
210           If UCase(Left(Trim(tvCaseDetails.SelectedItem.Text), 14)) <> "Frozen Section" Then
220               AddDefaultStains "L2" & UniqueId, tvCaseDetails.SelectedItem
230           End If
240       End If
250       tnode.Expanded = True
260       tnode.Selected = True
270   End With
280   DataChanged = True
290   TreeChanged = True

300   Exit Sub

mnuSingleBlock_Click_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmWorkSheet", "mnuSingleBlock_Click", intEL, strES

End Sub



Private Sub mnuSingleSlideLevel2_Click()
10    AddSingleSlide
End Sub
Private Sub AddSingleSlide()
      Dim iSlide As Integer
      Dim tnode As MSComctlLib.Node
      
      Dim NoOfSlides As Integer
      Dim UniqueId As String

10    On Error GoTo AddSingleSlide_Error

20    NoOfSlides = GetSlideNumber(tvCaseDetails.SelectedItem)
30    With tvCaseDetails.SelectedItem
40        UniqueId = GetUniqueID

50        If sysOptSlideNumberingFormat(0) = "1" Then
60            iSlide = NoOfSlides + 1
70            Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L3" & UniqueId, "Slide" & " " & iSlide, 1, 2)
80        Else
90            If NoOfSlides = 1 Then
100               .Child.Text = "Slide" & " A"
110           End If
120           iSlide = NoOfSlides + 65
130           Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L3" & UniqueId, "Slide" & " " & IIf(Chr(iSlide) = "A", "", Chr(iSlide)), 1, 2)
140       End If
150       .Expanded = True
160       tnode.Selected = True
170   End With
180   DataChanged = True
190   TreeChanged = True

200   Exit Sub

AddSingleSlide_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmWorkSheet", "AddSingleSlide", intEL, strES

End Sub

Private Sub optReport_Click(Index As Integer)
      Dim sql As String
      Dim tb As Recordset
      Dim sn As Recordset

10    If Index = 1 Then



20        sql = "SELECT * FROM CaseTree CT " & _
                "WHERE CT.CaseId = N'" & CaseNo & "' AND CT.LocationLevel = '1'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        Do While Not tb.EOF
60            If tb!LocationID & "" <> "" Then
70                sql = "SELECT * FROM CaseListLink WHERE CaseId = N'" & CaseNo & "' AND TissueTypeId = N'" & tb!LocationID & "' "
80                Set sn = New Recordset
90                RecOpenServer 0, sn, sql
100               If Not sn.EOF Then
110                   tb.MoveNext
120               Else

130                   If fraLinkedCase.Visible = False Or _
                         Left(txtCaseId, 1) = "C" Or _
                         Mid(txtCaseId, 2, 1) = "A" Then
140                       MsgBox ("Please enter M Code before Authorisation")
150                       optReport(1).Value = False
160                       Exit Sub
170                   Else
180                       Exit Do
190                   End If
200               End If
210           Else
220               tb.MoveNext
230           End If

240       Loop

250       If (txtPDescription = "" Or txtPCode = "") And (fraLinkedCase.Visible = False Or _
                                                          Left(txtCaseId, 1) = "C" Or _
                                                          Mid(txtCaseId, 2, 1) = "A") Then
260           MsgBox ("Please enter P Code before Authorisation")
270           optReport(1).Value = False
280           Exit Sub
290       End If

300       optState(0).Value = False
310       optState(1).Value = False
320       optState(2).Value = False
330       lblWithPathologist = ""
340       lblWithPathologistName = ""
350       lblCheckedBy = ""
360   End If
370   DataChanged = True
End Sub

Private Sub tvCaseDetails_Click()
10    On Error GoTo tvCaseDetails_Click_Error
20    If Not Activated Then Exit Sub

30    If tvCaseDetails.SelectedItem Is Nothing Then
40        If Not PrevNode Is Nothing Then
50            PrevNode.Selected = True
60        End If
70    End If

80    Exit Sub

tvCaseDetails_Click_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmWorkSheet", "tvCaseDetails_Click", intEL, strES


End Sub

Private Sub tvCaseDetails_NodeClick(ByVal Node As MSComctlLib.Node)

      Dim n As MSComctlLib.Node
10    On Error GoTo tvCaseDetails_NodeClick_Error

20    If Not Activated Then Exit Sub

30    Set n = Node
      Dim PrevParentNode As MSComctlLib.Node

      Dim SearchNode As MSComctlLib.Node

40    While Not n.Parent Is Nothing
50        Set n = n.Parent
60    Wend

70    If DataMode = DataModeEdit Then
80        If Not PrevNode Is Nothing Then
90            Set PrevParentNode = PrevNode
100           While Not PrevParentNode.Parent Is Nothing
110               Set PrevParentNode = PrevParentNode.Parent
120           Wend
130           If PrevParentNode <> n Then
140               If DataChanged = False Then
150                   Set PrevNode = Node
160                   LoadDetails n
170               Else

180                   If MsgBox("Save Changes", vbYesNo + vbQuestion + vbDefaultButton2) = 1 Then
190                       cmdSave_Click
200                   End If

210                   Set PrevNode = Node

220                   txtCaseId = ""
230                   cmbPatientId = PatientIdCombo
240                   txtPatientId = PatientIdText
250                   FillTree



260                   With tvCaseDetails
270                       For Each SearchNode In .Nodes
280                           If SearchNode.Key = Node.Key Then
290                               SearchNode.Selected = True
300                               Exit For
310                           End If
320                       Next
330                   End With
340                   LoadDetails n

350               End If

360           End If
370       Else
380           If Search Then
390               Select Case Left(tvCaseDetails.SelectedItem.Key, 2)
                  Case "L0"
400                   Set PrevNode = Node
410                   LoadDetails n
420               End Select
430           End If
440       End If



450   End If

460   Exit Sub

tvCaseDetails_NodeClick_Error:

      Dim strES As String
      Dim intEL As Integer

470   intEL = Erl
480   strES = Err.Description
490   LogError "frmWorkSheet", "tvCaseDetails_NodeClick", intEL, strES

End Sub
Private Sub LoadDetails(ByVal Node As MSComctlLib.Node)

10    ResetWorkSheet
20    txtPatientId = ""
30    txtCaseId = Node.Text
40    CaseNo = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
50    LoadDemographics CaseNo
60    FillWorkSheet CaseNo
70    If Not Search Then
80        fraWorkSheet.Enabled = True
90    End If
End Sub
'Zyam 26-07-24
Private Sub txtCaseId_KeyUp(KeyCode As Integer, Shift As Integer)
    UCase (txtCaseId.Text)
End Sub

'Zyam 26-07-24
Private Sub txtCaseId_LostFocus()
      Dim tb As New Recordset
      Dim sql As String
      Dim UniqueId As String
      Dim strCaseId As String

10    On Error GoTo txtCaseId_LostFocus_Error

      
20    If Trim$(txtCaseId) <> "" Then
30        strCaseId = UCase(Replace(Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", ""), " ", ""))
40        If VerifyCaseIdFormat(strCaseId) Then
50            If CaseIdDemoEntered(strCaseId) Then
60                DataMode = DataModeNew
70                If DataChanged = False Then
80                    sCaseLockedBy = CaseLockedBy(CaseNo)
90                    If IsValidCaseNo(txtCaseId) Then
100                       fraWorkSheet.Enabled = True
110                       Validated = False
120                       CaseNo = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
130                       sCaseLockedBy = CaseLockedBy(CaseNo)
140                       If sCaseLockedBy <> UserName And sCaseLockedBy <> "" Then
150                           lblCaseLocked = "RECORD BEING EDITED BY" & " " & sCaseLockedBy
160                           bLocked = True
170                           lblCaseLocked.BackColor = &H8080FF
180                       ElseIf sCaseLockedBy = "" Then
190                           LockCase CaseNo
200                           bLocked = False
210                           UnLockControl
220                           lblCaseLocked.BackColor = &H80FF80
230                           lblCaseLocked = "RECORD BEING EDITED BY YOU!"
240                       Else
250                           bLocked = False
260                           UnLockControl
270                           lblCaseLocked.BackColor = &H80FF80
280                           lblCaseLocked = "RECORD BEING EDITED BY YOU!"
290                       End If

300                       sql = "SELECT * FROM Cases WHERE CaseId = N'" & CaseNo & "'"
310                       Set tb = New Recordset
320                       RecOpenClient 0, tb, sql
330                       If Not tb.EOF Then
340                           txtPatientId = ""
350                           ResetWorkSheet
360                           tvCaseDetails.Nodes.Clear
370                           If FillTree Then
380                               DataMode = DataModeEdit
390                               tvCaseDetails.Nodes(1).Selected = True
400                               LoadDemographics CaseNo
410                               FillWorkSheet CaseNo
420                           End If
430                           ExpandAll tvCaseDetails
440                           If lngMaxDigits = 12 Then
450                               SetAutopsyWorksheet
460                           Else
470                               SetHisCytWorksheet
480                           End If
490                       Else
500                           If txtCaseId = mCaseId Then
510                               Exit Sub
520                           End If
530                           txtPatientId = ""
540                           EnableCase
550                           ResetWorkSheet
560                           UniqueId = GetUniqueID
570                           If txtCaseId = "" Then
580                               tvCaseDetails.Nodes.Clear
590                           Else
600                               tvCaseDetails.Nodes.Clear
610                               If lngMaxDigits = 12 Then
620                                   tvCaseDetails.Nodes.Add , , "L0" & UniqueId, Left(CaseNo, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(CaseNo, 2), 1, 2
630                                   SetAutopsyWorksheet
640                               Else
650                                   tvCaseDetails.Nodes.Add , , "L0" & UniqueId, Left(CaseNo, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(CaseNo, 2), 1, 2
660                                   SetHisCytWorksheet
670                               End If
680                               tvCaseDetails.Nodes(1).Selected = True
690                               TreeChanged = True

700                           End If

710                           DataMode = DataModeNew

720                       End If

730                       If bLocked Then
740                           LockControl
750                       End If
760                   Else
770                       txtCaseId = ""
780                       ResetSearch
790                       ResetWorkSheet
800                   End If
810               End If
820           Else
830               fraWorkSheet.Enabled = False
840               txtPatientId = ""
850               Set PrevNode = Nothing
860               tvCaseDetails.Nodes.Clear
870               cmdSearch.Enabled = True
880               TreeChanged = False
890               ResetWorkSheet
900           End If
910       Else
920           txtCaseId = ""
930       End If
940   End If

'950   cmdMCode.Enabled = False
'960   txtMCode.Enabled = False
'970   txtMDescription.Enabled = False
980   Set PrevNode = Nothing
990   Exit Sub

txtCaseId_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

1000  intEL = Erl
1010  strES = Err.Description
1020  LogError "frmWorkSheet", "txtCaseId_LostFocus", intEL, strES, sql


End Sub

Private Sub LockControl()
      Dim i As Integer

10    txtPCode.Enabled = False
20    txtPDescription.Enabled = False
30    txtMCode.Enabled = False
40    txtMDescription.Enabled = False
50    txtGross.Locked = True
60    txtMicro.Locked = True
70    txtContainerLabel.Enabled = False
80    txtNOS.Enabled = False
90    txtQCode.Enabled = False
100   txtQDescription.Enabled = False
110   cmdQCode.Enabled = False
120   cmdMCode.Enabled = False

130   For i = 0 To 3
140       grdTracker(i).Enabled = False
150   Next

160   fraCaseState.Enabled = False
170   fraReport.Enabled = False
180   cmdSave.Enabled = False
190   cmdCytoHist.Enabled = False
200   txtPatientId.Enabled = False
210   cmbState.Enabled = False

End Sub

Private Sub UnLockControl()

      Dim strCaseState As String

10    On Error GoTo UnLockControl_Error

20    strCaseState = AuthorisedOrAddendumAdded(Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", ""))

30    If strCaseState = "1" Then    'Not Authorised
40        caseStatusNotAuthorised
50    ElseIf strCaseState = "2" Then    'Authorised and No addendum added
60        caseStatusAuthorisedNoAddendum
70    ElseIf strCaseState = "3" Then    'Authorised and Addendum added
80        caseStatusAuthorisedAddendumAdded
90    Else
100       caseStatusUnknown
110   End If

120   Exit Sub

UnLockControl_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmWorkSheet", "UnLockControl", intEL, strES

End Sub

Private Sub caseStatusUnknown()
      Dim i As Integer

10    txtGross.Locked = True             'Gross Description
20    txtMicro.Locked = True             'Micro Description
30    txtPCode.Enabled = False             'P Code & Description
40    txtPDescription.Enabled = False
50    txtQCode.Enabled = False             'Q code and description
60    txtQDescription.Enabled = False
70    cmdQCode.Enabled = False             'Q code add button
80    fraCaseState.Enabled = True         'Frame for Case states
      'In Histology, With Pathologist, Awaiting Authorisation
90    fraReport.Enabled = False            'Prelimary and Authorised Report

100   For i = 0 To 3                      'Movement Tracker
110       grdTracker(i).Enabled = False
120   Next
130   cmdCytoHist.Enabled = False          'Histo link button
140   txtPatientId.Enabled = False         'Patient Id
150   cmbState.Enabled = False             'Case state combo for search filter
160   txtContainerLabel.Enabled = False    'Container labels
170   txtNOS.Enabled = False               'Nature Of Specimen
180   cmdSave.Enabled = False              'Save case command button

End Sub

Private Sub caseStatusAuthorisedAddendumAdded()
      Dim i As Integer

10    If UCase$(UserMemberOf) = "SCIENTIST" Or UCase$(UserMemberOf) = "MANAGER" Then
20        txtGross.Locked = True             'Gross Description
30        txtMicro.Locked = True             'Micro Description
40        txtPCode.Enabled = False             'P Code & Description
50        txtPDescription.Enabled = False
60        txtQCode.Enabled = True             'Q code and description
70        txtQDescription.Enabled = True
80        cmdQCode.Enabled = True             'Q code add button
90        fraCaseState.Enabled = True         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
100       fraReport.Enabled = False           'Prelimary and Authorised Report
110   ElseIf UCase$(UserMemberOf) = "CONSULTANT" Then
120       txtGross.Locked = True             'Gross Description
130       txtMicro.Locked = True             'Micro Description
140       txtPCode.Enabled = False             'P Code & Description
150       txtPDescription.Enabled = False
160       txtQCode.Enabled = True             'Q code and description
170       txtQDescription.Enabled = True
180       cmdQCode.Enabled = True             'Q code add button
190       fraCaseState.Enabled = True         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
200       fraReport.Enabled = True            'Prelimary and Authorised Report
210   ElseIf UCase$(UserMemberOf) = "CLERICAL" Then    'CLERICAL
220       txtGross.Locked = True             'Gross Description
230       txtMicro.Locked = True             'Micro Description
240       txtPCode.Enabled = False             'P Code & Description
250       txtPDescription.Enabled = False
260       txtQCode.Enabled = False             'Q code and description
270       txtQDescription.Enabled = False
280       cmdQCode.Enabled = False             'Q code add button
290       fraCaseState.Enabled = False         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
300       fraReport.Enabled = False            'Prelimary and Authorised Report
310   End If

320   For i = 0 To 3                      'Movement Tracker
330       grdTracker(i).Enabled = True
340   Next
350   cmdCytoHist.Enabled = True          'Histo link button
360   txtPatientId.Enabled = True         'Patient Id
370   cmbState.Enabled = True             'Case state combo for search filter
380   txtContainerLabel.Enabled = True    'Container labels
390   txtNOS.Enabled = True               'Nature Of Specimen
400   cmdSave.Enabled = True              'Save case command button


End Sub

Private Sub caseStatusAuthorisedNoAddendum()
      Dim i As Integer
10    If UCase$(UserMemberOf) = "SCIENTIST" Or UCase$(UserMemberOf) = "MANAGER" Then
20        txtGross.Locked = True             'Gross Description
30        txtMicro.Locked = True             'Micro Description
40        txtPCode.Enabled = False             'P Code & Description
50        txtPDescription.Enabled = False
60        txtQCode.Enabled = True             'Q code and description
70        txtQDescription.Enabled = True
80        cmdQCode.Enabled = True             'Q code add button
90        fraCaseState.Enabled = False         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
100       fraReport.Enabled = False           'Prelimary and Authorised Report
110   ElseIf UCase$(UserMemberOf) = "CONSULTANT" Then
120       txtGross.Locked = True             'Gross Description
130       txtMicro.Locked = True             'Micro Description
140       txtPCode.Enabled = False             'P Code & Description
150       txtPDescription.Enabled = False
160       txtQCode.Enabled = True             'Q code and description
170       txtQDescription.Enabled = True
180       cmdQCode.Enabled = True             'Q code add button
190       fraCaseState.Enabled = True         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
200       fraReport.Enabled = True            'Prelimary and Authorised Report
210   ElseIf UCase$(UserMemberOf) = "CLERICAL" Then    'CLERICAL
220       txtGross.Locked = True             'Gross Description
230       txtMicro.Locked = True             'Micro Description
240       txtPCode.Enabled = False             'P Code & Description
250       txtPDescription.Enabled = False
260       txtQCode.Enabled = False             'Q code and description
270       txtQDescription.Enabled = False
280       cmdQCode.Enabled = False             'Q code add button
290       fraCaseState.Enabled = False         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
300       fraReport.Enabled = False            'Prelimary and Authorised Report
310   End If

320   For i = 0 To 3                      'Movement Tracker
330       grdTracker(i).Enabled = True
340   Next
350   cmdCytoHist.Enabled = True          'Histo link button
360   txtPatientId.Enabled = True         'Patient Id
370   cmbState.Enabled = True             'Case state combo for search filter
380   txtContainerLabel.Enabled = True    'Container labels
390   txtNOS.Enabled = True               'Nature Of Specimen
400   cmdSave.Enabled = True              'Save case command button

End Sub

Private Sub caseStatusNotAuthorised()
      Dim i As Integer
10    If UCase$(UserMemberOf) = "SCIENTIST" Or UCase$(UserMemberOf) = "MANAGER" Then
20        txtGross.Locked = False             'Gross Description
30        txtMicro.Locked = False             'Micro Description
40        txtPCode.Enabled = True             'P Code & Description
50        txtPDescription.Enabled = True
60        txtQCode.Enabled = True             'Q code and description
70        txtQDescription.Enabled = True
80        cmdQCode.Enabled = True             'Q code add button
90        fraCaseState.Enabled = True         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
100       fraReport.Enabled = False           'Prelimary and Authorised Report
110   ElseIf UCase$(UserMemberOf) = "CONSULTANT" Then
120       txtGross.Locked = False             'Gross Description
130       txtMicro.Locked = False             'Micro Description
140       txtPCode.Enabled = True             'P Code & Description
150       txtPDescription.Enabled = True
160       txtQCode.Enabled = True             'Q code and description
170       txtQDescription.Enabled = True
180       cmdQCode.Enabled = True             'Q code add button
190       fraCaseState.Enabled = True         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
200       fraReport.Enabled = True            'Prelimary and Authorised Report
210   ElseIf UCase$(UserMemberOf) = "CLERICAL" Then    'CLERICAL
220       txtGross.Locked = False             'Gross Description
230       txtMicro.Locked = False             'Micro Description
240       txtPCode.Enabled = False             'P Code & Description
250       txtPDescription.Enabled = False
260       txtQCode.Enabled = False             'Q code and description
270       txtQDescription.Enabled = False
280       cmdQCode.Enabled = False             'Q code add button
290       fraCaseState.Enabled = False         'Frame for Case states
          'In Histology, With Pathologist, Awaiting Authorisation
300       fraReport.Enabled = False            'Prelimary and Authorised Report
310   End If

320   For i = 0 To 3                      'Movement Tracker
330       grdTracker(i).Enabled = True
340   Next
350   cmdCytoHist.Enabled = True          'Histo link button
360   txtPatientId.Enabled = True         'Patient Id
370   cmbState.Enabled = True             'Case state combo for search filter
380   txtContainerLabel.Enabled = True    'Container labels
390   txtNOS.Enabled = True               'Nature Of Specimen
400   cmdSave.Enabled = True              'Save case command button
End Sub

Private Sub SetAutopsyWorksheet()
10    txtMicro.Visible = False
20    txtGross.Height = 4000
End Sub

Private Sub SetHisCytWorksheet()
10    txtMicro.Visible = True
20    txtGross.Height = 1815
End Sub



Private Sub tvCaseDetails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Dim n As MSComctlLib.Node
      Dim PrevParentNode As MSComctlLib.Node
      Dim SearchNode As MSComctlLib.Node
      Dim i As Integer


10    On Error GoTo tvCaseDetails_MouseDown_Error
20    If Not Activated Then Exit Sub

30    Set tvCaseDetails.SelectedItem = tvCaseDetails.HitTest(X, Y)
40    If tvCaseDetails.SelectedItem Is Nothing Then

50        Exit Sub
60    End If
      'if record not locked by some other user
70    If Not bLocked Then
80        If Button = vbRightButton Then
90            Set n = tvCaseDetails.SelectedItem

              'finds the top node
100           While Not n.Parent Is Nothing
110               Set n = n.Parent
120           Wend


130           If Not PrevNode Is Nothing Then
140               Set PrevParentNode = PrevNode
150               While Not PrevParentNode.Parent Is Nothing
160                   Set PrevParentNode = PrevParentNode.Parent
170               Wend
                  'this is when there is more than one caseid in the tree (ie you use the search facility under the tree
180               If PrevParentNode <> n Then
190                   If DataChanged = False Then
200                       Set PrevNode = tvCaseDetails.SelectedItem
210                       If UCase$(UserMemberOf) <> "CLERICAL" Then
220                           CallTreePopupMenu tvCaseDetails, X, Y
230                       End If

240                   Else
250                       If MsgBox("Save Changes?", vbYesNo + vbQuestion + vbDefaultButton2) = 1 Then
260                           cmdSave_Click
270                       End If
280                       Set PrevNode = tvCaseDetails.SelectedItem
290                       i = tvCaseDetails.SelectedItem.Index
300                       If tvCaseDetails.SelectedItem Is Nothing Then
310                           Exit Sub
320                       Else
330                           txtCaseId = ""
340                           cmbPatientId = PatientIdCombo
350                           txtPatientId = PatientIdText
360 FillTree

370                           With tvCaseDetails
380                               For Each SearchNode In .Nodes
390                                   If SearchNode.Key = tvCaseDetails.Nodes(i).Key Then
400                                       SearchNode.Selected = True
410                                       Exit For
420                                   End If
430                               Next
440                           End With
450                           LoadDetails n
460                       End If
470                   End If
480               Else
490                   Set PrevNode = tvCaseDetails.SelectedItem
500                   If UCase$(UserMemberOf) <> "CLERICAL" Then
510                       CallTreePopupMenu tvCaseDetails, X, Y
520                   End If
530               End If
540           Else
550               Set PrevNode = tvCaseDetails.SelectedItem
560               If UCase$(UserMemberOf) <> "CLERICAL" Then
570                   CallTreePopupMenu tvCaseDetails, X, Y
580               End If
590           End If
600       End If
610   End If

620   Exit Sub

tvCaseDetails_MouseDown_Error:

      Dim strES As String
      Dim intEL As Integer

630   intEL = Erl
640   strES = Err.Description
650   LogError "frmWorkSheet", "tvCaseDetails_MouseDown", intEL, strES

End Sub
Private Sub FillMCodes(TissueTypeId As String, TissueTypeLetter As String)
      Dim s As String
      Dim i As Integer

10    On Error GoTo FillMCodes_Error



20    For i = 1 To grdTempMCode.Rows - 1
30        If grdTempMCode.TextMatrix(i, 3) = TissueTypeId And grdTempMCode.TextMatrix(i, 4) = TissueTypeLetter Then
40            s = grdTempMCode.TextMatrix(i, 0) _
                  & vbTab & grdTempMCode.TextMatrix(i, 1) _
                  & vbTab & grdTempMCode.TextMatrix(i, 2) _
                  & vbTab & grdTempMCode.TextMatrix(i, 3) _
                  & vbTab & grdTempMCode.TextMatrix(i, 4) _
                  & vbTab & "" _
                  & vbTab & grdTempMCode.TextMatrix(i, 6)
50            grdMCodes.AddItem s
60        End If
70    Next i

80    Exit Sub

FillMCodes_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmWorkSheet", "FillMCodes", intEL, strES


End Sub
Private Function CaseInCutUp() As Boolean
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CheckIfCutUp_Error

20    sql = "SELECT Phase FROM Cases WHERE State = N'" & "C" & "' " & _
            "AND Phase = N'" & "Cut-Up" & "' " & _
            "AND CaseId = N'" & CaseNo & "' "
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If Not tb.EOF Then
60        CaseInCutUp = True
70    Else
80        CaseInCutUp = False
90    End If

100   Exit Function

CheckIfCutUp_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmWorkSheet", "CheckIfCutUp", intEL, strES, sql

End Function
Private Sub CallTreePopupMenu(tvtemp As MSComctlLib.TreeView, X As Single, Y As Single)
      Dim oNode As MSComctlLib.Node
      Dim oChildNode As MSComctlLib.Node
      Dim coord As POINTAPI    ' receives coordinates of cursor

10    On Error GoTo CallTreePopupMenu_Error

20    If tvtemp.SelectedItem Is Nothing Then
30        Exit Sub
40    Else
50        Call GetCursorPos(coord)
60        Set tvtemp.SelectedItem = tvtemp.HitTest(X, Y)
70        TreePositionX = coord.X
80        TreePositionY = coord.Y
90        Select Case Left(tvtemp.SelectedItem.Key, 2)
          Case "L0"
              '***If first node in tree (ie Caseid node)
              '***if its not originally validated and validated option is not selected then

100           If optReport(1) <> True And Not Validated Then
110               If Not Search Then
                      '***if consultant don't want any menu items on this node to appear
120                   If UCase$(UserMemberOf) <> "CONSULTANT" And _
                         UCase$(UserMemberOf) <> "SPECIALIST REGISTRAR" Then

                          '***if the case is in cut up then show cut up details menu item
130                       If Not CaseInCutUp Then

140                           mnuAddCutUp.Visible = False
150                       Else
160                           mnuAddCutUp.Visible = True
170                       End If
180                       mnuAddTissueType.Visible = True
190                       mnuOpen.Visible = False

200                       mnuDisposeCase.Visible = False

                          '***view the menuitems in mnuPopUpLevel1 through menu editor
210                       PopupMenu mnuPopUpLevel1
220                   End If
230               Else
240                   mnuOpen.Visible = True
250                   mnuAddCutUp.Visible = False
260                   mnuAddTissueType.Visible = False
270                   PopupMenu mnuPopUpLevel1
280               End If
290           Else
                  '***view the menuitems in mnuPopUpLevel1 through menu editor
                  'If Tull or Port or Mull Autopsy case then show "Dispose Case" menu option
300               If Left(txtCaseId, 2) = "TA" Or Left(txtCaseId, 2) = "PA" Or Left(txtCaseId, 2) = "MA" Then
                      'If 14 days since inital Authorisation have passed
                      'And NOT previously disposed
310                   If DaysHavePassedSinceAuthorisation(txtCaseId, 14) And Not CaseAllDisposed(txtCaseId) Then
320                       mnuDisposeCase.Visible = True
330                   Else
340                       mnuDisposeCase.Visible = False
350                   End If
360               Else    'else don't
370                   mnuDisposeCase.Visible = False
380               End If

390               mnuAddCutUp.Visible = False
400               mnuOpen.Visible = False
410               mnuAddTissueType.Visible = False
420               PopupMenu mnuPopUpLevel1
430           End If
440       Case "L1"
450           If Not Search Then
                  '***If cytology don't add blocks or frozen section
460               If Left(CaseNo, 1) <> "C" Then
470                   mnuSingleBlock.Visible = True
480                   mnuMultipleBlocks.Visible = True
490                   mnuFrozenSection.Visible = True
500               Else
510                   mnuSingleBlock.Visible = False
520                   mnuMultipleBlocks.Visible = False
530                   mnuFrozenSection.Visible = False
540               End If
                  '***if consultant don't want any menu items on this node to appear
550               If UCase$(UserMemberOf) <> "CONSULTANT" And _
                     UCase$(UserMemberOf) <> "SPECIALIST REGISTRAR" Then
560                   Set oNode = tvtemp.SelectedItem

570                   mnuAllEmbedded.Visible = True
580                   mnuTouchPrep.Visible = True
590                   If Not (oNode Is Nothing) Then
600                       If Right(oNode.Text, 4) = "(AE)" Then
610                           mnuAllEmbedded.Checked = True
620                       End If
630                       If oNode.Children Then
640                           Set oChildNode = tvtemp.SelectedItem.Child
                              '***If one touch prep already added can't add any more so menu item does not appear
650                           Do Until oChildNode Is Nothing
660                               If UCase(Trim(oChildNode.Text)) = "Touch Prep" Then
670                                   mnuTouchPrep.Visible = False
680                               End If
690                               Set oChildNode = oChildNode.Next
700                           Loop
710                       End If
720                   End If

                      '***if Manager or Scientist show the delete menu item so they can delete tissuetype node
730                   If UCase$(UserMemberOf) = "MANAGER" Or _
                         UCase$(UserMemberOf) = "IT MANAGER" Or _
                         UCase$(UserMemberOf) = "SCIENTIST" Then
740                       mnuDelTissueType.Visible = True
750                       mnuSeperator3.Visible = True
760                   Else
770                       mnuDelTissueType.Visible = False
780                       mnuSeperator3.Visible = False
790                   End If

800                   If optReport(1) = True Or Validated Then
810                       mnuDelTissueType.Visible = False
820                       mnuSeperator3.Visible = False
830                       mnuEditTissueType.Visible = False
840                       mnuReferral.Visible = False
850                       mnuSeparator6.Visible = False
860                   End If

870                   PopupMenu mnuPopupLevel2
880               End If
890           End If
900       Case "L2"
910           If Not Search Then
                  '***if frozen section then popup same menu item as if you were on the tissue type node
920               If InStr(1, tvtemp.SelectedItem.Text, "Frozen Section") Then
930                   If UCase$(UserMemberOf) <> "CONSULTANT" And _
                         UCase$(UserMemberOf) <> "SPECIALIST REGISTRAR" Then
940                       mnuFrozenSection.Visible = False
950                       mnuTouchPrep.Visible = False
960                       mnuAllEmbedded.Visible = False
970                       mnuEditTissueType.Visible = False
980                       mnuSeparator6.Visible = False
990                       If Left(CaseNo, 1) <> "C" Then
1000                          mnuSingleBlock.Visible = True
1010                          mnuMultipleBlocks.Visible = True
1020                      Else
1030                          mnuSingleBlock.Visible = False
1040                          mnuMultipleBlocks.Visible = False
1050                      End If

1060                      If optReport(1) = True Or Validated Then
1070                          mnuDelTissueType.Visible = False
1080                          mnuSeperator3.Visible = False
1090                          mnuEditTissueType.Visible = False
1100                          mnuReferral.Visible = False
1110                          mnuSeparator6.Visible = False
1120                      End If

1130                      PopupMenu mnuPopupLevel2
1140                  End If
                      '***if touch prep then popup same menu item as if you were on the tissue type node
1150              ElseIf UCase(Trim(tvtemp.SelectedItem.Text)) = "Touch Prep" Then
1160                  If UCase$(UserMemberOf) <> "CONSULTANT" And _
                         UCase$(UserMemberOf) <> "SPECIALIST REGISTRAR" Then
1170                      mnuFrozenSection.Visible = False
1180                      mnuTouchPrep.Visible = False
1190                      mnuAllEmbedded.Visible = False
1200                      If Left(CaseNo, 1) <> "C" Then
1210                          mnuSingleBlock.Visible = True
1220                          mnuMultipleBlocks.Visible = True
1230                      Else
1240                          mnuSingleBlock.Visible = False
1250                          mnuMultipleBlocks.Visible = False
1260                      End If
1270                      mnuEditTissueType.Visible = False
1280                      mnuSeparator6.Visible = False

1290                      If optReport(1) = True Or Validated Then
1300                          mnuDelTissueType.Visible = False
1310                          mnuSeperator3.Visible = False
1320                          mnuEditTissueType.Visible = False
1330                          mnuReferral.Visible = False
1340                          mnuSeparator6.Visible = False
1350                      End If

1360                      PopupMenu mnuPopupLevel2
1370                  End If
1380              Else
                      '***if the selected item is last sibling and a manager or scientist
                      '***then allow to delete
1390                  If tvtemp.SelectedItem = tvtemp.SelectedItem.LastSibling Then
1400                      If UCase$(UserMemberOf) = "MANAGER" Or _
                             UCase$(UserMemberOf) = "IT MANAGER" Or _
                             UCase$(UserMemberOf) = "SCIENTIST" Then
1410                          mnuDelBlock.Visible = True
1420                          mnuSeperator7.Visible = True
1430                      Else
1440                          mnuDelBlock.Visible = False
1450                          mnuSeperator7.Visible = False
1460                      End If
1470                  Else
1480                      mnuDelBlock.Visible = False
1490                      mnuSeperator7.Visible = False
1500                  End If

1510                  If UCase$(UserMemberOf) <> "CONSULTANT" And _
                         UCase$(UserMemberOf) <> "SPECIALIST REGISTRAR" Then

1520                      mnuAddExtraLevels.Visible = False

1530                      If optReport(1) = True Or Validated Then
1540                          mnuDelBlock.Visible = False
1550                          mnuPrnBlockNumber.Visible = False
1560                          mnuAddExtraLevels.Visible = False
1570                          mnuSeperator7.Visible = False
1580                      End If

1590                      If InStr(tvtemp.SelectedItem.Text, "Block") Then
1600                          mnuAddControlLevel3.Visible = False
1610                          mnuSingleSlideLevel3.Visible = True
1620                          mnuMultipleSlidesLevel3.Visible = True
1630                          mnuSeperator8.Visible = True
1640                          mnuRoutineStainLevel3.Visible = True
1650                          mnuSpecialStainLevel3.Visible = True
1660                          mnuImmunoStainLevel3.Visible = True
1670                          mnuNoOfLevelsLevel3.Visible = False
1680                      Else
1690                          mnuNoOfLevelsLevel3.Visible = True
1700                          mnuSingleSlideLevel3.Visible = False
1710                          mnuMultipleSlidesLevel3.Visible = False
1720                          If tvtemp.SelectedItem.Children <> 0 Then
1730                              If tvtemp.SelectedItem.Children = 1 Then
1740                                  If UCase(Trim(tvtemp.SelectedItem.Child.Text)) <> "Control" Then

1750                                      mnuSeperator8.Visible = False
1760                                      mnuRoutineStainLevel3.Visible = False
1770                                      mnuSpecialStainLevel3.Visible = False
1780                                      mnuImmunoStainLevel3.Visible = False
1790                                  End If
1800                              Else
1810                                  mnuSeperator8.Visible = False
1820                                  mnuRoutineStainLevel3.Visible = False
1830                                  mnuSpecialStainLevel3.Visible = False
1840                                  mnuImmunoStainLevel3.Visible = False
1850                              End If
1860                          Else
1870                              mnuSeperator8.Visible = True
1880                              mnuRoutineStainLevel3.Visible = True
1890                              mnuSpecialStainLevel3.Visible = True
1900                              mnuImmunoStainLevel3.Visible = True
1910                          End If


1920                          Set oNode = tvtemp.SelectedItem
1930                          mnuAddControlLevel3.Visible = True
1940                          If Not (oNode Is Nothing) Then
1950                              If oNode.Children Then
1960                                  Set oChildNode = tvtemp.SelectedItem.Child

1970                                  Do Until oChildNode Is Nothing
1980                                      If UCase(Trim(oChildNode.Text)) = "Control" Then
1990                                          mnuAddControlLevel3.Visible = False
2000                                      End If
2010                                      Set oChildNode = oChildNode.Next
2020                                  Loop
2030                              End If
2040                          End If
2050                      End If





2060                      PopupMenu mnuPopupLevel3
2070                  ElseIf Left(CaseNo, 1) <> "C" Then
2080                      If bWithPathologist Then
2090                          mnuAddExtraLevels.Visible = True
2100                      Else
2110                          mnuAddExtraLevels.Visible = False
2120                      End If

2130                      If optReport(1) = True Or Validated Then
2140                          mnuDelBlock.Visible = False
2150                          mnuPrnBlockNumber.Visible = False
2160                          mnuAddExtraLevels.Visible = False
2170                          mnuSeperator7.Visible = False
2180                      End If

2190                      If InStr(1, tvtemp.SelectedItem.Text, "Block") Then
2200                          mnuAddControlLevel3.Visible = False
2210                          mnuSingleSlideLevel3.Visible = True
2220                          mnuMultipleSlidesLevel3.Visible = True
2230                          mnuNoOfLevelsLevel3.Visible = False
2240                      Else
2250                          mnuNoOfLevelsLevel3.Visible = True
2260                          mnuSingleSlideLevel3.Visible = False
2270                          mnuMultipleSlidesLevel3.Visible = False

2280                          Set oNode = tvtemp.SelectedItem
2290                          mnuAddControlLevel3.Visible = True
2300                          If Not (oNode Is Nothing) Then
2310                              If oNode.Children Then
2320                                  Set oChildNode = tvtemp.SelectedItem.Child

2330                                  Do Until oChildNode Is Nothing
2340                                      If UCase(Trim(oChildNode.Text)) = "Control" Then
2350                                          mnuAddControlLevel3.Visible = False
2360                                      End If
2370                                      Set oChildNode = oChildNode.Next
2380                                  Loop
2390                              End If
2400                          End If
2410                      End If

2420                      PopupMenu mnuPopupLevel3
2430                  End If
2440              End If
2450          End If
2460      Case "L3"
2470          If Not Search Then
2480              If UCase$(UserMemberOf) <> "CONSULTANT" And _
                     UCase$(UserMemberOf) <> "SPECIALIST REGISTRAR" Then

                      '***if on a block that is a frozen section block then show menu item that would appear on a normal block
2490                  If InStr(1, tvtemp.SelectedItem.Parent.Text, "Frozen Section") Then
2500                      If tvtemp.SelectedItem = tvtemp.SelectedItem.LastSibling Then
2510                          If UCase$(UserMemberOf) = "MANAGER" Or _
                                 UCase$(UserMemberOf) = "IT MANAGER" Or _
                                 UCase$(UserMemberOf) = "SCIENTIST" Then
2520                              mnuDelBlock.Visible = True
2530                              mnuSeperator7.Visible = True
2540                          Else
2550                              mnuDelBlock.Visible = False
2560                              mnuSeperator7.Visible = False
2570                          End If
2580                      Else
2590                          mnuDelBlock.Visible = False
2600                          mnuSeperator7.Visible = False
2610                      End If
2620                      mnuAddExtraLevels.Visible = False

2630                      If optReport(1) = True Or Validated Then
2640                          mnuDelBlock.Visible = False
2650                          mnuPrnBlockNumber.Visible = False
2660                          mnuAddExtraLevels.Visible = False
2670                          mnuSeperator7.Visible = False
2680                      End If

2690                      If InStr(1, tvtemp.SelectedItem.Text, "Block") Then
2700                          mnuAddControlLevel3.Visible = False
2710                          mnuSingleSlideLevel3.Visible = True
2720                          mnuMultipleSlidesLevel3.Visible = True
2730                          mnuSeperator8.Visible = True
2740                          mnuRoutineStainLevel3.Visible = True
2750                          mnuSpecialStainLevel3.Visible = True
2760                          mnuImmunoStainLevel3.Visible = True
2770                      Else
2780                          mnuSingleSlideLevel3.Visible = False
2790                          mnuMultipleSlidesLevel3.Visible = False
2800                          If tvtemp.SelectedItem.Children <> 0 Then
2810                              If tvtemp.SelectedItem.Children = 1 Then

2820                                  If UCase(Trim(tvtemp.SelectedItem.Child.Text)) <> "Control" Then

2830                                      mnuSeperator8.Visible = False
2840                                      mnuRoutineStainLevel3.Visible = False
2850                                      mnuSpecialStainLevel3.Visible = False
2860                                      mnuImmunoStainLevel3.Visible = False
2870                                  End If
2880                              Else
2890                                  mnuSeperator8.Visible = False
2900                                  mnuRoutineStainLevel3.Visible = False
2910                                  mnuSpecialStainLevel3.Visible = False
2920                                  mnuImmunoStainLevel3.Visible = False
2930                              End If
2940                          Else
2950                              mnuSeperator8.Visible = True
2960                              mnuRoutineStainLevel3.Visible = True
2970                              mnuSpecialStainLevel3.Visible = True
2980                              mnuImmunoStainLevel3.Visible = True
2990                          End If


3000                          Set oNode = tvtemp.SelectedItem
3010                          mnuAddControlLevel3.Visible = True
3020                          If Not (oNode Is Nothing) Then
3030                              If oNode.Children Then
3040                                  Set oChildNode = tvtemp.SelectedItem.Child
                                      '***if control added then don't allow to add another one
3050                                  Do Until oChildNode Is Nothing
3060                                      If UCase(Trim(oChildNode.Text)) = "Control" Then
3070                                          mnuAddControlLevel3.Visible = False
3080                                      End If
3090                                      Set oChildNode = oChildNode.Next
3100                                  Loop
3110                              End If
3120                          End If
3130                      End If

3140                      PopupMenu mnuPopupLevel3
3150                  ElseIf tvtemp.SelectedItem.Children = 0 Then
3160                      mnuRoutineStainLevel4.Visible = True
3170                      mnuSpecialStainLevel4.Visible = True
3180                      mnuImmunoStainLevel4.Visible = True
3190                      If tvtemp.SelectedItem = tvtemp.SelectedItem.LastSibling Then
3200                          If UCase$(UserMemberOf) = "MANAGER" Or _
                                 UCase$(UserMemberOf) = "IT MANAGER" Or _
                                 UCase$(UserMemberOf) = "SCIENTIST" Then
3210                              mnuDelSlide.Visible = True
3220                              mnuSeperator8.Visible = True
3230                          Else
3240                              mnuDelSlide.Visible = False
3250                              mnuSeperator8.Visible = False
3260                          End If
3270                      Else
3280                          mnuDelSlide.Visible = False
3290                          mnuSeperator8.Visible = False
3300                      End If

3310                      If optReport(1) = True Or Validated Then
3320                          mnuDelSlide.Visible = False
3330                          mnuSeperator8.Visible = False
3340                      End If

3350                      Set oNode = tvtemp.SelectedItem
3360                      mnuAddControlLevel4.Visible = True
3370                      If Not (oNode Is Nothing) Then
3380                          If oNode.Children Then
3390                              Set oChildNode = tvtemp.SelectedItem.Child

3400                              Do Until oChildNode Is Nothing
3410                                  If UCase(Trim(oChildNode.Text)) = "Control" Then
3420                                      mnuAddControlLevel4.Visible = False
3430                                  End If
3440                                  Set oChildNode = oChildNode.Next
3450                              Loop
3460                          End If
3470                      End If


3480                      PopupMenu mnuPopupLevel4
3490                  Else
3500                      If tvtemp.SelectedItem = tvtemp.SelectedItem.LastSibling Then
3510                          If UCase$(UserMemberOf) = "MANAGER" Or _
                                 UCase$(UserMemberOf) = "IT MANAGER" Or _
                                 UCase$(UserMemberOf) = "SCIENTIST" Then
3520                              mnuDelSlide.Visible = True
3530                              mnuSeperator8.Visible = True
3540                          Else
3550                              mnuDelSlide.Visible = False
3560                              mnuSeperator8.Visible = False
3570                          End If

3580                          If tvtemp.SelectedItem.Children <> 0 Then
3590                              If tvtemp.SelectedItem.Children = 1 Then
3600                                  If UCase(Trim(tvtemp.SelectedItem.Child.Text)) = "Control" Then
3610                                      mnuRoutineStainLevel4.Visible = True
3620                                      mnuSpecialStainLevel4.Visible = True
3630                                      mnuImmunoStainLevel4.Visible = True
3640                                      mnuAddControlLevel4.Visible = False
3650                                  Else
3660                                      mnuSeperator8.Visible = False
3670                                      mnuRoutineStainLevel4.Visible = False
3680                                      mnuSpecialStainLevel4.Visible = False
3690                                      mnuImmunoStainLevel4.Visible = False
3700                                  End If
3710                              Else
3720                                  mnuSeperator8.Visible = False
3730                                  mnuRoutineStainLevel4.Visible = False
3740                                  mnuSpecialStainLevel4.Visible = False
3750                                  mnuImmunoStainLevel4.Visible = False
3760                              End If


3770                              Set oNode = tvtemp.SelectedItem
3780                              mnuAddControlLevel4.Visible = True
3790                              If Not (oNode Is Nothing) Then
3800                                  If oNode.Children Then
3810                                      Set oChildNode = tvtemp.SelectedItem.Child

3820                                      Do Until oChildNode Is Nothing
3830                                          If UCase(Trim(oChildNode.Text)) = "Control" Then
3840                                              mnuAddControlLevel4.Visible = False
3850                                          End If
3860                                          Set oChildNode = oChildNode.Next
3870                                      Loop
3880                                  End If
3890                              End If





3900                              If optReport(1) = True And Validated Then
3910                                  mnuDelSlide.Visible = False
3920                                  mnuSeperator8.Visible = False
3930                              End If
3940                              PopupMenu mnuPopupLevel4

3950                          Else
3960                              mnuDelSlide.Visible = False
3970                              mnuSeperator8.Visible = False
3980                              mnuNoOfLevelsLevel4.Visible = False

3990                              Set oNode = tvtemp.SelectedItem
4000                              mnuAddControlLevel4.Visible = True
4010                              If Not (oNode Is Nothing) Then
4020                                  If oNode.Children Then
4030                                      Set oChildNode = tvtemp.SelectedItem.Child

4040                                      Do Until oChildNode Is Nothing
4050                                          If UCase(Trim(oChildNode.Text)) = "Control" Then
4060                                              mnuAddControlLevel4.Visible = False
4070                                          End If
4080                                          Set oChildNode = oChildNode.Next
4090                                      Loop
4100                                  End If
4110                              End If
4120                              PopupMenu mnuPopupLevel4
4130                          End If


4140                      Else
4150                          If optReport(1) <> True And Not Validated Then
4160                              mnuSeperator8.Visible = False
4170                              mnuRoutineStainLevel4.Visible = False
4180                              mnuSpecialStainLevel4.Visible = False
4190                              mnuImmunoStainLevel4.Visible = False
4200                              mnuDelSlide.Visible = False

4210                              Set oNode = tvtemp.SelectedItem
4220                              mnuAddControlLevel4.Visible = True
4230                              If Not (oNode Is Nothing) Then
4240                                  If oNode.Children Then
4250                                      Set oChildNode = tvtemp.SelectedItem.Child

4260                                      Do Until oChildNode Is Nothing
4270                                          If UCase(Trim(oChildNode.Text)) = "Control" Then
4280                                              mnuAddControlLevel4.Visible = False
4290                                          End If
4300                                          Set oChildNode = oChildNode.Next
4310                                      Loop
4320                                  End If
4330                              End If
4340                              PopupMenu mnuPopupLevel4
4350                          End If
4360                      End If


4370                  End If
4380              Else

4390                  If InStr(1, tvtemp.SelectedItem.Parent.Text, "Frozen Section") Then
4400                      mnuDelBlock.Visible = False
4410                      mnuSeperator7.Visible = False
4420                      If bWithPathologist Then
4430                          mnuAddExtraLevels.Visible = True
4440                      Else
4450                          mnuAddExtraLevels.Visible = False
4460                      End If

4470                      If optReport(1) = True Or Validated Then
4480                          mnuDelBlock.Visible = False
4490                          mnuPrnBlockNumber.Visible = False
4500                          mnuAddExtraLevels.Visible = False
4510                          mnuSeperator7.Visible = False
4520                      End If

4530                      If InStr(1, tvtemp.SelectedItem.Text, "Block") Then
4540                          mnuAddControlLevel3.Visible = False
4550                          mnuSingleSlideLevel3.Visible = True
4560                          mnuMultipleSlidesLevel3.Visible = True
4570                          mnuSeperator8.Visible = True
4580                          mnuRoutineStainLevel3.Visible = True
4590                          mnuSpecialStainLevel3.Visible = True
4600                          mnuImmunoStainLevel3.Visible = True
4610                          mnuNoOfLevelsLevel3.Visible = False
4620                      Else
4630                          mnuNoOfLevelsLevel3.Visible = True
4640                          mnuSingleSlideLevel3.Visible = False
4650                          mnuMultipleSlidesLevel3.Visible = False
4660                          If tvtemp.SelectedItem.Children <> 0 Then
4670                              If tvtemp.SelectedItem.Children = 1 Then
4680                                  If UCase(Trim(tvtemp.SelectedItem.Child.Text)) <> "Control" Then

4690                                      mnuSeperator8.Visible = False
4700                                      mnuRoutineStainLevel3.Visible = False
4710                                      mnuSpecialStainLevel3.Visible = False
4720                                      mnuImmunoStainLevel3.Visible = False
4730                                  End If
4740                              Else
4750                                  mnuSeperator8.Visible = False
4760                                  mnuRoutineStainLevel3.Visible = False
4770                                  mnuSpecialStainLevel3.Visible = False
4780                                  mnuImmunoStainLevel3.Visible = False
4790                              End If
4800                          Else
4810                              mnuSeperator8.Visible = True
4820                              mnuRoutineStainLevel3.Visible = True
4830                              mnuSpecialStainLevel3.Visible = True
4840                              mnuImmunoStainLevel3.Visible = True
4850                          End If


4860                          Set oNode = tvtemp.SelectedItem
4870                          mnuAddControlLevel3.Visible = True
4880                          If Not (oNode Is Nothing) Then
4890                              If oNode.Children Then
4900                                  Set oChildNode = tvtemp.SelectedItem.Child

4910                                  Do Until oChildNode Is Nothing
4920                                      If UCase(Trim(oChildNode.Text)) = "Control" Then
4930                                          mnuAddControlLevel3.Visible = False
4940                                      End If
4950                                      Set oChildNode = oChildNode.Next
4960                                  Loop
4970                              End If
4980                          End If
4990                      End If


5000                      PopupMenu mnuPopupLevel3
5010                  End If

5020              End If
5030          End If
5040      Case "L4"
5050          If Not Search Then
5060              If UCase$(UserMemberOf) <> "CONSULTANT" And _
                     UCase$(UserMemberOf) <> "SPECIALIST REGISTRAR" Then
5070                  If InStr(1, tvtemp.SelectedItem.Parent.Parent.Text, "Frozen Section") Then
5080                      If tvtemp.SelectedItem.Children = 0 Then
5090                          mnuRoutineStainLevel4.Visible = True
5100                          mnuSpecialStainLevel4.Visible = True
5110                          mnuImmunoStainLevel4.Visible = True
5120                          If tvtemp.SelectedItem = tvtemp.SelectedItem.LastSibling Then
5130                              If UCase$(UserMemberOf) = "MANAGER" Or _
                                     UCase$(UserMemberOf) = "IT MANAGER" Or _
                                     UCase$(UserMemberOf) = "SCIENTIST" Then
5140                                  mnuDelSlide.Visible = True
5150                                  mnuSeperator8.Visible = True
5160                              Else
5170                                  mnuDelSlide.Visible = False
5180                                  mnuSeperator8.Visible = False
5190                              End If
5200                          Else
5210                              mnuDelSlide.Visible = False
5220                              mnuSeperator8.Visible = False
5230                          End If

5240                          If optReport(1) = True Or Validated Then
5250                              mnuDelSlide.Visible = False
5260                              mnuSeperator8.Visible = False
5270                          End If

5280                          Set oNode = tvtemp.SelectedItem
5290                          mnuAddControlLevel4.Visible = True
5300                          If Not (oNode Is Nothing) Then
5310                              If oNode.Children Then
5320                                  Set oChildNode = tvtemp.SelectedItem.Child

5330                                  Do Until oChildNode Is Nothing
5340                                      If UCase(Trim(oChildNode.Text)) = "Control" Then
5350                                          mnuAddControlLevel4.Visible = False
5360                                      End If
5370                                      Set oChildNode = oChildNode.Next
5380                                  Loop
5390                              End If
5400                          End If

5410                          PopupMenu mnuPopupLevel4
5420                      ElseIf tvtemp.SelectedItem = tvtemp.SelectedItem.LastSibling Then
5430                          If UCase$(UserMemberOf) = "MANAGER" Or _
                                 UCase$(UserMemberOf) = "IT MANAGER" Or _
                                 UCase$(UserMemberOf) = "SCIENTIST" Then
5440                              mnuDelSlide.Visible = True
5450                              mnuSeperator8.Visible = True
5460                          Else
5470                              mnuDelSlide.Visible = False
5480                              mnuSeperator8.Visible = False
5490                          End If

5500                          If tvtemp.SelectedItem.Children <> 0 Then
5510                              mnuSeperator8.Visible = False
5520                              mnuRoutineStainLevel4.Visible = False
5530                              mnuSpecialStainLevel4.Visible = False
5540                              mnuImmunoStainLevel4.Visible = False

5550                              Set oNode = tvtemp.SelectedItem
5560                              mnuAddControlLevel4.Visible = True
5570                              If Not (oNode Is Nothing) Then
5580                                  If oNode.Children Then
5590                                      Set oChildNode = tvtemp.SelectedItem.Child

5600                                      Do Until oChildNode Is Nothing
5610                                          If UCase(Trim(oChildNode.Text)) = "Control" Then
5620                                              mnuAddControlLevel4.Visible = False
5630                                          End If
5640                                          Set oChildNode = oChildNode.Next
5650                                      Loop
5660                                  End If
5670                              End If
5680                              If optReport(1) = True And Validated Then
5690                                  mnuDelSlide.Visible = False
5700                              End If
5710                              PopupMenu mnuPopupLevel4
5720                          Else
5730                              mnuDelSlide.Visible = False
5740                              mnuSeperator8.Visible = False
5750                              mnuNoOfLevelsLevel4.Visible = False

5760                              Set oNode = tvtemp.SelectedItem
5770                              mnuAddControlLevel4.Visible = True
5780                              If Not (oNode Is Nothing) Then
5790                                  If oNode.Children Then
5800                                      Set oChildNode = tvtemp.SelectedItem.Child

5810                                      Do Until oChildNode Is Nothing
5820                                          If UCase(Trim(oChildNode.Text)) = "Control" Then
5830                                              mnuAddControlLevel4.Visible = False
5840                                          End If
5850                                          Set oChildNode = oChildNode.Next
5860                                      Loop
5870                                  End If
5880                              End If
5890                              PopupMenu mnuPopupLevel4
5900                          End If

5910                      Else
5920                          If optReport(1) <> True And Not Validated Then
5930                              mnuSeperator8.Visible = False
5940                              mnuRoutineStainLevel4.Visible = False
5950                              mnuSpecialStainLevel4.Visible = False
5960                              mnuImmunoStainLevel4.Visible = False
5970                              mnuDelSlide.Visible = False

5980                              Set oNode = tvtemp.SelectedItem
5990                              mnuAddControlLevel4.Visible = True
6000                              If Not (oNode Is Nothing) Then
6010                                  If oNode.Children Then
6020                                      Set oChildNode = tvtemp.SelectedItem.Child

6030                                      Do Until oChildNode Is Nothing
6040                                          If UCase(Trim(oChildNode.Text)) = "Control" Then
6050                                              mnuAddControlLevel4.Visible = False
6060                                          End If
6070                                          Set oChildNode = oChildNode.Next
6080                                      Loop
6090                                  End If
6100                              End If
6110                              PopupMenu mnuPopupLevel4
6120                          End If
6130                      End If

6140                  Else
6150                      If optReport(1) <> True And Not Validated Then
6160                          If UCase$(UserMemberOf) = "MANAGER" Or _
                                 UCase$(UserMemberOf) = "IT MANAGER" Or _
                                 UCase$(UserMemberOf) = "SCIENTIST" Then
6170                              PopupMenu mnuPopupLevel5
6180                          End If
6190                      End If
6200                  End If
6210              End If
6220          End If
6230      Case "L5"
6240          If optReport(1) <> True And Not Validated Then
6250              If Not Search Then
6260                  If UCase$(UserMemberOf) = "MANAGER" Or _
                         UCase$(UserMemberOf) = "IT MANAGER" Or _
                         UCase$(UserMemberOf) = "SCIENTIST" Then
6270                      PopupMenu mnuPopupLevel5
6280                  End If
6290              End If
6300          End If
6310      End Select
6320  End If

6330  Exit Sub

CallTreePopupMenu_Error:

      Dim strES As String
      Dim intEL As Integer

6340  intEL = Erl
6350  strES = Err.Description
6360  LogError "frmWorkSheet", "CallTreePopupMenu", intEL, strES

End Sub

Private Sub InitializeGridTracker()

      Dim i As Integer

10    For i = 0 To 3
20        With grdTracker(i)
30            .Clear
40            .Rows = 2: .FixedRows = 1
50            .Cols = 11: .FixedCols = 0
60            .Rows = 1
70            .Font.Size = fgcFontSize
80            .Font.name = fgcFontName
90            .ForeColor = fgcForeColor
100           .BackColor = fgcBackColor
110           .ForeColorFixed = fgcForeColorFixed
120           .BackColorFixed = fgcBackColorFixed
130           .ScrollBars = flexScrollBarBoth
              '<Type                  |<Destination    |<Out     |Received
140           .TextMatrix(0, 0) = "Description": .ColWidth(0) = 900: .ColAlignment(0) = flexAlignLeftCenter
150           .TextMatrix(0, 1) = "ID": .ColWidth(1) = 300: .ColAlignment(1) = flexAlignLeftCenter
160           .TextMatrix(0, 2) = "Sent": .ColWidth(2) = 950: .ColAlignment(2) = flexAlignLeftCenter
170           .TextMatrix(0, 3) = "Returned": .ColWidth(3) = 950: .ColAlignment(3) = flexAlignLeftCenter
180           .TextMatrix(0, 4) = "Code": .ColWidth(4) = 0: .ColAlignment(4) = flexAlignLeftCenter
190           .TextMatrix(0, 5) = "Unique ID": .ColWidth(5) = 0: .ColAlignment(5) = flexAlignLeftCenter
200           .TextMatrix(0, 6) = "Completed": .ColWidth(6) = 775: .ColAlignment(6) = flexAlignLeftCenter
210           .TextMatrix(0, 7) = "Type": .ColWidth(7) = 0: .ColAlignment(7) = flexAlignLeftCenter
220           .TextMatrix(0, 8) = "Referred To": .ColWidth(8) = 0: .ColAlignment(8) = flexAlignLeftCenter
230           .TextMatrix(0, 9) = "Reason of Referal": .ColWidth(9) = 0: .ColAlignment(9) = flexAlignLeftCenter
240           .TextMatrix(0, 10) = "Details": .ColWidth(10) = 600: .ColAlignment(10) = flexAlignLeftCenter

250       End With
260   Next
End Sub

Private Sub txtContainerLabel_KeyPress(KeyAscii As Integer)
10    If txtContainerLabel.Locked = True Then
20        DataChanged = True
30    End If
End Sub

Private Sub txtFindCode_KeyPress(KeyAscii As Integer)
10    Set frmList.txtCode = txtFindCode
20    Set frmList.txtDescription = txtFindDescription
30    frmList.SearchByCode = True
40    frmList.Show
50    frmList.Move Me.Left + txtFindCode.Left + 50, Me.Top + fraWorkSheet.Top + txtFindCode.Top + 625

End Sub

Private Sub txtFindDescription_KeyPress(KeyAscii As Integer)
10    Set frmList.txtCode = txtFindCode
20    Set frmList.txtDescription = txtFindDescription
30    frmList.SearchByCode = False
40    frmList.Show
50    frmList.Move Me.Left + txtFindCode.Left + 50, Me.Top + fraWorkSheet.Top + txtFindCode.Top + 625
End Sub



Private Sub txtGross_DblClick()
10    With frmRichText
20        If Validated Or bLocked Then
30            .rtb.Locked = True
40        End If
50        .rtbTextBox = "GROSS"
60        .Show 1
70    End With
End Sub

Private Sub txtGross_KeyPress(KeyAscii As Integer)
10    If txtGross.Locked = False Then
20        DataChanged = True
30    End If
End Sub

Private Sub txtGross_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    If Button = vbRightButton Then
20        PopupMenu mnuPopupFormatGrossText
30    End If
End Sub

Private Sub txtMCode_KeyPress(KeyAscii As Integer)
10    frmList.PrevCode = txtMCode
20    frmList.PrevDesc = txtMDescription
30    Set frmList.txtCode = txtMCode
40    Set frmList.txtDescription = txtMDescription
50    frmList.SearchByCode = True
60    frmList.ListType = "M"
70    frmList.Show
80    frmList.Move Me.Left + fraWorkSheet.Left + txtMCode.Left + 50, Me.Top + fraWorkSheet.Top + txtMCode.Top + 625
90    If KeyAscii = 13 Then
100       KeyAscii = 0
110   End If
End Sub

Private Sub txtMDescription_KeyPress(KeyAscii As Integer)
10    frmList.PrevCode = txtMCode
20    frmList.PrevDesc = txtMDescription
30    Set frmList.txtCode = txtMCode
40    Set frmList.txtDescription = txtMDescription
50    frmList.SearchByCode = False
60    frmList.ListType = "M"
70    frmList.Show
80    frmList.Move Me.Left + fraWorkSheet.Left + txtMCode.Left + 50, Me.Top + fraWorkSheet.Top + txtMCode.Top + 625
90    If KeyAscii = 13 Then
100       KeyAscii = 0
110   End If
End Sub

Private Sub txtMicro_DblClick()
10    With frmRichText
20        If Validated Or bLocked Then
30            .rtb.Locked = True
40        End If
50        .rtbTextBox = "MICRO"
60        .Show 1
70    End With
End Sub

Private Sub txtMicro_KeyPress(KeyAscii As Integer)
10    If txtMicro.Locked = False Then
20        DataChanged = True
30    End If
End Sub

Private Sub txtMicro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    If Button = vbRightButton Then
20        PopupMenu mnuPopupFormatMicroText
30    End If
End Sub




Private Sub txtNOS_KeyPress(KeyAscii As Integer)
10    If txtNOS.Locked = False Then
20        DataChanged = True
30    End If
End Sub

Private Sub txtPatientId_LostFocus()
      Dim tb As New Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo txtPatientId_LostFocus_Error

20    If UCase(sysOptCaseIdValidation(0)) = "LIMERICK" Then
30        ConnectToLimNetacquireDb

40        sql = "SELECT * FROM PatientIfs "

50        If cmbPatientId = "NOPAS" Then
60            sql = sql & "WHERE  Nopas = N'" & txtPatientId & "'"
70        ElseIf cmbPatientId = "MRN" Then
80            sql = sql & "WHERE  Mrn = N'" & txtPatientId & "'"
90        Else
100           sql = sql & "WHERE  AandE = N'" & txtPatientId & "'"
110       End If

120       Set tb = New Recordset
130       Set tb = Cnxn(1).Execute(sql)

140       If Not tb.EOF Then
150           If txtCaseId <> "" Then
160               fraDemographics.Visible = True
170               lblPatientName.Caption = tb!PatName & ""
180               lblSex.Caption = "(" & tb!Sex & ")"
190               lblPatientAddress1.Caption = tb!Address0 & ""
200               lblPatientAddress2.Caption = tb!Address1 & ""
210               lblPatientAddress3.Caption = tb!Address2 & ""
220               lblPatientBorn.Caption = "Born " & tb!DoB & ""

230               lblPatientWard.Caption = tb!Ward & ""
240               lblPatientDoctor.Caption = tb!Clinician & ""
250               lblNopas.Caption = tb!Nopas & ""
260               lblMrn.Caption = tb!MRN & ""
270               lblAandE.Caption = tb!AandE & ""
280               cmdClinicalHist.Visible = True
290               CheckClinicalHist
300               cmdComments.Visible = True
310               CheckGeneralComments
320               cmdDiscrepancyLog.Visible = True
330               CheckDiscrepancyLog
340               If Val(GetOptionSetting("DemographicEntry", "0")) <> 0 Then
350                   cmdEditDemo.Visible = True
360               End If
370               cmdCytoHist.Visible = True
380               cmdAudit.Visible = True
390               s = "Chart No:  " & lblMrn & vbCrLf & _
                      "Name:      " & lblPatientName & vbCrLf & _
                      "Sex:       " & lblSex & vbCrLf & _
                      "Address1:  " & lblPatientAddress1 & vbCrLf & _
                      "Address2:  " & lblPatientAddress2 & vbCrLf & _
                      "Address3:  " & lblPatientAddress3 & vbCrLf & _
                      "DOB:       " & lblPatientBorn & vbCrLf & _
                      "Ward:      " & lblPatientWard & vbCrLf
400               s = s & "Clinician: " & lblPatientDoctor & vbCrLf
410               mclsToolTip.ToolText(fraDemographics) = s
420           End If
430       End If
440   End If

450   Exit Sub

txtPatientId_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

460   intEL = Erl
470   strES = Err.Description
480   LogError "frmWorkSheet", "txtPatientId_LostFocus", intEL, strES, sql


End Sub

Private Sub txtPCode_KeyPress(KeyAscii As Integer)
10    frmList.PrevCode = txtPCode
20    frmList.PrevDesc = txtPDescription
30    Set frmList.txtCode = txtPCode
40    Set frmList.txtDescription = txtPDescription
50    frmList.SearchByCode = True
60    frmList.ListType = "P"
70    frmList.Show
80    frmList.Move Me.Left + fraWorkSheet.Left + txtPCode.Left + 50, Me.Top + fraWorkSheet.Top + txtPCode.Top + 625
90    If KeyAscii = 13 Then
100       KeyAscii = 0
110   End If
120   DataChanged = True
End Sub


Private Sub txtPDescription_KeyPress(KeyAscii As Integer)
10    frmList.PrevCode = txtPCode
20    frmList.PrevDesc = txtPDescription
30    Set frmList.txtCode = txtPCode
40    Set frmList.txtDescription = txtPDescription
50    frmList.SearchByCode = False
60    frmList.ListType = "P"
70    frmList.Show
80    frmList.Move Me.Left + fraWorkSheet.Left + txtPCode.Left + 50, Me.Top + fraWorkSheet.Top + txtPCode.Top + 625
90    If KeyAscii = 13 Then
100       KeyAscii = 0
110   End If
120   DataChanged = True
End Sub


Private Sub txtQCode_KeyPress(KeyAscii As Integer)
10    frmList.PrevCode = txtQCode
20    frmList.PrevDesc = txtQDescription
30    Set frmList.txtCode = txtQCode
40    Set frmList.txtDescription = txtQDescription
50    frmList.SearchByCode = True
60    frmList.ListType = "Q"
70    frmList.Show
80    frmList.Move Me.Left + fraWorkSheet.Left + txtQCode.Left + 50, Me.Top + fraWorkSheet.Top + txtQCode.Top + 625
90    If KeyAscii = 13 Then
100       KeyAscii = 0
110   End If
End Sub



Private Sub txtQDescription_KeyPress(KeyAscii As Integer)
10    frmList.PrevCode = txtQCode
20    frmList.PrevDesc = txtQDescription
30    Set frmList.txtCode = txtQCode
40    Set frmList.txtDescription = txtQDescription
50    frmList.SearchByCode = False
60    frmList.ListType = "Q"
70    frmList.Show
80    frmList.Move Me.Left + fraWorkSheet.Left + txtQCode.Left + 50, Me.Top + fraWorkSheet.Top + txtQCode.Top + 625
90    If KeyAscii = 13 Then
100       KeyAscii = 0
110   End If
End Sub


Private Sub ResetWorkSheet()

      Dim i As Integer

10    lblPatientAddress1 = ""
20    lblPatientAddress2 = ""
30    lblPatientAddress3 = ""
40    lblAge = ""
50    lblSex = ""
60    lblPatientBorn = ""
70    lblPatientDoctor = ""
80    lblPatientGP = ""
90    lblPatientName = ""
100   lblPatientWard = ""
110   DTSampleTaken.CustomFormat = " "
120   DTSampleTaken.Value = Date
130   txtSampleTakenTime = "__:__"
140   DTSampleRec.CustomFormat = " "
150   DTSampleRec.Value = Date
160   txtSampleRecTime = "__:__"
170   lblPreReportDate = ""
180   lblValReportDate = ""
190   txtPCode = ""
200   txtPDescription = ""
210   txtGross.Text = ""
220   txtMCode = ""
230   txtMDescription = ""
240   txtMicro.Text = ""
250   txtQCode = ""
260   txtQDescription = ""
270   lblNopas = ""
280   lblMrn = ""
290   lblAandE = ""
300   lblPatientGP = ""
310   lblClinicalHist = ""
320   lblGeneralComments = ""

330   InitializeGridTracker
340   For i = 1 To 3
350       grdTracker(i).ColWidth(1) = 0
360   Next
370   InitializeGridCodes grdMCodes
380   InitializeGridCodes grdTempMCode
390   InitializeGridCodes grdQCodes
400   InitializeGridAmendments
410   InitializeGridDelete
420   optState(0).Value = True
430   txtPatientId.Enabled = True
440   DataChanged = False
450   TreeChanged = False
460   optReport(0).Value = False
470   optReport(1).Value = False
480   cmdClinicalHist.Visible = False
490   cmdAudit.Visible = False
500   cmdDiscrepancyLog.Visible = False
510   fraDemographics.Visible = False
520   cmdEditDemo.Visible = False
530   cmdComments.Visible = False
540   cmdCytoHist.Visible = False
550   fraLinkedCase.Visible = False
560   cmdLinkedCaseId.Caption = ""
570   txtNOS = ""
580   txtContainerLabel = ""
End Sub
Private Sub InitializeGridAmendments()
10    With grdAmendments
20        .Clear
    .Rows = 2:
30        .Cols = 4
40        .FixedCols = 0
50        .FixedRows = 1
60        .Rows = 1
70        .ColWidth(0) = 1200
80        .ColWidth(1) = grdAmendments.Width - 1200 - 250
90        .ColWidth(2) = 0
100       .ColWidth(3) = 0
110       .TextMatrix(0, 0) = "Date"
120       .TextMatrix(0, 1) = "Description"
130       .TextMatrix(0, 2) = "Amend Id"
140       .TextMatrix(0, 3) = "Code"
150       .GridLines = flexGridNone

160   End With
End Sub
Private Sub InitializeGridDelete()
10    With grdDelete
20        .Clear
    .Rows = 2:
30        .Cols = 4
40        .FixedCols = 0
50        .FixedRows = 1
60        .Rows = 1
70        .ColWidth(0) = 1000
80        .ColWidth(1) = 1000
90        .ColWidth(2) = 1000
100       .ColWidth(3) = 0
110       .TextMatrix(0, 0) = "Unqiue ID"
120       .TextMatrix(0, 1) = "Code"
130       .TextMatrix(0, 2) = "Table"
140       .TextMatrix(0, 3) = "Path"
150       .GridLines = flexGridNone

160   End With
End Sub

Private Sub ResetSearch()
10    txtCaseId = ""
20    fraWorkSheet.Enabled = False
30    txtPatientId = ""
40    Set PrevNode = Nothing
50    tvCaseDetails.Nodes.Clear
60    cmdSearch.Enabled = True
70    TreeChanged = False

End Sub


Private Sub txtSampleRecTime_GotFocus()
10    txtSampleRecTime.SelStart = 0
20    txtSampleRecTime.SelLength = Len(txtSampleRecTime.FormattedText)

End Sub


Private Sub txtSampleRecTime_Validate(Cancel As Boolean)

10    If DTSampleRec = DTSampleTaken Then
20        If txtSampleTakenTime <> "" Then
30            If Format(txtSampleRecTime.FormattedText, "HH:mm") < Format(txtSampleTakenTime.FormattedText, "HH:mm") Then
40                frmMsgBox.Msg "SampleDate After RecDate Please Amend", , , mbExclamation
50                Cancel = True
60            End If
70        End If
80    End If


End Sub

Private Sub txtSampleTakenTime_GotFocus()
10    txtSampleTakenTime.SelStart = 0
20    txtSampleTakenTime.SelLength = Len(txtSampleTakenTime.FormattedText)

End Sub

Private Sub txtSampleTakenTime_Validate(Cancel As Boolean)

10    If DTSampleRec = DTSampleTaken Then
20        If txtSampleTakenTime <> "" Then
30            If Format(txtSampleTakenTime.FormattedText, "HH:mm") > Format(txtSampleRecTime.FormattedText, "HH:mm") Then
40                frmMsgBox.Msg "SampleDate After RecDate Please Amend", , , mbExclamation
50                Cancel = True
60            End If
70        End If
80    End If

End Sub

