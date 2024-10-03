VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmDemographics 
   Caption         =   "Demographics"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDemographics 
      ForeColor       =   &H000000FF&
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   9840
         Picture         =   "frmDemographics.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   240
         Width           =   375
      End
      Begin VB.ListBox lstClinician 
         Height          =   855
         IntegralHeight  =   0   'False
         Left            =   6120
         TabIndex        =   25
         Top             =   3360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ListBox lstWard 
         Height          =   855
         IntegralHeight  =   0   'False
         Left            =   4200
         TabIndex        =   23
         Top             =   2640
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ListBox lstGP 
         Height          =   855
         IntegralHeight  =   0   'False
         Left            =   4200
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ListBox lstCoronerClin 
         Height          =   855
         IntegralHeight  =   0   'False
         Left            =   4200
         TabIndex        =   82
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtDateOfDeath 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Date of Birth"
         Text            =   "88/88/8888"
         Top             =   4560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame fraAutopsy 
         Height          =   1335
         Left            =   4440
         TabIndex        =   19
         Top             =   3960
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton optRequest 
            Caption         =   "Paediatric"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optRequest 
            Caption         =   "Non-Coroner"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optRequest 
            Caption         =   "Coroner"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ListBox lstCounty 
         Height          =   855
         IntegralHeight  =   0   'False
         Left            =   4200
         TabIndex        =   57
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cmbSource 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox txtDOB 
         Height          =   285
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Date of Birth"
         Text            =   "88/88/8888"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtChartNo 
         Height          =   285
         Left            =   7680
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtPhone 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         TabIndex        =   16
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         TabIndex        =   14
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtCounty 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txtAddress3 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtAddress2 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtAddress1 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtCaseId 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   30
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   53
         _StockProps     =   15
         Caption         =   "SSPanel1"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   30
         Left            =   120
         TabIndex        =   44
         Top             =   3840
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   53
         _StockProps     =   15
         Caption         =   "SSPanel1"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTSampleTaken 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   250019843
         CurrentDate     =   40207
      End
      Begin MSMask.MaskEdBox txtSampleRecTime 
         Height          =   285
         Left            =   9120
         TabIndex        =   6
         Top             =   840
         Width           =   585
         _ExtentX        =   1032
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
         Left            =   3240
         TabIndex        =   4
         Top             =   840
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DTSampleRec 
         Height          =   285
         Left            =   7680
         TabIndex        =   5
         Top             =   840
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   250019843
         CurrentDate     =   40207
      End
      Begin VB.Frame fraButtons 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   120
         TabIndex        =   79
         Top             =   5640
         Width           =   10095
         Begin VB.CommandButton cmdCopyTo 
            Caption         =   "CC"
            Height          =   615
            Left            =   6360
            Picture         =   "frmDemographics.frx":037E
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtSpecimenLabelled 
            Height          =   855
            Left            =   6480
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   1440
            Width           =   3135
         End
         Begin VB.TextBox txtNatureSpecimen 
            Height          =   1095
            Left            =   6480
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   120
            Width           =   3135
         End
         Begin VB.CheckBox chkUrgent 
            Caption         =   "Urgent"
            Height          =   255
            Left            =   1440
            TabIndex        =   37
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox txtPatComments 
            Height          =   855
            Left            =   1440
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   1440
            Width           =   2895
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "&New"
            Height          =   615
            Left            =   7215
            Picture         =   "frmDemographics.frx":08B0
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   2640
            Width           =   735
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   615
            Left            =   8070
            Picture         =   "frmDemographics.frx":09B2
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   615
            Left            =   9165
            Picture         =   "frmDemographics.frx":0D79
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox txtClinicalHistory 
            Height          =   1095
            Left            =   1440
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   120
            Width           =   2895
         End
         Begin VB.CheckBox chkNoHistology 
            Caption         =   "No Histology Taken"
            Height          =   255
            Left            =   1440
            TabIndex        =   38
            Top             =   3000
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblCaseLocked 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   675
            Left            =   3540
            TabIndex        =   92
            Top             =   2640
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.Label lblLoggedIn 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   960
            TabIndex        =   91
            Top             =   3360
            Width           =   525
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Logged In : "
            Height          =   195
            Left            =   0
            TabIndex        =   90
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label lblCopyTo 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   6120
            TabIndex        =   88
            Top             =   2880
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Container Labelled"
            Height          =   195
            Left            =   5040
            TabIndex        =   84
            Top             =   1485
            Width           =   1320
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nature Of Specimen"
            Height          =   195
            Left            =   4935
            TabIndex        =   83
            Top             =   165
            Width           =   1440
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Patient Comments"
            Height          =   195
            Left            =   0
            TabIndex        =   81
            Top             =   1440
            Width           =   1375
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Clinical Details"
            Height          =   195
            Left            =   255
            TabIndex        =   80
            Top             =   120
            Width           =   1050
         End
      End
      Begin VB.Frame fraAutopsyDetails 
         BorderStyle     =   0  'None
         Caption         =   "Autopsy Details"
         Height          =   1455
         Left            =   6240
         TabIndex        =   68
         Top             =   4080
         Visible         =   0   'False
         Width           =   3735
         Begin VB.ComboBox cmbType 
            Height          =   315
            Left            =   1440
            TabIndex        =   32
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtCoronerClin 
            Height          =   285
            Left            =   1440
            TabIndex        =   29
            Top             =   0
            Width           =   2055
         End
         Begin VB.TextBox txtMothersName 
            Height          =   285
            Left            =   1440
            TabIndex        =   30
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtMothersDOB 
            Height          =   285
            Left            =   1455
            MaxLength       =   10
            TabIndex        =   31
            Tag             =   "Date of Birth"
            Text            =   "88/88/8888"
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label lblMothersName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mothers Name"
            Height          =   195
            Left            =   270
            TabIndex        =   72
            Top             =   420
            Width           =   1035
         End
         Begin VB.Label lblCoronerClin 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Coroner/Clinician"
            Height          =   195
            Left            =   90
            TabIndex        =   71
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblType 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Type"
            Height          =   195
            Left            =   960
            TabIndex        =   70
            Top             =   1125
            Width           =   360
         End
         Begin VB.Label lblMothersDOB 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DOB"
            Height          =   195
            Left            =   990
            TabIndex        =   69
            Top             =   765
            Width           =   345
         End
      End
      Begin VB.Frame fraGP 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1575
         Left            =   6840
         TabIndex        =   74
         Top             =   4080
         Visible         =   0   'False
         Width           =   3255
         Begin VB.TextBox txtGP 
            Height          =   285
            Left            =   840
            TabIndex        =   28
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label lblGP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GP"
            Height          =   195
            Left            =   500
            TabIndex        =   75
            Top             =   30
            Width           =   225
         End
      End
      Begin VB.Frame fraHospital 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1575
         Left            =   6300
         TabIndex        =   73
         Top             =   4080
         Visible         =   0   'False
         Width           =   3675
         Begin VB.TextBox txtWard 
            Height          =   285
            Left            =   1380
            TabIndex        =   24
            Top             =   0
            Width           =   2055
         End
         Begin VB.TextBox txtClinician 
            Height          =   285
            Left            =   1380
            TabIndex        =   26
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblWardGP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ward"
            Height          =   195
            Left            =   75
            TabIndex        =   77
            Top             =   60
            Width           =   1110
         End
         Begin VB.Label lblClinician 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Clinician"
            Height          =   195
            Left            =   150
            TabIndex        =   76
            Top             =   405
            Width           =   1110
         End
      End
      Begin VB.Label lblSampleTakenStar 
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
         Left            =   3000
         TabIndex        =   85
         Top             =   840
         Width           =   75
      End
      Begin VB.Label lblDateOfDeath 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date Of Death"
         Height          =   195
         Left            =   360
         TabIndex        =   78
         Top             =   4590
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblGpId 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   9000
         TabIndex        =   67
         Top             =   4080
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label23 
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
         Left            =   3960
         TabIndex        =   66
         Top             =   4080
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
         Left            =   9840
         TabIndex        =   65
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label Label19 
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
         Left            =   9840
         TabIndex        =   64
         Top             =   840
         Width           =   75
      End
      Begin VB.Label Label18 
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
         Left            =   3960
         TabIndex        =   63
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label17 
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
         Left            =   3960
         TabIndex        =   62
         Top             =   1920
         Width           =   75
      End
      Begin VB.Label Label16 
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
         Left            =   3960
         TabIndex        =   61
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label Label15 
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
         Left            =   3960
         TabIndex        =   58
         Top             =   3360
         Width           =   75
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sample Received"
         Height          =   195
         Left            =   6315
         TabIndex        =   56
         Top             =   870
         Width           =   1260
      End
      Begin VB.Label lblSampleTaken 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sample Taken"
         Height          =   195
         Left            =   255
         TabIndex        =   55
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Left            =   945
         TabIndex        =   45
         Top             =   4140
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Chart Number"
         Height          =   195
         Left            =   6630
         TabIndex        =   47
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Phone"
         Height          =   195
         Left            =   6720
         TabIndex        =   43
         Top             =   3405
         Width           =   825
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Gender"
         Height          =   195
         Left            =   7065
         TabIndex        =   54
         Top             =   2325
         Width           =   525
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Left            =   6720
         TabIndex        =   52
         Top             =   1965
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Left            =   6720
         TabIndex        =   50
         Top             =   1605
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "County"
         Height          =   195
         Left            =   945
         TabIndex        =   42
         Top             =   3405
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   270
         TabIndex        =   53
         Top             =   2325
         Width           =   1170
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Surname"
         Height          =   195
         Left            =   210
         TabIndex        =   51
         Top             =   1605
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Firstname"
         Height          =   195
         Left            =   285
         TabIndex        =   49
         Top             =   1965
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Case ID"
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   1170
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   9120
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDemographics.frx":10BB
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDemographics.frx":11CD
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDemographics.frx":12DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDemographics.frx":13F1
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   89
      Top             =   0
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label Label21 
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
      Left            =   8760
      TabIndex        =   60
      Top             =   1800
      Width           =   75
   End
   Begin VB.Label Label20 
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
      Left            =   3840
      TabIndex        =   59
      Top             =   1080
      Width           =   75
   End
End
Attribute VB_Name = "frmDemographics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pAddNew As Boolean
Private DemoChanged As Boolean
Private PrevCounty As String
Private YYYY As String
Private pLink As Boolean
Private LinkedCaseId As String



Public Property Let AddNew(ByVal Value As Boolean)

    pAddNew = Value

End Property
Public Property Let Link(ByVal Value As Boolean)

    pLink = Value

End Property

Private Sub cmdCopyTo_Click()
    With frmCopyTo
        .Move frmDemographics.Left + fraButtons.Left + cmdCopyTo.Left, frmDemographics.Top + fraButtons.Top + cmdCopyTo.Top - .Height
        .Show vbModal
    End With

    CheckCC

End Sub

Private Sub LoadCopyTo()
    Dim sql As String
    Dim tb As Recordset


    On Error GoTo LoadCopyTo_Error

    sql = "SELECT Consultant FROM SendCopyTo WHERE CaseId = N'" & CaseNo & "'"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql

    Do While Not tb.EOF
        If lblCopyTo <> "" Then
            lblCopyTo = lblCopyTo & vbCrLf & tb!Consultant
        Else
            lblCopyTo = tb!Consultant
        End If
        tb.MoveNext
    Loop
    CheckCC

    Exit Sub

LoadCopyTo_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "LoadCopyTo", intEL, strES, sql
End Sub

Private Sub SaveCopyTo(ByVal strCaseId As String)
    Dim sql As String
    Dim Y As Variant
    Dim strArray() As String


    On Error GoTo SaveCopyTo_Error

    sql = "Delete from SendCopyTo where " & _
          "CaseId = N'" & strCaseId & "'"
    Cnxn(0).Execute sql

    strArray = Split(lblCopyTo, vbCrLf)


    For Each Y In strArray

        sql = "Insert into SendCopyTo " & _
              "(CaseId, Consultant, Username) VALUES " & _
              "(N'" & strCaseId & "', " & _
            " N'" & AddTicks(Y) & "', " & _
            " N'" & AddTicks(UserName) & "')"

        Cnxn(0).Execute sql
    Next

    Exit Sub

SaveCopyTo_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "SaveCopyTo", intEL, strES, sql

End Sub


Private Sub CheckCC()
    Dim sql As String

    On Error GoTo CheckCC_Error

    cmdCopyTo.Caption = "CC"
    cmdCopyTo.Font.Bold = False
    cmdCopyTo.BackColor = &H8000000F


    If lblCopyTo <> "" Then
        cmdCopyTo.Caption = "CC"
        cmdCopyTo.Font.Bold = True
        cmdCopyTo.BackColor = &H8080FF
    End If

    Exit Sub

CheckCC_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmWorkSheet", "CheckCC", intEL, strES, sql


End Sub

Private Sub cmbSource_Click()

    If Left(txtCaseId, 2) <> "PA" _
       And Left(txtCaseId, 2) <> "MA" _
       And Left(txtCaseId, 2) <> "TA" Then
        If InStr(1, cmbSource.Text, "GP") = 0 Then
            fraHospital.Visible = True
            fraGP.Visible = False
            txtGP = ""
            lblGpId = ""

            fraAutopsyDetails.Visible = False
        ElseIf cmbSource.Text = "" Then
            fraHospital.Visible = False
            fraGP.Visible = False

            fraAutopsyDetails.Visible = False
        Else
            fraGP.Visible = True
            fraHospital.Visible = False
            txtClinician = ""
            txtWard = ""

            fraAutopsyDetails.Visible = False
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    ResetForm
End Sub

Private Sub cmdSave_Click()
    Dim strCaseId As String
    Dim blnContinueSave As Boolean
    
    If CheckValidation Then    'Check if mandatory fields are selected
        strCaseId = UCase(Replace(Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", ""), " ", ""))
        If VerifyCaseIdFormat(strCaseId) Then
            If pAddNew Or pLink Then    'a new case or a linked case / And Not for Edit Save
                blnContinueSave = SaveCaseDetails(strCaseId)
                If Not blnContinueSave Then
                  iMsg "Case Id not saved correct!"
                  ResetForm    'Clear form for new entry
                  DemoChanged = False
                  Exit Sub
                End If
                SaveCases strCaseId
            End If
            
            SaveDemographics strCaseId   'Save Demographics
            
            '******SAVE COPY TO HERE
            If lblCopyTo.Caption <> "" Then
                SaveCopyTo strCaseId
            End If
            
            ResetForm    'Clear form for new entry
            DemoChanged = False

            If pAddNew = False Then    'If editing log Edit event
                CaseAddLogEvent strCaseId, DemographicsEdited
                Unload Me
            Else    'Log add event
                CaseAddLogEvent strCaseId, DemographicsAdded
            End If

            If pLink Then
                frmWorkSheet.CheckLinkedCase strCaseId
            End If
            If pAddNew Then
                txtCaseId.SetFocus
            End If
        Else
            iMsg "Case Id incorrect!"
        End If
    End If

End Sub
Private Function CheckValidation() As Boolean


    If DTSampleTaken.CustomFormat = " " Or DTSampleRec.CustomFormat = " " _
       Or txtSampleRecTime = "__:__" _
       Or txtCounty = "" Or txtCaseId = "" _
       Or txtFirstName = "" Or txtSurname = "" _
       Or txtDOB = "" Or cmbSource.Text = "" Or txtSex = "" Then
        frmMsgBox.Msg "Please fill in mandatory fields", mbOKOnly, "Demographics", mbExclamation
        CheckValidation = False
        Exit Function
    Else
        CheckValidation = True
    End If

End Function

Private Sub cmdSearch_Click()
    With frmSearch
        .FromEdit = True
        .Show 1
    End With
End Sub


Private Sub Form_DblClick()

If IsIDE Then
    If sysOptCurrentLanguage = "Russian" Then
        sysOptCurrentLanguage = "English"
    Else
        sysOptCurrentLanguage = "Russian"
    End If
    
    LoadLanguage sysOptCurrentLanguage
'    frmDemographics_ChangeLanguage
End If

End Sub

Private Sub Form_Load()

    ChangeFont Me, "Arial"
    
'    frmDemographics_ChangeLanguage
    If pAddNew Then
        ResetForm
        FillLists
        lngMaxDigits = 11
    Else
        ResetForm
        FillLists
        LoadDemographics

        If cmbSource <> "" Then
            cmbSource_Click
        End If
        If pLink Then
            DisableForm
            txtCaseId.Enabled = True
            txtCaseId = ""

        Else
            YYYY = 2000 + Val(Right(txtCaseId, 2))
        End If
        If bLocked Then
            DisableForm
            cmdNew.Enabled = False
            cmdSave.Enabled = False
            lblCaseLocked.Visible = True
            lblCaseLocked = "RECORD BEING EDITED BY " & sCaseLockedBy
            lblCaseLocked.BackColor = &H8080FF
        End If
    End If
'    'ALI
'    txtPatComments.Text = "Patiet Comments"
'    'ALI-------
    lblLoggedIn = UserName
    If blnIsTestMode Then EnableTestMode Me
End Sub
Private Sub LoadDemographics()
    Dim sql As String
    Dim tb As Recordset
    Dim p As Integer


    On Error GoTo LoadDemographics_Error

    sql = "SELECT * FROM Demographics " & _
          "WHERE CaseID = N'" & CaseNo & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    If Not tb.EOF Then
        If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
            lngMaxDigits = 12
            txtCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
            lblDateOfDeath.Visible = True
            txtDateOfDeath.Visible = True
            fraAutopsy.Visible = True
            lblSampleTaken.Caption = "Sample" & " " & "Taken"  ' "Autopsy Date"
            lblSampleTakenStar.Visible = False
            fraButtons.Top = 5640
            fraDemographics.Height = 9255
            Me.Height = 9960
        Else
            lngMaxDigits = 11
            txtCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
            lblSampleTaken.Caption = "Sample" & " " & "Taken"
            lblSampleTakenStar.Visible = True
            fraButtons.Top = 4700
            fraDemographics.Height = 8335
            Me.Height = 9025
        End If
        If pLink Then
            LinkedCaseId = txtCaseId
        End If
        txtCaseId.Enabled = False
        txtChartNo = tb!MRN & ""
        txtFirstName = tb!FirstName & ""
        txtSurname = tb!Surname & ""
        txtAddress1 = tb!Address1 & ""
        txtAddress2 = tb!Address2 & ""
        txtAddress3 = tb!Address3 & ""
        txtCounty.MaxLength = 0
        txtCounty = tb!County & ""
        If Not IsNull(tb!DateOfBirth) Then
            txtDOB = Format(tb!DateOfBirth, "DD/MM/YYYY")
        Else
            txtDOB = ""
        End If
        txtAge = tb!Age & ""
'        Select Case tb!Sex & ""
'        Case "Male": txtSex = "M"
'        Case "Female": txtSex = "F"
'        Case Else: txtSex = ""
'        End Select
        txtSex = tb!Sex
        txtPhone = tb!Phone & ""
        cmbSource = tb!Source & ""
        txtClinician = tb!Clinician & ""
        txtWard = tb!Ward & ""
        txtGP = tb!GP & ""
        lblGpId = tb!GpId & ""
        txtPatComments = tb!Comments & ""
        chkUrgent = IIf(IsNull(tb!Urgent), 0, tb!Urgent)

        txtMothersName = tb!MothersName & ""
        txtMothersDOB = tb!MothersDOB & ""
        cmbType = tb!PaedType & ""
        For p = 0 To 2
            If optRequest(p).Caption = tb!AutopsyFor & "" Then
                optRequest(p).Value = True
            End If
        Next

        txtCoronerClin = tb!AutopsyRequestedBy & ""
        If Not IsNull(tb!DateOfDeath) Then
            txtDateOfDeath = Format(tb!DateOfDeath, "DD/MM/YYYY")
        Else
            txtDateOfDeath = ""
        End If
        txtMothersName = tb!MothersName & ""
        If Not IsNull(tb!MothersDOB) Then
            txtMothersDOB = Format(tb!MothersDOB, "DD/MM/YYYY")
        Else
            txtMothersDOB = ""
        End If
        txtClinicalHistory = tb!ClinicalHistory & ""
        txtNatureSpecimen = tb!NatureOfSpecimen & ""
        txtSpecimenLabelled = tb!SpecimenLabelled & ""

        cmbType = tb!PaedType & ""

        chkNoHistology = IIf(IsNull(tb!NoHistTaken), 0, tb!NoHistTaken)

        lstCoronerClin.Visible = False
        lstClinician.Visible = False
        lstWard.Visible = False
        lstGP.Visible = False
        If pAddNew = False Then
            DTSampleTaken.CustomFormat = "dd/MM/yyyy"
            DTSampleTaken = frmWorkSheet.DTSampleTaken
            DTSampleRec.CustomFormat = "dd/MM/yyyy"
            DTSampleRec = frmWorkSheet.DTSampleRec
            txtSampleTakenTime = frmWorkSheet.txtSampleTakenTime
            txtSampleRecTime = frmWorkSheet.txtSampleRecTime
            DTSampleTaken.Enabled = False
            DTSampleRec.Enabled = False
            txtSampleTakenTime.Enabled = False
            txtSampleRecTime.Enabled = False
        End If
        LoadCopyTo

    End If

    Exit Sub

LoadDemographics_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "LoadDemographics", intEL, strES, sql

End Sub

Private Sub FillLists()
    Dim sql As String
    Dim tb As Recordset

    On Error GoTo FillLists_Error

    cmbSource.AddItem ""
    sql = "SELECT * FROM Lists WHERE ListType = 'Source'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    Do While Not tb.EOF
        cmbSource.AddItem tb!Description & ""
        tb.MoveNext
    Loop
    cmbSource.ListIndex = -1
    cmbType.AddItem ""
    cmbType.AddItem "Stillborn"
    cmbType.AddItem "Miscarriage"
    cmbType.AddItem "Post Natal Death"
    cmbType.ListIndex = -1

    Exit Sub

FillLists_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "FillLists", intEL, strES, sql
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        Me.Top = 0
        Me.Left = Screen.Width / 2 - Me.Width / 2
    End If
End Sub


Private Sub ResetForm()

    txtCaseId = ""
    txtChartNo = ""
    txtFirstName = ""
    txtSurname = ""
    txtAddress1 = ""
    txtAddress2 = ""
    txtAddress3 = ""

    txtCounty = ""
    txtDOB = ""
    txtAge = ""
    txtSex = ""
    txtPhone = ""
    cmbSource = ""
    txtClinician = ""
    txtWard = ""
    txtGP = ""
    txtClinicalHistory = ""

    txtPatComments = ""
    chkUrgent.Value = 0
    DTSampleTaken.CustomFormat = " "
    DTSampleRec.CustomFormat = " "
    DTSampleTaken.Value = Date
    DTSampleRec.Value = Date
    txtSampleTakenTime = "__:__"
    txtSampleRecTime = "__:__"
    fraHospital.Visible = False
    fraGP.Visible = False
    fraAutopsyDetails.Visible = False
    fraAutopsy.Visible = False
    txtCoronerClin = ""
    txtMothersName = ""
    txtMothersDOB = ""
    cmbType = ""
    lblDateOfDeath.Visible = False

    txtDateOfDeath.Visible = False
    txtDateOfDeath = ""
    fraButtons.Top = 4700
    fraDemographics.Height = 8335
    Me.Height = 9025
    lblCoronerClin.Caption = "Coroner/Clinician"
    txtNatureSpecimen = ""
    txtSpecimenLabelled = ""
    lblCopyTo = ""

    CheckCC
End Sub
Private Sub DisableForm()
    txtChartNo.Enabled = False
    txtFirstName.Enabled = False
    txtSurname.Enabled = False
    txtAddress1.Enabled = False
    txtAddress2.Enabled = False
    txtAddress3.Enabled = False
    txtCounty.Enabled = False
    txtDOB.Enabled = False
    txtAge.Enabled = False
    txtSex.Enabled = False
    txtPhone.Enabled = False
    cmbSource.Enabled = False
    txtClinician.Enabled = False
    txtWard.Enabled = False
    txtGP.Enabled = False
    txtClinicalHistory.Enabled = False
    txtPatComments.Enabled = False
    DTSampleTaken.Enabled = False
    DTSampleRec.Enabled = False
    txtSampleTakenTime.Enabled = False
    txtSampleRecTime.Enabled = False
    fraHospital.Enabled = False
    fraGP.Enabled = False
    fraAutopsyDetails.Enabled = False
    fraAutopsy.Enabled = False
    txtCoronerClin.Enabled = False
    txtMothersName.Enabled = False
    txtMothersDOB.Enabled = False
    cmbType.Enabled = False
    lblDateOfDeath.Enabled = False
    txtDateOfDeath.Enabled = False
    txtDateOfDeath.Enabled = False
    txtNatureSpecimen.Enabled = False
    txtSpecimenLabelled.Enabled = False
    DemoChanged = False
End Sub

Private Sub SaveDemographics(ByVal strCaseId As String)

    Dim sql As String
    Dim tb As Recordset
    Dim p As Integer
    Dim gender As String


    On Error GoTo SaveDemographics_Error

    sql = "SELECT * FROM Demographics WHERE CaseID = N'" & strCaseId & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If tb.EOF Then tb.AddNew

    tb!CaseId = strCaseId
    tb!FirstName = txtFirstName
    tb!Surname = txtSurname
    tb!PatientName = txtFirstName & " " & txtSurname
'    ali
   Select Case Trim(CStr(txtSex.Text))
    Case "Female"
        gender = "F"
    Case "Male"
        gender = "M"
    Case Else
        gender = ""
End Select

tb!Sex = gender
'---------------------

    tb!Address1 = txtAddress1
    tb!Address2 = txtAddress2
    tb!Address3 = txtAddress3

    tb!County = txtCounty
    If Trim(txtDOB) <> "" Then
        tb!DateOfBirth = txtDOB
    End If
    tb!Age = CalcAge(txtDOB, DTSampleRec)
    tb!Ward = txtWard
    tb!Clinician = txtClinician
    tb!GP = txtGP
    tb!GpId = lblGpId
    tb!MRN = txtChartNo
    tb!Urgent = chkUrgent.Value
    tb!ClinicalHistory = txtClinicalHistory
    tb!Comments = txtPatComments
    tb!Source = cmbSource
    If Mid(strCaseId, 2, 1) = "A" Then
        For p = 0 To 2
            If optRequest(p).Value = True Then
                tb!AutopsyFor = optRequest(p).Caption
            End If
        Next
        tb!AutopsyRequestedBy = txtCoronerClin

        If Trim(txtDateOfDeath) <> "" Then
            tb!DateOfDeath = txtDateOfDeath
        End If
        tb!MothersName = txtMothersName
        If Trim(txtMothersDOB) <> "" Then
            tb!MothersDOB = txtMothersDOB
        End If
        tb!PaedType = cmbType
        tb!NoHistTaken = chkNoHistology.Value
    End If
    tb!NatureOfSpecimen = txtNatureSpecimen
    tb!SpecimenLabelled = txtSpecimenLabelled
    tb!Year = YYYY
    tb!UserName = UserName
    tb.Update

    Exit Sub

SaveDemographics_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "SaveDemographics", intEL, strES, sql

End Sub

Private Sub SaveCases(ByVal strCaseId As String)
    Dim sql As String
    Dim tb As New Recordset

    On Error GoTo SaveCases_Error

    sql = "SELECT * FROM Cases WHERE CaseID = N'" & strCaseId & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    If tb.EOF Then tb.AddNew

    tb!CaseId = strCaseId
    If txtSampleTakenTime <> "" Then
        tb!SampleTaken = Format(DTSampleTaken & " " & txtSampleTakenTime.FormattedText, "dd/MMM/yyyy hh:mm")
    Else
        tb!SampleTaken = Format(DTSampleTaken, "dd/MMM/yyyy")
    End If
    tb!SampleReceived = Format(DTSampleRec & " " & txtSampleRecTime.FormattedText, "dd/MMM/yyyy hh:mm")
    tb!State = "In Histology"
    If Left(strCaseId, 1) = "C" Then
        tb!Phase = "Cytology"
    Else
        tb!Phase = "Cut-Up"
    End If
    If pLink Then
        tb!LinkedCaseId = Replace(LinkedCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
    End If
    tb!UserName = UserName
    tb!Year = YYYY
    tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")
    tb.Update

    If pLink Then
        sql = "SELECT * FROM Cases WHERE CaseID = N'" & Replace(LinkedCaseId, " " & sysOptCaseIdSeperator(0) & " ", "") & "'"
        Set tb = New Recordset
        RecOpenServer 0, tb, sql

        If Not tb.EOF Then
            tb!LinkedCaseId = strCaseId
            tb.Update
        End If
    End If

    Exit Sub

SaveCases_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "SaveCases", intEL, strES, sql


End Sub

Private Function SaveCaseDetails(ByVal strCaseId As String) As Boolean
    Dim sql As String
    Dim UniqueId As String
    Dim tb As Recordset
    
    On Error GoTo SaveCaseDetails_Error

    SaveCaseDetails = False
    
    UniqueId = GetUniqueID 'Get unique ID

    'INSERT 1st branch into CaseTree
    sql = "Insert into CaseTree " & _
          "(CaseId, LocationID, LocationName, LocationParentID, LocationLevel, UserName) VALUES " & _
          "(N'" & strCaseId & "', " & _
        " '" & UniqueId & "', " & _
        " N'" & Trim(txtCaseId.Text) & "', " & _
        " '" & 0 & "', " & _
        " '" & 0 & "', " & _
        " N'" & AddTicks(UserName) & "')"
    Cnxn(0).Execute sql

    'Verify INSERT if NOT ok try again

    sql = "SELECT * from CaseTree where CaseId = N'" & strCaseId & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    If tb.EOF Then
        tb.AddNew
        tb!CaseId = strCaseId
        tb!LocationID = UniqueId
        tb!LocationName = txtCaseId
        tb!LocationParentID = 0
        tb!LocationLevel = 0
        tb!UserName = UserName
        tb.Update
    Else
        SaveCaseDetails = True
        Exit Function
    End If
    tb.Close
    
     
    sql = "SELECT * from CaseTree where CaseId = N'" & strCaseId & "'"
    Set tb = New Recordset
    Set tb = Cnxn(0).Execute(sql)

    If tb.EOF Then
        SaveCaseDetails = False
    Else
        SaveCaseDetails = True
    End If
    tb.Close
    
    Exit Function

SaveCaseDetails_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "SaveCaseDetails", intEL, strES, sql

End Function



Private Sub Form_Unload(Cancel As Integer)

    If pAddNew = True Then
        If DemoChanged = False Then
            frmWorklist.Enabled = True
            frmWorklist.tmrRefresh.Enabled = True

            frmWorklist.sCaseId = ""

            DataMode = 0
        Else
            If frmMsgBox.Msg("Alert!! Do you want to save your changes?", mbYesNo, , mbQuestion) = 1 = 1 Then
                cmdSave_Click
            End If
            frmWorklist.Enabled = True
            frmWorklist.tmrRefresh.Enabled = True

            frmWorklist.sCaseId = ""

            DataMode = 0
            DemoChanged = False
        End If
    Else
        frmWorkSheet.LoadDemographics (CaseNo)
    End If
End Sub



Private Sub lstGP_Click()
    lblGpId = lstGP.ItemData(lstGP.ListIndex)
End Sub

Private Sub optRequest_Click(Index As Integer)
    If UCase(optRequest(Index).Caption) = "CORONER" Then
        fraHospital.Visible = False
        fraGP.Visible = False
        fraAutopsyDetails.Visible = True
        lblMothersDOB.Visible = False
        txtMothersDOB.Visible = False
        lblType.Visible = False
        cmbType.Visible = False
        lblMothersName.Visible = False
        txtMothersName.Visible = False
        lblCoronerClin.Caption = "Coroners"
        txtCoronerClin = ""
        txtMothersName = ""
        txtMothersDOB = ""
        cmbType = ""
    ElseIf UCase(optRequest(Index).Caption) = "NON-CORONER" Then
        fraHospital.Visible = False
        fraGP.Visible = False
        fraAutopsyDetails.Visible = True
        lblMothersDOB.Visible = False
        txtMothersDOB.Visible = False
        lblType.Visible = False
        cmbType.Visible = False
        lblMothersName.Visible = False
        txtMothersName.Visible = False
        lblCoronerClin.Caption = "Clinician"
        txtCoronerClin = ""
        txtMothersName = ""
        txtMothersDOB = ""
        cmbType = ""
    Else
        fraHospital.Visible = False
        fraGP.Visible = False
        fraAutopsyDetails.Visible = True
        lblMothersDOB.Visible = True
        txtMothersDOB.Visible = True
        lblType.Visible = True
        cmbType.Visible = True
        lblMothersName.Visible = True
        txtMothersName.Visible = True
        lblCoronerClin.Caption = "Clinician"
        txtCoronerClin = ""
        txtMothersName = ""
        txtMothersDOB = ""
        cmbType = ""
    End If

End Sub



Private Sub optRequest_GotFocus(Index As Integer)
    If UCase(optRequest(Index).Caption) = "CORONER" Then
        fraHospital.Visible = False
        fraGP.Visible = False
        fraAutopsyDetails.Visible = True
        lblMothersDOB.Visible = False
        txtMothersDOB.Visible = False
        lblType.Visible = False
        cmbType.Visible = False
        lblMothersName.Visible = False
        txtMothersName.Visible = False
        lblCoronerClin.Caption = "Coroners"
        txtCoronerClin = ""
        txtMothersName = ""
        txtMothersDOB = ""
        cmbType = ""
    ElseIf UCase(optRequest(Index).Caption) = "NON-CORONER" Then
        fraHospital.Visible = False
        fraGP.Visible = False
        fraAutopsyDetails.Visible = True
        lblMothersDOB.Visible = False
        txtMothersDOB.Visible = False
        lblType.Visible = False
        cmbType.Visible = False
        lblMothersName.Visible = False
        txtMothersName.Visible = False
        lblCoronerClin.Caption = "Clinician"
        txtCoronerClin = ""
        txtMothersName = ""
        txtMothersDOB = ""
        cmbType = ""
    Else
        fraHospital.Visible = False
        fraGP.Visible = False
        fraAutopsyDetails.Visible = True
        lblMothersDOB.Visible = True
        txtMothersDOB.Visible = True
        lblType.Visible = True
        cmbType.Visible = True
        lblMothersName.Visible = True
        txtMothersName.Visible = True
        lblCoronerClin.Caption = "Clinician"
        txtCoronerClin = ""
        txtMothersName = ""
        txtMothersDOB = ""
        cmbType = ""
    End If

End Sub

Private Sub txtAddress1_LostFocus()
    txtAddress1 = initial2upper(txtAddress1)
End Sub



Private Sub txtAddress2_LostFocus()
    txtAddress2 = initial2upper(txtAddress2)
End Sub



Private Sub txtAddress3_LostFocus()
    txtAddress3 = initial2upper(txtAddress3)
End Sub

Private Sub txtCaseId_KeyPress(KeyAscii As Integer)
    Dim lngSel As Long, lngLen As Long

    On Error GoTo txtCaseId_KeyPress_Error
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'MsgBox UCase(sysOptCaseIdValidation(0))
    If UCase(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
        Call ValidateTullCaseId(KeyAscii, Me)
    Else
        Call ValidateLimCaseId(KeyAscii, Me)
    End If
   
    Exit Sub

txtCaseId_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "txtCaseId_KeyPress", intEL, strES


End Sub

Private Sub txtCaseId_LostFocus()
    Dim sql As String
    Dim tb As Recordset
    Dim strCaseId As String

    On Error GoTo txtCaseId_LostFocus_Error
    
    'txtCaseId = initial2upper(txtCaseId)
    txtCaseId.Text = UCase(txtCaseId.Text)
    If IsValidCaseNo(txtCaseId) Then
        If pAddNew Or pLink Then
            strCaseId = UCase(Replace(Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", ""), " ", ""))

            sql = "SELECT * FROM Cases WHERE CaseId = N'" & strCaseId & "'"
            Set tb = New Recordset
            RecOpenClient 0, tb, sql
            If Not tb.EOF Then
                frmMsgBox.Msg "Case ID already Exists"
                If pAddNew Then
                    ResetForm
                Else
                    txtCaseId = ""
                End If
                Exit Sub
            End If
        End If
        If Left(txtCaseId, 2) = "PA" _
           Or Left(txtCaseId, 2) = "MA" _
           Or Left(txtCaseId, 2) = "TA" Then

            fraAutopsy.Visible = True
            fraHospital.Visible = False
            fraGP.Visible = False
            If optRequest(0) <> True And _
               optRequest(1) <> True And _
               optRequest(2) <> True Then
                fraAutopsyDetails.Visible = False
            End If
            lblSampleTaken.Caption = "Sample Taken" ' "Autopsy Date"
            lblSampleTakenStar.Visible = False
            lblDateOfDeath.Visible = True
            txtDateOfDeath.Visible = True
            fraButtons.Top = 5640
            fraDemographics.Height = 9255
            Me.Height = 9960
            chkNoHistology.Visible = True

        Else

            fraAutopsy.Visible = False
            fraAutopsyDetails.Visible = False
            lblSampleTaken.Caption = "Sample Taken" ' "Sample Taken"
            lblDateOfDeath.Visible = False
            txtDateOfDeath.Visible = False
            chkNoHistology.Visible = False

        End If
        YYYY = 2000 + Val(Right(txtCaseId, 2))
        DemoChanged = True
        Exit Sub
    Else

        txtCaseId = ""
        txtCoronerClin = ""
        txtMothersDOB = ""
        txtMothersName = ""
        cmbType = ""
        fraAutopsy.Visible = False
        fraAutopsyDetails.Visible = False
    End If

    Exit Sub

txtCaseId_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "txtCaseId_LostFocus", intEL, strES, sql

End Sub

Private Sub txtChartNo_LostFocus()
    On Error GoTo ErrorHandler
1    txtChartNo = UCase(txtChartNo)

2    If txtChartNo <> "" And pAddNew Then
3        LoadDemo txtChartNo
4    End If
Exit Sub
ErrorHandler:

MsgBox (Err.Description & " " & Erl)

End Sub

Private Sub LoadDemo(ByVal IDNumber As String)
    Dim tb As New Recordset
    Dim sn As New Recordset
    Dim sql As String
    Dim f As Form

    On Error GoTo LoadDemo_Error

    sql = "SELECT COUNT(*) AS RecordNumber FROM Demographics WHERE MRN = '" & IDNumber & "'"

    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb.EOF Then
        If tb!RecordNumber > 1 Then
            If txtFirstName = "" _
               And txtSurname = "" _
               And txtCounty = "" _
               And txtDOB = "" Then

                Set f = New frmDemoConflict
                With f
                    .IDNumber = IDNumber
                    .lblNumber = "MRN: " & IDNumber
                    .Show 1
                End With
                Unload f
                Set f = Nothing
            End If
        Else
            sql = "SELECT TOP 1 * FROM Demographics " & _
                  "WHERE MRN = N'" & IDNumber & "' " & _
                  "ORDER BY DateTimeOfRecord DESC"

            Set sn = New Recordset
            RecOpenServer 0, sn, sql

            If Not sn.EOF Then
                txtFirstName = sn!FirstName & ""
                txtSurname = sn!Surname & ""
                txtAddress1 = sn!Address1 & ""
                txtAddress2 = sn!Address2 & ""
                txtAddress3 = sn!Address3 & ""
                txtCounty.MaxLength = 0
                txtCounty = sn!County & ""
                If Not IsNull(sn!DateOfBirth) Then
                    txtDOB = Format(sn!DateOfBirth, "DD/MM/YYYY")
                Else
                    txtDOB = ""
                End If
                If DTSampleRec.CustomFormat <> " " Then
                    txtAge = CalcAge(txtDOB, DTSampleRec)
                Else
                    txtAge = CalcAge(txtDOB, Now)
                End If
                'ali
                Select Case Trim(CStr(sn!Sex))
                Case "F": txtSex = "Female"
                Case "M": txtSex = "Male"
                Case Else: txtSex = ""
                End Select
                '-----
'                Select Case sn!Sex & ""
'                Case "F": txtSex = "Female"
'                Case "M": txtSex = "Male"
'                Case Else: txtSex = ""
'                End Select
                txtPhone = sn!Phone & ""
                txtPatComments = sn!Comments & ""
            End If
            sn.Close
        End If


    End If
    tb.Close



    Exit Sub

LoadDemo_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "LoadDemo", intEL, strES, sql


End Sub
Private Sub DTSampleTaken_GotFocus()
    If DTSampleTaken.CustomFormat = " " Then
        If Not IsIDE Then
            SendKeys "{RIGHT}"
        End If
    End If
    DTSampleTaken.CustomFormat = "dd/MM/yyyy"

End Sub

Private Sub DTSampleTaken_Validate(Cancel As Boolean)
    If DTSampleTaken > DTSampleRec Then
        'Comment
        'frmMsgBox.Msg "Sample Date is after Received Date, Please Amend", , , mbExclamation
        MsgBox ("Sample Date is after Received Date. Please Change")
        Cancel = True
    ElseIf DTSampleRec = DTSampleTaken Then
        If txtSampleTakenTime <> "" Then
            If Format(txtSampleTakenTime.FormattedText, "HH:mm") > Format(txtSampleRecTime.FormattedText, "HH:mm") Then
                'frmMsgBox.Msg "Sample Date is after Received Date, Please Amend", , ,
                MsgBox ("Sample Date is after Received Date. Please Change")
                Cancel = True
            End If
        End If
    End If
End Sub






Private Sub txtClinician_Validate(Cancel As Boolean)
    txtClinician = QueryKnown("Clinician", txtClinician, cmbSource)
    lstClinician.Visible = False
End Sub

Private Sub txtCoronerClin_Validate(Cancel As Boolean)
    txtCoronerClin = QueryKnown(lblCoronerClin, txtCoronerClin, cmbSource)
    lstCoronerClin.Visible = False
End Sub

Private Sub txtCounty_LostFocus()
    FillCountyNames
End Sub



Private Sub txtDateOfDeath_KeyPress(KeyAscii As Integer)
    KeyAscii = VI(KeyAscii, NumericSlash)
End Sub

Private Sub txtDateOfDeath_LostFocus()
    txtDateOfDeath = Convert62Date(txtDateOfDeath, BACKWARD)

    If Len(txtDateOfDeath) = 8 And Not IsDate(txtDateOfDeath) Then
        txtDateOfDeath = Left(txtDateOfDeath, 2) & "/" & Mid(txtDateOfDeath, 3, 2) & "/" & Right(txtDateOfDeath, 4)
    End If
    If Not IsDate(txtDateOfDeath) Then
        txtDateOfDeath = ""
        Exit Sub
    End If

    If Format$(txtDateOfDeath, "yyyymmdd") > Format$(Now, "yyyymmdd") Then
        txtDateOfDeath = ""
        Exit Sub
    End If
End Sub

Private Sub txtDOB_KeyPress(KeyAscii As Integer)
    KeyAscii = VI(KeyAscii, NumericSlash)
End Sub

Private Sub txtDOB_LostFocus()
    txtDOB = Convert62Date(txtDOB, BACKWARD)

    If Len(txtDOB) = 8 And Not IsDate(txtDOB) Then
        txtDOB = Left(txtDOB, 2) & "/" & Mid(txtDOB, 3, 2) & "/" & Right(txtDOB, 4)
    End If
    If Not IsDate(txtDOB) Then
        txtDOB = ""
        Exit Sub
    End If

    If Format$(txtDOB, "yyyymmdd") > Format$(Now, "yyyymmdd") Then
        txtDOB = ""
        Exit Sub
    End If

    If DTSampleRec.CustomFormat <> " " Then
        txtAge = CalcAge(txtDOB, DTSampleRec)
    Else
        txtAge = CalcAge(txtDOB, Now)
    End If
End Sub


Private Sub txtFirstName_LostFocus()
    Dim strName As String
    Dim strSex As String

    On Error GoTo txtFirstName_LostFocus_Error

    strName = txtFirstName
    strSex = txtSex

    NameLostFocus strName, strSex

    txtFirstName = initial2upper(txtFirstName)
    txtSex = strSex

    Exit Sub

txtFirstName_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "txtFirstName_LostFocus", intEL, strES


End Sub


Private Sub txtGP_Validate(Cancel As Boolean)
    Dim sTempSource As String

    sTempSource = Replace(cmbSource, "GP", "")
    sTempSource = Replace(sTempSource, "-", "")
    txtGP = QueryKnown("GP", txtGP, Trim(sTempSource))
    lstGP.Visible = False

End Sub

Private Sub txtMothersDOB_KeyPress(KeyAscii As Integer)
    KeyAscii = VI(KeyAscii, NumericSlash)
End Sub

Private Sub txtMothersDOB_LostFocus()
    On Error GoTo txtMothersDOB_LostFocus_Error

    txtMothersDOB = Convert62Date(txtMothersDOB, BACKWARD)

    If Len(txtMothersDOB) = 8 And Not IsDate(txtMothersDOB) Then
        txtMothersDOB = Left(txtMothersDOB, 2) & "/" & Mid(txtMothersDOB, 3, 2) & "/" & Right(txtMothersDOB, 4)
    End If
    If Not IsDate(txtMothersDOB) Then
        txtMothersDOB = ""
        Exit Sub
    End If

    If Format$(txtMothersDOB, "yyyymmdd") > Format$(Now, "yyyymmdd") Then
        txtMothersDOB = ""
        Exit Sub
    End If

    Exit Sub

txtMothersDOB_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "txtMothersDOB_LostFocus", intEL, strES

End Sub



Private Sub txtMothersName_LostFocus()
    txtMothersName = initial2upper(txtMothersName)
End Sub

Private Sub txtSampleTakenTime_GotFocus()
    txtSampleTakenTime.SelStart = 0
    txtSampleTakenTime.SelLength = Len(txtSampleTakenTime.FormattedText)
End Sub



Private Sub txtSampleTakenTime_Validate(Cancel As Boolean)

    If DTSampleRec = DTSampleTaken Then
        If txtSampleTakenTime <> "" Then
            If Format(txtSampleTakenTime.FormattedText, "HH:mm") > Format(txtSampleRecTime.FormattedText, "HH:mm") Then
                'frmMsgBox.Msg LS(csSampleDateAfterRecDatePleaseAmend), , , mbExclamation
                MsgBox "Sample Date is after Received Date, Please Change", vbExclamation
                Cancel = True
            End If
        End If
    End If

End Sub

Private Sub DTSampleRec_GotFocus()
    If DTSampleRec.CustomFormat = " " Then
        If Not IsIDE Then
            SendKeys "{RIGHT}"
        End If
    End If
    DTSampleRec.CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub DTSampleRec_Validate(Cancel As Boolean)
    If DTSampleRec < DTSampleTaken Then
        frmMsgBox.Msg "Sample Received Date cannot be before Sample Taken Date", , , mbExclamation
        'MsgBox , vbExclamation
        Cancel = True
    ElseIf DTSampleRec = DTSampleTaken Then
        If txtSampleTakenTime <> "" Then
            If Format(txtSampleRecTime.FormattedText, "HH:mm") < Format(txtSampleTakenTime.FormattedText, "HH:mm") Then
                frmMsgBox.Msg "Sample Received Date cannot be before Sample Taken Date", , , mbExclamation
                'MsgBox "Sample Date is after Received Date, Please Change", vbExclamation
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub txtSampleRecTime_GotFocus()
    txtSampleRecTime.SelStart = 0
    txtSampleRecTime.SelLength = Len(txtSampleRecTime.FormattedText)
End Sub



Private Sub txtSampleRecTime_Validate(Cancel As Boolean)

    If DTSampleRec = DTSampleTaken Then
        If txtSampleTakenTime <> "" Then
            If Format(txtSampleRecTime.FormattedText, "HH:mm") < Format(txtSampleTakenTime.FormattedText, "HH:mm") Then
                'frmMsgBox.Msg LS(csSampleDateAfterRecDatePleaseAmend), , , mbExclamation
                MsgBox "Sample Date is after Received Date, Please Change", vbExclamation
                Cancel = True
            End If
        End If
    End If

End Sub

Private Sub FillClinicianNames()

    Dim tb As Recordset
    Dim sql As String

    On Error GoTo FillClinicianNames_Error

    sql = "SELECT TOP 50 Description FROM SourceItemLists WHERE " & _
          "ListType = 'Clinician' " & _
          "AND Description LIKE N'%" & AddTicks(txtClinician) & "%' " & _
          "AND (Source = N'" & ListCodeFor("Source", cmbSource) & "' " & _
          "OR Source = '' OR Source IS NULL) " & _
          "AND Inuse = 1 ORDER BY Description"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
    lstClinician.Clear
    Do While Not tb.EOF
        lstClinician.AddItem tb!Description & ""
        tb.MoveNext
    Loop

    Exit Sub

FillClinicianNames_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "FillClinicianNames", intEL, strES, sql


End Sub

Private Sub FillCoronerClinNames()

    Dim tb As Recordset
    Dim sql As String



    On Error GoTo FillCoronerNames_Error

    sql = "SELECT TOP 50 Description FROM SourceItemLists WHERE " & _
          "ListType = N'" & "Coroner" & "' " & _
          "AND Description LIKE N'%" & AddTicks(txtCoronerClin) & "%' " & _
          "AND (Source = N'" & ListCodeFor("Source", cmbSource) & "' " & _
          "OR Source = '' OR Source IS NULL) " & _
          "AND Inuse = 1 ORDER BY Description"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
    lstCoronerClin.Clear
    Do While Not tb.EOF
        lstCoronerClin.AddItem tb!Description & ""
        tb.MoveNext
    Loop




    Exit Sub

FillCoronerNames_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "FillCoronerNames", intEL, strES, sql


End Sub

Private Sub txtClinician_GotFocus()
    lstClinician.Left = fraHospital.Left + txtClinician.Left
    lstClinician.Top = fraHospital.Top + txtClinician.Top + txtClinician.Height
End Sub


Private Sub lstClinician_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtClinician = lstClinician
        txtClinician.SetFocus
        lstClinician.Visible = False
    End If

    If KeyAscii = 8 Then    'backspace
        txtClinician.SetFocus
        If Len(txtClinician) > 0 Then
            txtClinician = Left$(txtClinician, Len(txtClinician) - 1)
        End If
        txtClinician.SelStart = 9999
    End If

End Sub

Private Sub lstClinician_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtClinician = lstClinician
    txtClinician.SetFocus
    lstClinician.Visible = False

End Sub

Private Sub txtClinician_Change()


    If Trim$(txtClinician) = "" Then
        lstClinician.Visible = False
        Exit Sub
    End If

    lstClinician.Visible = True
    FillClinicianNames

End Sub

Private Sub txtClinician_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        lstClinician.Clear
        lstClinician.Visible = False
        KeyCode = 0
        Exit Sub
    End If
    If KeyCode = vbKeyDown Then
        If lstClinician.ListCount > 0 Then
            If lstClinician.Visible = True Then
                lstClinician.SetFocus
                lstClinician.Selected(0) = True
            End If
        End If
    End If

End Sub
Private Sub lstCoronerClin_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtCoronerClin = lstCoronerClin
        txtCoronerClin.SetFocus
        lstCoronerClin.Visible = False
    End If

    If KeyAscii = 8 Then    'backspace
        txtCoronerClin.SetFocus
        If Len(txtCoronerClin) > 0 Then
            txtCoronerClin = Left$(txtCoronerClin, Len(txtCoronerClin) - 1)
        End If
        txtCoronerClin.SelStart = 9999
    End If

End Sub

Private Sub lstCoronerClin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtCoronerClin = lstCoronerClin
    txtCoronerClin.SetFocus
    lstCoronerClin.Visible = False

End Sub
Private Sub txtCoronerClin_Change()


    If Trim$(txtCoronerClin) = "" Then
        lstCoronerClin.Visible = False
        Exit Sub
    End If

    lstCoronerClin.Visible = True



    FillCoronerClinNames

End Sub
Private Sub txtCoronerClin_GotFocus()
    lstCoronerClin.Left = fraAutopsyDetails.Left + txtCoronerClin.Left
    lstCoronerClin.Top = fraAutopsyDetails.Top + txtCoronerClin.Top + txtCoronerClin.Height
End Sub

Private Sub txtCoronerClin_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        lstCoronerClin.Clear
        lstCoronerClin.Visible = False
        KeyCode = 0
        Exit Sub
    End If
    If KeyCode = vbKeyDown Then
        If lstCoronerClin.ListCount > 0 Then
            If lstCoronerClin.Visible = True Then
                lstCoronerClin.SetFocus
                lstCoronerClin.Selected(0) = True
            End If
        End If
    End If

End Sub

Private Sub FillCountyNames()

    Dim tb As Recordset
    Dim sql As String

    On Error GoTo FillCountyNames_Error

    sql = "SELECT Description FROM Lists WHERE " & _
          "ListType = 'County' " & _
          "AND Code = N'" & AddTicks(txtCounty) & "' " & _
          "AND Inuse = 1 ORDER BY Description"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql

    If Not tb.EOF Then
        txtCounty.MaxLength = 0
        txtCounty = tb!Description & ""
    Else
        If PrevCounty <> "" Then
            txtCounty.MaxLength = 0
        End If
        txtCounty = PrevCounty
    End If

    Exit Sub

FillCountyNames_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "FillCountyNames", intEL, strES, sql


End Sub

Private Sub txtCounty_GotFocus()
    PrevCounty = txtCounty
    txtCounty.MaxLength = 2
End Sub


Private Sub FillGPNames()

    Dim tb As Recordset
    Dim sql As String
    Dim sCounty As String
    Dim sTempSource As String

    On Error GoTo FillGPNames_Error

    sTempSource = Replace(cmbSource, "GP", "")
    sTempSource = Replace(sTempSource, "-", "")

    sCounty = Trim$(sTempSource)


    sql = "SELECT distinct TOP 10 (GPName),GPid FROM GPs WHERE " & _
          "(GPName LIKE N'" & AddTicks(txtGP) & "%' " & _
          "OR SurName LIKE N'" & AddTicks(txtGP) & "%' " & _
          "OR FirstName LIKE N'" & AddTicks(txtGP) & "%') " & _
          "AND County = N'" & sCounty & "' " & _
          "AND Inuse = 1 ORDER BY GPName"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
    lstGP.Clear
    Do While Not tb.EOF
        lstGP.AddItem tb!GPName & ""
        lstGP.ItemData(lstGP.NewIndex) = tb!GpId & ""
        tb.MoveNext
    Loop

    Exit Sub

FillGPNames_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "FillGPNames", intEL, strES, sql


End Sub

Private Sub txtGP_GotFocus()
    lstGP.Left = fraGP.Left + txtGP.Left
    lstGP.Top = fraGP.Top + txtGP.Top + txtGP.Height
End Sub


Private Sub lstGP_KeyPress(KeyAscii As Integer)
    Dim sql As String
    Dim tb As New Recordset
    Dim f As Form

    On Error GoTo lstGP_KeyPress_Error

    If KeyAscii = 13 Then
        sql = "SELECT COUNT(*) AS RecordNumber FROM GPs WHERE GPName = N'" & AddTicks(lstGP) & "'"
        Set tb = New Recordset
        RecOpenClient 0, tb, sql
        If tb!RecordNumber > 1 Then
            Set f = New frmGPConflict
            With f
                .GPName = lstGP
                .lblGP = lstGP
                .Show 1
            End With
            Unload f
            Set f = Nothing
        Else
            txtGP = lstGP
        End If
        txtGP.SetFocus
        lstGP.Visible = False
    End If

    If KeyAscii = 8 Then    'backspace
        txtGP.SetFocus
        If Len(txtGP) > 0 Then
            txtGP = Left$(txtGP, Len(txtGP) - 1)
        End If
        txtGP.SelStart = 9999
    End If

    Exit Sub

lstGP_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "lstGP_KeyPress", intEL, strES, sql


End Sub

Private Sub lstGP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim sql As String
    Dim tb As New Recordset
    Dim f As Form

    sql = "SELECT COUNT(*) AS RecordNumber FROM GPs WHERE GPName = N'" & AddTicks(lstGP) & "'"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
    If tb!RecordNumber > 1 Then
        Set f = New frmGPConflict
        With f
            .GPName = lstGP
            .lblGP = lstGP
            .Show 1
        End With
        Unload f
        Set f = Nothing
    Else
        txtGP = lstGP
    End If
    txtGP.SetFocus
    lstGP.Visible = False

End Sub

Private Sub txtGP_Change()


    If Trim$(txtGP) = "" Then
        lblGpId = ""
        lstGP.Visible = False
        Exit Sub
    End If

    lstGP.Visible = True
    FillGPNames

End Sub

Private Sub txtGP_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        lstGP.Clear
        lstGP.Visible = False
        KeyCode = 0
        Exit Sub
    End If
    If KeyCode = vbKeyDown Then
        If lstGP.ListCount > 0 Then
            If lstGP.Visible = True Then
                lstGP.SetFocus
                lstGP.Selected(0) = True
            End If
        End If
    End If

End Sub


Private Sub FillWardNames()

    Dim tb As Recordset
    Dim sql As String

    On Error GoTo FillWardNames_Error

    sql = "SELECT TOP 50 Description FROM SourceItemLists WHERE " & _
          "ListType = 'Ward' " & _
          "AND Description LIKE N'%" & AddTicks(txtWard) & "%' " & _
          "AND (Source = N'" & ListCodeFor("Source", cmbSource) & "' " & _
          "OR Source = '' OR Source IS NULL) " & _
          "AND Inuse = 1 ORDER BY Description"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
    lstWard.Clear
    Do While Not tb.EOF
        lstWard.AddItem tb!Description & ""
        tb.MoveNext
    Loop

    Exit Sub

FillWardNames_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmDemographics", "FillWardNames", intEL, strES, sql


End Sub

Private Sub txtSex_Click()
    'Zyam 26-2-24
    Select Case Trim$(txtSex)
    Case "": txtSex = "Male" 'ls(csMale)
    Case "Male": txtSex = "Female" 'ls(csFemale)
    Case "Female": txtSex = ""
    Case Else: txtSex = ""
    End Select
    'Zyam 26-2-24
End Sub

Private Sub txtSex_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    txtSex_Click
End Sub

Private Sub txtSex_LostFocus()
    SexLostFocus txtSex, txtFirstName
End Sub
Private Sub txtSurname_LostFocus()
    txtSurname = initial2upper(txtSurname)
End Sub

Private Sub txtWard_GotFocus()
    lstWard.Left = fraHospital.Left + txtWard.Left
    lstWard.Top = fraHospital.Top + txtWard.Top + txtWard.Height
End Sub


Private Sub lstWard_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtWard = lstWard
        txtWard.SetFocus
        lstWard.Visible = False
    End If

    If KeyAscii = 8 Then    'backspace
        txtWard.SetFocus
        If Len(txtWard) > 0 Then
            txtWard = Left$(txtWard, Len(txtWard) - 1)
        End If
        txtWard.SelStart = 9999
    End If

End Sub

Private Sub lstWard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtWard = lstWard
    txtWard.SetFocus
    lstWard.Visible = False

End Sub

Private Sub txtWard_Change()


    If Trim$(txtWard) = "" Then
        lstWard.Visible = False
        Exit Sub
    End If

    lstWard.Visible = True
    FillWardNames

End Sub

Private Sub txtWard_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        lstWard.Clear
        lstWard.Visible = False
        KeyCode = 0
        Exit Sub
    End If
    If KeyCode = vbKeyDown Then
        If lstWard.ListCount > 0 Then
            If lstWard.Visible = True Then
                lstWard.SetFocus

                lstWard.Selected(0) = True
            End If
        End If
    End If

End Sub


Private Sub txtWard_Validate(Cancel As Boolean)
    txtWard = QueryKnown("Ward", txtWard, cmbSource)
    lstWard.Visible = False
End Sub
