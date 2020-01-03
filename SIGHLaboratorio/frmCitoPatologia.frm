VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form frmCitoPatologia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CITOPATOLOGÍA"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "frmCitoPatologia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame CPA001 
      Caption         =   "Citología Cérvico-Vaginal (Papanicolau)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6495
      Left            =   60
      TabIndex        =   61
      Top             =   2400
      Visible         =   0   'False
      Width           =   7155
      Begin TabDlg.SSTab CPA001_15 
         Height          =   6135
         Left            =   60
         TabIndex        =   113
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Toma de Muestra"
         TabPicture(0)   =   "frmCitoPatologia.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "CPA001_01(18)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "CPA001_01(14)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "CPA001_00(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CPA001_00(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CPA001_00(2)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "CPA001_00(3)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "CPA001_26"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Resultados"
         TabPicture(1)   =   "frmCitoPatologia.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CPA001_00(8)"
         Tab(1).Control(1)=   "CPA001_00(6)"
         Tab(1).Control(2)=   "CPA001_00(5)"
         Tab(1).Control(3)=   "CPA001_00(4)"
         Tab(1).Control(4)=   "CPA001_00(7)"
         Tab(1).Control(5)=   "CPA001_01(15)"
         Tab(1).ControlCount=   6
         Begin VB.TextBox CPA001_26 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1665
            TabIndex        =   0
            Top             =   480
            Width           =   5115
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "5.- Recomendaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   700
            Index           =   8
            Left            =   -74880
            TabIndex        =   84
            Top             =   5310
            Width           =   6735
            Begin VB.TextBox CPA001_25 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   60
               MultiLine       =   -1  'True
               TabIndex        =   60
               Top             =   210
               Width           =   6555
            End
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "4.- Responsable de toma de PAP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Index           =   3
            Left            =   120
            TabIndex        =   82
            Top             =   4440
            Width           =   6735
            Begin VB.TextBox CPA001_20 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1485
               TabIndex        =   21
               Top             =   210
               Width           =   5115
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre y Apellidos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   83
               Top             =   240
               Width           =   1365
            End
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "3.- Observaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1215
            Index           =   2
            Left            =   120
            TabIndex        =   80
            Top             =   3090
            Width           =   6735
            Begin VB.TextBox CPA001_19 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1005
               MultiLine       =   -1  'True
               TabIndex        =   20
               Top             =   720
               Width           =   5595
            End
            Begin VB.CheckBox CPA001_05 
               Caption         =   "Anormal"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   3
               Left            =   5400
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox CPA001_05 
               Caption         =   "Erosión"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   3840
               TabIndex        =   14
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox CPA001_05 
               Caption         =   "Aparentemente sano"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox CPA001_05 
               Caption         =   "Cervicitis"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   2280
               TabIndex        =   13
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox CPA001_05 
               Caption         =   "Prolapso"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   5400
               TabIndex        =   19
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox CPA001_05 
               Caption         =   "Leucorrea"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   3840
               TabIndex        =   18
               Top             =   480
               Width           =   1215
            End
            Begin VB.CheckBox CPA001_05 
               Caption         =   "Sangrado anormal"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox CPA001_05 
               Caption         =   "Prolapso"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   2280
               TabIndex        =   17
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Comentarios"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   81
               Top             =   750
               Width           =   900
            End
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "2.- Situación Gineco-Obstetrica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1095
            Index           =   1
            Left            =   120
            TabIndex        =   77
            Top             =   1920
            Width           =   6735
            Begin Threed.SSOption CPA001_04 
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   6
               Top             =   600
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Premenopausia"
            End
            Begin VB.TextBox CPA001_17 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   375
               TabIndex        =   4
               Top             =   210
               Width           =   2595
            End
            Begin Threed.SSOption CPA001_04 
               Height          =   195
               Index           =   1
               Left            =   2385
               TabIndex        =   7
               Top             =   600
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Postmenopausia"
            End
            Begin Threed.SSOption CPA001_04 
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   9
               Top             =   840
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "DIU"
            End
            Begin Threed.SSOption CPA001_04 
               Height          =   195
               Index           =   2
               Left            =   5040
               TabIndex        =   8
               Top             =   600
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Uso de Hormonas"
            End
            Begin Threed.SSOption CPA001_04 
               Height          =   195
               Index           =   5
               Left            =   5040
               TabIndex        =   11
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Embarazo actual"
            End
            Begin Threed.SSOption CPA001_04 
               Height          =   195
               Index           =   4
               Left            =   2385
               TabIndex        =   10
               Top             =   840
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Histerectomía"
            End
            Begin VB.TextBox CPA001_18 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4920
               TabIndex        =   5
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "GP"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   195
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha última regla"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   3570
               TabIndex        =   78
               Top             =   270
               Width           =   1305
            End
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "1.- Citología"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   0
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   6735
            Begin Threed.SSOption CPA001_03 
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   1
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Primera vez"
            End
            Begin Threed.SSOption CPA001_03 
               Height          =   195
               Index           =   1
               Left            =   1530
               TabIndex        =   2
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Subsecuente"
            End
            Begin VB.TextBox CPA001_16 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   3375
               MultiLine       =   -1  'True
               TabIndex        =   3
               Top             =   450
               Visible         =   0   'False
               Width           =   3195
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Diagnóstico Último PAP"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   1695
               TabIndex        =   76
               Top             =   480
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "3.- Hallazgos adicionales"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1030
            Index           =   6
            Left            =   -74880
            TabIndex        =   71
            Top             =   2950
            Width           =   6735
            Begin VB.CheckBox CPA001_11 
               Caption         =   "Sospechoso de Virus Herpes"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   4200
               TabIndex        =   44
               Top             =   240
               Width           =   2415
            End
            Begin VB.CheckBox CPA001_11 
               Caption         =   "Cambios reactivos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox CPA001_11 
               Caption         =   "Trichomonas"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   4200
               TabIndex        =   47
               Top             =   480
               Width           =   1215
            End
            Begin VB.CheckBox CPA001_11 
               Caption         =   "Sospechoso de PVH"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   2040
               TabIndex        =   43
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox CPA001_11 
               Caption         =   "Atrofia"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   45
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox CPA001_11 
               Caption         =   "Bacterias"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   4
               Left            =   2040
               TabIndex        =   46
               Top             =   480
               Width           =   1215
            End
            Begin VB.CheckBox CPA001_11 
               Caption         =   "Hongos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   48
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox CPA001_24 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2805
               TabIndex        =   50
               Top             =   690
               Visible         =   0   'False
               Width           =   3795
            End
            Begin VB.CheckBox CPA001_11 
               Caption         =   "Otros"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   2040
               TabIndex        =   49
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "2.- Evaluación citológica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1095
            Index           =   5
            Left            =   -74880
            TabIndex        =   67
            Top             =   1780
            Width           =   6735
            Begin Threed.SSOption CPA001_09 
               Height          =   195
               Index           =   1
               Left            =   4305
               TabIndex        =   38
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   344
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Anormal"
            End
            Begin VB.CheckBox CPA001_10 
               Caption         =   "Proceso inflamatorio"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   510
               Width           =   1935
            End
            Begin VB.TextBox CPA001_22 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2820
               TabIndex        =   40
               Top             =   450
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.TextBox CPA001_23 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2820
               TabIndex        =   41
               Top             =   750
               Visible         =   0   'False
               Width           =   2715
            End
            Begin Threed.SSOption CPA001_09 
               Height          =   315
               Index           =   0
               Left            =   1500
               TabIndex        =   37
               Top             =   180
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Negativo a cáncer (Normal)"
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Estado de células:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   9
               Left            =   135
               TabIndex        =   70
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Grado:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   10
               Left            =   2250
               TabIndex        =   69
               Top             =   480
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Células Endometriales en > 50 años:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   11
               Left            =   150
               TabIndex        =   68
               Top             =   750
               Visible         =   0   'False
               Width           =   2625
            End
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "1.- Características de la muestra"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1335
            Index           =   4
            Left            =   -74880
            TabIndex        =   62
            Top             =   360
            Width           =   6735
            Begin VB.TextBox CPA001_21 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   530
               TabIndex        =   29
               Top             =   930
               Width           =   6075
            End
            Begin VB.Frame CPA001_00 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   11
               Left            =   1860
               TabIndex        =   104
               Top             =   720
               Width           =   3255
               Begin Threed.SSOption CPA001_08 
                  Height          =   195
                  Index           =   1
                  Left            =   1830
                  TabIndex        =   28
                  Top             =   0
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Ausentes"
               End
               Begin Threed.SSOption CPA001_08 
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   27
                  Top             =   0
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Presentes"
               End
            End
            Begin VB.Frame CPA001_00 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   10
               Left            =   1860
               TabIndex        =   103
               Top             =   480
               Width           =   3495
               Begin Threed.SSOption CPA001_07 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   25
                  Top             =   -60
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Muestra con sangre"
               End
               Begin Threed.SSOption CPA001_07 
                  Height          =   195
                  Index           =   1
                  Left            =   1830
                  TabIndex        =   26
                  Top             =   0
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Muestra mal fijada"
               End
            End
            Begin VB.Frame CPA001_00 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   9
               Left            =   1860
               TabIndex        =   102
               Top             =   240
               Width           =   4815
               Begin Threed.SSOption CPA001_06 
                  Height          =   195
                  Index           =   1
                  Left            =   1830
                  TabIndex        =   23
                  Top             =   0
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Limitada"
               End
               Begin Threed.SSOption CPA001_06 
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   22
                  Top             =   0
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Adecuada"
               End
               Begin Threed.SSOption CPA001_06 
                  Height          =   195
                  Index           =   2
                  Left            =   3540
                  TabIndex        =   24
                  Top             =   0
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Inadecuada"
               End
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Otro"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   8
               Left            =   150
               TabIndex        =   66
               Top             =   960
               Width           =   330
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Motivo:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   150
               TabIndex        =   65
               Top             =   480
               Width           =   1680
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Característica:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   150
               TabIndex        =   64
               Top             =   240
               Width           =   1680
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Células Endocervicales:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   150
               TabIndex        =   63
               Top             =   720
               Width           =   1680
            End
         End
         Begin VB.Frame CPA001_00 
            Caption         =   "4.- Diagnóstico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1215
            Index           =   7
            Left            =   -74880
            TabIndex        =   72
            Top             =   4050
            Width           =   6735
            Begin VB.Frame CPA001_00 
               BorderStyle     =   0  'None
               Height          =   210
               Index           =   14
               Left            =   1800
               TabIndex        =   107
               Top             =   960
               Width           =   4815
               Begin Threed.SSOption CPA001_14 
                  Height          =   315
                  Index           =   2
                  Left            =   2790
                  TabIndex        =   59
                  Top             =   -60
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   556
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Maligno no especificado"
               End
               Begin Threed.SSOption CPA001_14 
                  Height          =   195
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   58
                  Top             =   0
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Adenocarcinoma"
               End
               Begin Threed.SSOption CPA001_14 
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   57
                  Top             =   0
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Invasor"
               End
            End
            Begin VB.Frame CPA001_00 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   13
               Left            =   1800
               TabIndex        =   106
               Top             =   720
               Width           =   4695
               Begin Threed.SSOption CPA001_13 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   55
                  Top             =   -60
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Bajo Grado"
               End
               Begin Threed.SSOption CPA001_13 
                  Height          =   195
                  Index           =   1
                  Left            =   1350
                  TabIndex        =   56
                  Top             =   0
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "Alto Grado"
               End
            End
            Begin VB.Frame CPA001_00 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   105
               Top             =   480
               Width           =   5175
               Begin Threed.SSOption CPA001_12 
                  Height          =   195
                  Index           =   2
                  Left            =   1830
                  TabIndex        =   53
                  Top             =   0
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "ASC-H"
               End
               Begin Threed.SSOption CPA001_12 
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   52
                  Top             =   0
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "ASC-US"
               End
               Begin Threed.SSOption CPA001_12 
                  Height          =   195
                  Index           =   3
                  Left            =   3540
                  TabIndex        =   54
                  Top             =   0
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   344
                  _Version        =   262144
                  ForeColor       =   0
                  Caption         =   "AGUS"
               End
            End
            Begin Threed.SSOption CPA001_12 
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   51
               Top             =   180
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   262144
               ForeColor       =   0
               Caption         =   "Negativo a neoplasia maligna"
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Lesión intraepitelial de:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   12
               Left            =   60
               TabIndex        =   74
               Top             =   720
               Width           =   1650
            End
            Begin VB.Label CPA001_01 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Sospechoso CA:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   13
               Left            =   540
               TabIndex        =   73
               Top             =   960
               Width           =   1170
            End
         End
         Begin VB.Label CPA001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   5690
            Index           =   14
            Left            =   60
            TabIndex        =   93
            Top             =   360
            Width           =   6830
         End
         Begin VB.Label CPA001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Servicio Procedencia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   135
            TabIndex        =   108
            Top             =   510
            Width           =   1470
         End
         Begin VB.Label CPA001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   5690
            Index           =   15
            Left            =   -74940
            TabIndex        =   94
            Top             =   360
            Width           =   6830
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   114
      Top             =   1740
      Width           =   7155
      Begin VB.ComboBox cmbResponsable 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   180
         Width           =   3105
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   5580
         TabIndex        =   116
         Top             =   150
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Realiza Prueba"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   118
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F.Resultado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4620
         TabIndex        =   117
         Top             =   180
         Width           =   945
      End
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   60
      TabIndex        =   98
      Top             =   8940
      Width           =   7155
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmCitoPatologia.frx":0D02
         DownPicture     =   "frmCitoPatologia.frx":11C6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3675
         Picture         =   "frmCitoPatologia.frx":16B2
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprime (F3)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   60
         Picture         =   "frmCitoPatologia.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmCitoPatologia.frx":2077
         DownPicture     =   "frmCitoPatologia.frx":24D7
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2235
         Picture         =   "frmCitoPatologia.frx":294C
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   180
         Width           =   1365
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1695
      Left            =   60
      TabIndex        =   97
      Top             =   15
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2990
   End
   Begin VB.Frame CPA002 
      Caption         =   "Líquidos Especiales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   60
      TabIndex        =   85
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin TabDlg.SSTab CPA002_00 
         Height          =   4575
         Left            =   60
         TabIndex        =   112
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8070
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Toma de Muestra"
         TabPicture(0)   =   "frmCitoPatologia.frx":2DC1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "CPA001_01(16)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "CPA002_01(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "CPA002_01(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CPA002_01(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CPA002_01(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "CPA001_01(19)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "CPA002_02"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "CPA002_03"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "CPA002_04"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "CPA002_05"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "CPA002_09"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "Resultado"
         TabPicture(1)   =   "frmCitoPatologia.frx":2DDD
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CPA001_01(17)"
         Tab(1).Control(1)=   "CPA002_01(4)"
         Tab(1).Control(2)=   "CPA002_01(5)"
         Tab(1).Control(3)=   "CPA002_01(6)"
         Tab(1).Control(4)=   "CPA002_06"
         Tab(1).Control(5)=   "CPA002_08"
         Tab(1).Control(6)=   "CPA002_07"
         Tab(1).Control(7)=   "CPA002_10"
         Tab(1).ControlCount=   8
         Begin VB.ComboBox CPA002_10 
            Height          =   315
            Left            =   -73110
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox CPA002_09 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1650
            TabIndex        =   109
            Top             =   780
            Width           =   5115
         End
         Begin VB.TextBox CPA002_07 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   -74850
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1200
            Width           =   6675
         End
         Begin VB.TextBox CPA002_08 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   765
            Left            =   -74850
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   2190
            Width           =   6675
         End
         Begin VB.TextBox CPA002_06 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   -72060
            TabIndex        =   34
            Top             =   450
            Width           =   2595
         End
         Begin VB.TextBox CPA002_05 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   885
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   33
            Top             =   3580
            Width           =   6675
         End
         Begin VB.TextBox CPA002_04 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   885
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   2280
            Width           =   6675
         End
         Begin VB.TextBox CPA002_03 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   1290
            Width           =   6675
         End
         Begin VB.TextBox CPA002_02 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1650
            MaxLength       =   35
            TabIndex        =   30
            Top             =   450
            Width           =   2595
         End
         Begin VB.Label CPA001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Servicio Procedencia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   110
            Top             =   810
            Width           =   1470
         End
         Begin VB.Label CPA002_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Procedimiento realizado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   -74880
            TabIndex        =   92
            Top             =   870
            Width           =   1695
         End
         Begin VB.Label CPA002_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   -74880
            TabIndex        =   91
            Top             =   1980
            Width           =   825
         End
         Begin VB.Label CPA002_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de realización del procedimiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   -74880
            TabIndex        =   90
            Top             =   480
            Width           =   2760
         End
         Begin VB.Label CPA002_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Historia Breve"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   89
            Top             =   3360
            Width           =   1005
         End
         Begin VB.Label CPA002_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Antecedentes de Importancia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   88
            Top             =   2040
            Width           =   2130
         End
         Begin VB.Label CPA002_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Muestra"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   87
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label CPA002_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de toma"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label CPA001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   4140
            Index           =   16
            Left            =   60
            TabIndex        =   95
            Top             =   360
            Width           =   6825
         End
         Begin VB.Label CPA001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   4140
            Index           =   17
            Left            =   -74940
            TabIndex        =   96
            Top             =   360
            Width           =   6825
         End
      End
   End
End
Attribute VB_Name = "frmCitoPatologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultados de CitoPatología
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable

Dim ml_idUsuario As Long
Dim ml_idOrden As Long
Dim ml_nombrePrueba As String
Dim ml_idAnalisis As Long
Dim ml_idPaciente As Long
Dim ml_resultado As String
Dim ml_observacion As String
Dim ml_nombreMedico As String
Dim ml_nombrePaciente As String
Dim ml_nombreRealiza As Long
Dim ml_areaTrabajo As Long
Dim ml_CodigoPruebaSeleccionada As String
Dim ml_idPrueba As String
Dim ml_DetalleOrden As New ADODB.Recordset
Dim ml_idOrdenLab As Long
Dim ml_FechaNacimiento As Date
Dim ml_idTipoSexo As Long
Dim ml_NoMuestraBotonGrabar As Boolean
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let NoMuestraBotonGrabar(lValue As Boolean)
   ml_NoMuestraBotonGrabar = lValue
   If ml_NoMuestraBotonGrabar = True Then
      cmdGrabar.Visible = False
   End If
End Property




Property Let idTipoSexo(lValue As Long)
    ml_idTipoSexo = lValue
End Property
Property Let FechaNacimiento(lValue As Date)
    ml_FechaNacimiento = lValue
End Property
Property Let idOrdenLab(lValue As Long)
   ml_idOrdenLab = lValue
End Property

Property Let CodigoPruebaSeleccionada(lValue As String)
   ml_CodigoPruebaSeleccionada = lValue
End Property

Property Let DetalleOrden(lValue As ADODB.Recordset)
  Set ml_DetalleOrden = lValue
End Property

Sub CargaDataCombos()
  mo_cmbResponsable.BoundColumn = "idEmpleado"
  mo_cmbResponsable.ListField = "ApNom"
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =17")
  Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
  If mo_CabeceraReportes.NOpuedeModificarResponsable(sghAgregar, sighentidades.Usuario, mo_cmbResponsable.RowSource) Then
     mo_cmbResponsable.BoundText = Trim(Str(sighentidades.Usuario))
     Me.cmbResponsable.Enabled = False
  End If
  Set mo_CabeceraReportes = Nothing
End Sub

Property Let AreaTrabajo(lValue As Long)
    ml_areaTrabajo = lValue
End Property

Property Get AreaTrabajo() As Long
  AreaTrabajo = ml_areaTrabajo
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

Property Let idOrden(lValue As Long)
   ml_idOrden = lValue
End Property

Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property

Property Let idPrueba(lValue As String)
   ml_idPrueba = lValue
End Property

Property Get idPrueba() As String
   idPrueba = ml_idPrueba
End Property

Property Let nombrePrueba(lValue As String)
   ml_nombrePrueba = lValue
End Property

Property Get nombrePrueba() As String
   nombrePrueba = ml_nombrePrueba
End Property

Property Let idAnalisis(lValue As Long)
   ml_idAnalisis = lValue
End Property

Property Get idAnalisis() As Long
   idAnalisis = ml_idAnalisis
End Property

Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property

Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property

Property Let nombreMedico(lValue As String)
   ml_nombreMedico = lValue
End Property

Property Get nombreMedico() As String
   nombreMedico = ml_nombreMedico
End Property

Property Let nombrePaciente(lValue As String)
   ml_nombrePaciente = lValue
End Property

Property Get nombrePaciente() As String
   nombrePaciente = ml_nombrePaciente
End Property

Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyReturn
      SendKeys "{TAB}"
    Case vbKeyF3
      cmdImprimir_Click
    Case vbKeyEscape
      cmdCancelar_Click
    Case vbKeyF2
      cmdGrabar_Click
  End Select
End Sub

Private Sub TopBoton(Fra As Frame)
  'If EmpleadoTrabajaEnLaboratorio(sighEntidades.Usuario) = True Then
    Fra.Enabled = True
  'Else
  '  Fra.Enabled = False
  'End If
  Fra.Visible = True
  Fra.Caption = ml_nombrePrueba
  Me.Caption = Fra.Caption
  Fra.Top = UcPacienteDatos1.Top + UcPacienteDatos1.Height + 700   '350
  fraBoton.Top = Fra.Top + Fra.Height
  Me.Height = fraBoton.Top + fraBoton.Height + 500
End Sub

Private Sub cmbResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdGrabar_Click()
  If cmbResponsable.Text = "" Then
    MsgBox "Debe Seleccionar el personal que realizó la prueba", vbInformation, "SIGH "
    cmbResponsable.SetFocus
    Exit Sub
  End If
  If txtFresultado.Text = sighentidades.FECHA_VACIA_DMY Then
     MsgBox "Por favor ingrese la Fecha del Resultado", vbInformation, "SIGH "
     Exit Sub
  End If
  ml_nombreRealiza = mo_cmbResponsable.BoundText
  If ml_CodigoPruebaSeleccionada = "CPA001" Then 'Papanicolau
    'CPA001
    ml_resultado = CPA001_26.Text & "\" & _
                   CPA001_03(0).Value & "\" & CPA001_03(1).Value & "\" & _
                   CPA001_16.Text & "\" & CPA001_17.Text & "\" & CPA001_18.Text & "\" & _
                   CPA001_04(0).Value & "\" & CPA001_04(1).Value & "\" & CPA001_04(2).Value & "\" & CPA001_04(3).Value & "\" & CPA001_04(4).Value & "\" & CPA001_04(5).Value & "\" & _
                   CPA001_05(0).Value & "\" & CPA001_05(1).Value & "\" & CPA001_05(2).Value & "\" & CPA001_05(3).Value & "\" & CPA001_05(4).Value & "\" & CPA001_05(5).Value & "\" & CPA001_05(6).Value & "\" & CPA001_05(7).Value & "\" & _
                   CPA001_19.Text & "\" & CPA001_20.Text & "\" & _
                   CPA001_06(0).Value & "\" & CPA001_06(1).Value & "\" & CPA001_06(2).Value & "\" & _
                   CPA001_07(0).Value & "\" & CPA001_07(1).Value & "\" & _
                   CPA001_08(0).Value & "\" & CPA001_08(1).Value & "\" & _
                   CPA001_21.Text & "\" & _
                   CPA001_09(0).Value & "\" & CPA001_09(1).Value & "\" & _
                   CPA001_10.Value & "\" & _
                   CPA001_22.Text & "\" & CPA001_23.Text & "\" & _
                   CPA001_11(0).Value & "\" & CPA001_11(1).Value & "\" & CPA001_11(2).Value & "\" & CPA001_11(3).Value & "\" & CPA001_11(4).Value & "\" & CPA001_11(5).Value & "\" & CPA001_11(6).Value & "\" & CPA001_11(7).Value & "\" & _
                   CPA001_24.Text & "\" & _
                   CPA001_12(0).Value & "\" & CPA001_12(1).Value & "\" & CPA001_12(2).Value & "\" & CPA001_12(3).Value & "\" & _
                   CPA001_13(0).Value & "\" & CPA001_13(1).Value & "\" & _
                   CPA001_14(0).Value & "\" & CPA001_14(1).Value & "\" & CPA001_14(2).Value & "\" & _
                   CPA001_25.Text
  ElseIf ml_CodigoPruebaSeleccionada = "CPA002" Then
    'BAAF 'Impronta 'PAP de liquidos y Fluidos
    'CPA002
    ml_resultado = CPA002_02.Text & "\" & CPA002_09.Text & "\" & CPA002_03.Text & "\" & CPA002_04.Text & "\" & CPA002_05.Text & "\" & CPA002_06.Text & "\" & CPA002_10.Text & "\" & CPA002_07.Text & "\" & CPA002_08.Text
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  mo_ReglasLaboratorio.LabIngresaResultados idPrueba, idOrden, ml_resultado, ml_observacion, idUsuario, ml_nombreRealiza, ml_DetalleOrden, ml_idOrdenLab, "", "", 0, CDate(txtFresultado.Text), mo_lcNombrePc, mo_lnIdTablaLISTBARITEMS, Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre, Me.Caption
End Sub

Private Sub cmdImprimir_Click()
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_CodigoPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
    mo_ReglasLaboratorio.LabImprimeCabeceraResultados UcPacienteDatos1.idPaciente, nombrePaciente, UcPacienteDatos1.NroHistoriaClinica, ldFechaResultado, _
                         nombreMedico + mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(ml_idOrden)
    mo_ReglasLaboratorio.LabImprimeResultadosCPA ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
    mo_ReglasLaboratorio.LabImprimePieResultados
  Else
    MsgBox "Debe grabar los resultados antes de poder imprimirlos", vbInformation, ""
  End If
End Sub

Private Sub CPA001_03_Click(Index As Integer, Value As Integer)
'  MsgBox Index
  If Index = 0 Then
    If Value = 1 Then
      CPA001_01(0).Visible = True
      CPA001_16.Visible = True
      CPA001_16.Enabled = True
    Else
      CPA001_16.Visible = False
      CPA001_01(0).Visible = False
      CPA001_16.Enabled = False
      CPA001_16.Text = ""
    End If
  ElseIf Index = 1 Then
    If Value = 1 Then
      CPA001_16.Visible = False
      CPA001_01(0).Visible = False
      CPA001_16.Enabled = False
      CPA001_16.Text = ""
    Else
      CPA001_01(0).Visible = True
      CPA001_16.Visible = True
      CPA001_16.Enabled = True
    End If
  End If
End Sub

Private Sub CPA001_03_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_04_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_05_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_06_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_07_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_08_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_09_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_10_Click()
  If CPA001_10.Value = 1 Then
    CPA001_22.Visible = True
    CPA001_22.Enabled = True
    CPA001_01(10).Visible = True
  Else
    CPA001_01(10).Visible = False
    CPA001_22.Visible = False
    CPA001_22.Enabled = False
  End If
End Sub

Private Sub CPA001_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_11_Click(Index As Integer)
  If CPA001_11(7).Value = 1 Then
    CPA001_24.Visible = True
  Else
    CPA001_24.Visible = False
    CPA001_24.Text = ""
  End If
End Sub

Private Sub CPA001_11_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_12_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_13_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_14_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_16_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_17_GotFocus()
  SeleccionaTexto CPA001_17
End Sub

Private Sub CPA001_17_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_18_GotFocus()
  SeleccionaTexto CPA001_18
End Sub

Private Sub CPA001_18_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_19_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_20_GotFocus()
  SeleccionaTexto CPA001_20
End Sub

Private Sub CPA001_20_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_21_GotFocus()
  SeleccionaTexto CPA001_21
End Sub

Private Sub CPA001_21_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_22_GotFocus()
  SeleccionaTexto CPA001_22
End Sub

Private Sub CPA001_22_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_23_GotFocus()
  SeleccionaTexto CPA001_23
End Sub

Private Sub CPA001_23_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_24_GotFocus()
  SeleccionaTexto CPA001_24
End Sub

Private Sub CPA001_24_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_25_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA001_26_GotFocus()
  SeleccionaTexto CPA001_26
End Sub

Private Sub CPA001_26_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_02_GotFocus()
  SeleccionaTexto CPA002_02
End Sub

Private Sub CPA002_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_06_GotFocus()
  SeleccionaTexto CPA002_06
End Sub

Private Sub CPA002_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_09_GotFocus()
  SeleccionaTexto CPA002_09
End Sub

Private Sub CPA002_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub CPA002_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
  Set mo_cmbResponsable.MiComboBox = cmbResponsable
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
  Me.txtFresultado.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  Me.UcPacienteDatos1.idPaciente = ml_idPaciente
  Me.UcPacienteDatos1.FechaRegistro = Now
  If ml_idPaciente = 0 Then
     Me.UcPacienteDatos1.idTipoSexo = ml_idTipoSexo
     Me.UcPacienteDatos1.FechaNacimiento = ml_FechaNacimiento
     Me.UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta ml_nombrePaciente
  Else
     Me.UcPacienteDatos1.CargarDatosDePacienteALosControles
  End If
  Me.UcPacienteDatos1.DeshabilitarFrames True
  CargaDataCombos
  cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, sighentidades.NombreUsuario)
  'If EmpleadoTrabajaEnLaboratorio(sighEntidades.Usuario) = True Then
    cmdGrabar.Enabled = True
  'Else
  '  cmdGrabar.Enabled = False
  'End If
  ml_resultado = ""
  ml_observacion = ""
  
  If ml_CodigoPruebaSeleccionada = "CPA001" Then  'Papanicolao
    TopBoton CPA001
    If UcPacienteDatos1.Edad > 49 Then
      CPA001_01(11).Visible = True
      CPA001_23.Visible = True
      CPA001_23.Enabled = True
    Else
      CPA001_01(11).Visible = False
      CPA001_23.Visible = False
      CPA001_23.Enabled = False
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "CPA002" Then
    'BAAF 'Impronta 'PAP de liquidos y Fluidos
    TopBoton CPA002
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  'Recupera información si es que ya esta grabado
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_resultado = "" Or Val(ml_nombreRealiza) = 0 Then Exit Sub
  Me.txtFresultado.Text = Format(IIf(ldFechaResultado = 0, Now, ldFechaResultado), sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, mo_ReglasLaboratorio.LabEmpleado(ml_nombreRealiza))
  'If cmbResponsable.Text <> "" Then cmbResponsable.Enabled = False
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim Temp As String
  'Asigna la información recuperada en el formulario
  If ml_CodigoPruebaSeleccionada = "CPA001" Then  'Papanicolau
    CPA001_26.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_03(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_03(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_16.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_17.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_18.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_04(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_04(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_04(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_04(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_04(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_04(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_05(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_05(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_05(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_05(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_05(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_05(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_05(6).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_05(7).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_19.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_20.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_06(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_06(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_06(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_07(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_07(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_08(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_08(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_21.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_09(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_09(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_10.Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_22.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_23.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_11(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_11(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_11(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_11(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_11(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_11(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_11(6).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_11(7).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_24.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_12(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_12(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_12(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_12(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_13(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_13(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_14(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_14(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_14(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA001_25.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
  ElseIf ml_CodigoPruebaSeleccionada = "CPA002" Then
    'BAAF 'Impronta 'PAP de liquidos y Fluidos
    CPA002_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA002_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA002_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA002_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA002_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA002_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA002_10.ListIndex = Ubica_En_Combo(CPA002_10, Temp)
    CPA002_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    CPA002_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
  End If
End Sub

