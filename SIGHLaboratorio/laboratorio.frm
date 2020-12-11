VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form frmLaboratorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Órdenes de Laboratorio"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   ForeColor       =   &H80000004&
   Icon            =   "laboratorio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosAtencion 
      Caption         =   "Datos de Cabecera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   13755
      Begin VB.CommandButton btnTamizaje 
         BackColor       =   &H8000000D&
         Caption         =   "Tamizaje"
         Height          =   360
         Left            =   9195
         TabIndex        =   68
         Top             =   3270
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   60
         TabIndex        =   57
         Top             =   3135
         Width           =   8895
         Begin VB.TextBox txtEG 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2250
            MaxLength       =   30
            TabIndex        =   13
            ToolTipText     =   "Edad Gestacional= (Hoy - FUM)/7......   (1mes gestacional=28 días)"
            Top             =   135
            Width           =   525
         End
         Begin MSMask.MaskEdBox txtFum 
            Height          =   315
            Left            =   480
            TabIndex        =   12
            ToolTipText     =   "Fecha de última mestruación"
            Top             =   135
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "FUM"
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
            Left            =   90
            TabIndex        =   59
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "EG"
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
            Left            =   1950
            TabIndex        =   58
            Top             =   165
            Width           =   225
         End
      End
      Begin Threed.SSOption ssoptPaciente 
         Height          =   255
         Left            =   345
         TabIndex        =   46
         Top             =   645
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   262144
         Caption         =   "Paciente (con N° Cuenta)"
         Value           =   -1
      End
      Begin Threed.SSOption ssoptExterno 
         Height          =   255
         Left            =   3600
         TabIndex        =   47
         Top             =   630
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   450
         _Version        =   262144
         Caption         =   "Externo (con Boleta)"
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3150
         MaxLength       =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   795
      End
      Begin VB.TextBox txtNmovimiento 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox txtDx 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6795
         MaxLength       =   30
         TabIndex        =   28
         ToolTipText     =   "Ingrese el Dx (4 dígitos)"
         Top             =   4665
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtNombreDx 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8265
         TabIndex        =   30
         Top             =   4665
         Visible         =   0   'False
         Width           =   5445
      End
      Begin VB.CommandButton cmdBuscaDx 
         Caption         =   "..."
         Height          =   315
         Left            =   7905
         TabIndex        =   29
         Top             =   4665
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtResultadoFinal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6795
         MaxLength       =   100
         TabIndex        =   32
         Top             =   4965
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.CheckBox chkDxDefinitivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Dx Definitivo"
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
         Left            =   10410
         TabIndex        =   33
         Top             =   4380
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtZonaCuerpo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6795
         MaxLength       =   100
         TabIndex        =   24
         Top             =   4095
         Visible         =   0   'False
         Width           =   6885
      End
      Begin VB.ComboBox cmbPersonaRecoje 
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
         Left            =   6795
         TabIndex        =   26
         Top             =   4365
         Visible         =   0   'False
         Width           =   6900
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   4560
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SIGHLaboratorio.UcPacienteDatos UcPacienteDatos1 
         Height          =   3015
         Left            =   9120
         TabIndex        =   18
         Top             =   150
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5318
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   2415
         Left            =   60
         TabIndex        =   36
         Top             =   660
         Width           =   8955
         Begin VB.TextBox txtColegiatura 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7515
            MaxLength       =   30
            TabIndex        =   61
            Top             =   1650
            Width           =   1245
         End
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
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1270
            Width           =   4350
         End
         Begin VB.ComboBox cmbMedicos 
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
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1630
            Visible         =   0   'False
            Width           =   4020
         End
         Begin VB.TextBox txtMedico 
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
            Left            =   1500
            TabIndex        =   14
            Top             =   1630
            Visible         =   0   'False
            Width           =   4350
         End
         Begin VB.Frame fraExterno 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1020
            Left            =   90
            TabIndex        =   41
            Top             =   195
            Visible         =   0   'False
            Width           =   8685
            Begin VB.CommandButton cmdBuscarPago 
               Caption         =   "..."
               Height          =   315
               Left            =   3090
               TabIndex        =   50
               TabStop         =   0   'False
               ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
               Top             =   60
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.TextBox txtNroOrden 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7005
               MaxLength       =   30
               TabIndex        =   9
               Top             =   60
               Width           =   1605
            End
            Begin VB.TextBox txtNserie 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1365
               MaxLength       =   30
               TabIndex        =   7
               Top             =   60
               Width           =   525
            End
            Begin VB.TextBox txtNboleta 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1935
               MaxLength       =   30
               TabIndex        =   8
               Top             =   60
               Width           =   1125
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "N° Orden"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   6180
               TabIndex        =   43
               Top             =   60
               Width           =   795
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "N° Boleta"
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
               Left            =   60
               TabIndex        =   42
               Top             =   120
               Width           =   780
            End
         End
         Begin VB.Frame fraHospitalizado 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1095
            Left            =   120
            TabIndex        =   37
            Top             =   210
            Visible         =   0   'False
            Width           =   8760
            Begin VB.CommandButton cmbBuscaReceta 
               Height          =   330
               Left            =   8355
               Picture         =   "laboratorio.frx":0CCA
               Style           =   1  'Graphical
               TabIndex        =   67
               Top             =   15
               Width           =   300
            End
            Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
               Height          =   330
               Left            =   2685
               Picture         =   "laboratorio.frx":1254
               Style           =   1  'Graphical
               TabIndex        =   66
               Top             =   15
               Width           =   300
            End
            Begin VB.TextBox txtNcita 
               Alignment       =   1  'Right Justify
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
               Left            =   7260
               MaxLength       =   30
               TabIndex        =   64
               Top             =   705
               Width           =   1380
            End
            Begin VB.TextBox txtNreceta 
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
               Left            =   7380
               MaxLength       =   30
               TabIndex        =   52
               Top             =   30
               Width           =   945
            End
            Begin VB.CheckBox chkPlanNoCubre 
               Alignment       =   1  'Right Justify
               Caption         =   "IAFA NO cubre"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7065
               TabIndex        =   48
               Top             =   390
               Width           =   1575
            End
            Begin VB.TextBox txtDatosDeCuenta 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3000
               TabIndex        =   4
               Top             =   30
               Width           =   3315
            End
            Begin VB.TextBox txtNcuenta 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1380
               MaxLength       =   30
               TabIndex        =   3
               Top             =   30
               Width           =   1245
            End
            Begin VB.TextBox txtProcedencia 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1380
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   360
               Width           =   3075
            End
            Begin VB.TextBox txtPlan 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1380
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   690
               Width           =   4365
            End
            Begin MSDataListLib.DataCombo cmbFormaPago 
               Height          =   330
               Left            =   4515
               TabIndex        =   49
               Top             =   360
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "N° Cita"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6570
               TabIndex        =   65
               Top             =   750
               Width           =   570
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "N° Receta"
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
               Left            =   6510
               TabIndex        =   53
               Top             =   90
               Width           =   840
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "N° Cuenta"
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
               Left            =   0
               TabIndex        =   40
               Top             =   60
               Width           =   855
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Procedencia"
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
               Left            =   0
               TabIndex        =   39
               Top             =   390
               Width           =   990
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Fte.Finan/IAFA"
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
               Left            =   0
               TabIndex        =   38
               Top             =   750
               Width           =   1215
            End
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "N° Colegiatura"
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
            Left            =   6360
            TabIndex        =   62
            Top             =   1680
            Width           =   1170
         End
         Begin VB.Label lblDx 
            AutoSize        =   -1  'True
            Caption         =   "Diagnóstico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   120
            TabIndex        =   51
            Top             =   2025
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Registra Orden"
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
            Left            =   120
            TabIndex        =   45
            Top             =   1300
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Médico solicita"
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
            Left            =   120
            TabIndex        =   44
            Top             =   1665
            Width           =   1155
         End
      End
      Begin MSMask.MaskEdBox txtFrealizaCpt 
         Height          =   315
         Left            =   7170
         TabIndex        =   55
         Top             =   270
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Label lblOrdenPago 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N° Orden de Pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   11940
         TabIndex        =   63
         Top             =   3270
         Width           =   1515
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "F.Realiza CPT"
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
         Left            =   6120
         TabIndex        =   56
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Left            =   2580
         TabIndex        =   22
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Movimiento"
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
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Reg"
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
         Left            =   4080
         TabIndex        =   21
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Resultado Final"
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
         Left            =   5415
         TabIndex        =   31
         Top             =   4995
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Diagnóstico"
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
         Left            =   5415
         TabIndex        =   27
         Top             =   4575
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Muestra"
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
         Left            =   5415
         TabIndex        =   23
         Top             =   4110
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Recoje"
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
         Left            =   5415
         TabIndex        =   25
         Top             =   4290
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin SIGHLaboratorio.ucInsumocpt ucProductos 
      Height          =   2895
      Left            =   60
      TabIndex        =   15
      Top             =   3870
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   6006
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   34
      Top             =   8460
      Width           =   13755
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   180
         Picture         =   "laboratorio.frx":17DE
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "laboratorio.frx":1CB7
         DownPicture     =   "laboratorio.frx":2117
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5340
         Picture         =   "laboratorio.frx":258C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "laboratorio.frx":2A01
         DownPicture     =   "laboratorio.frx":2EC5
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6870
         Picture         =   "laboratorio.frx":33B1
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdConsumoPaciente 
      Height          =   1605
      Left            =   90
      TabIndex        =   54
      Top             =   6810
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   2831
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "laboratorio.frx":389D
      Caption         =   "Exámenes históricos del Paciente (Consulta Externa, Hospitalización, Emergencia)"
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "N° Boleta"
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
      Left            =   8250
      TabIndex        =   35
      Top             =   1830
      Width           =   780
   End
End
Attribute VB_Name = "frmLaboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento para Anatomía Patológica, Banco de Sangre y Patología Clínica
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_IdMovimiento As Long
Dim mi_Opcion As sghOpciones
Dim ms_MensajeError As String
Dim ml_idUsuario As Long
Dim ml_puntoCarga As Long
Dim mb_ExistenDatos As Boolean
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim wxParametro302 As String, lnIdTipoServicio As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdEstado As New sighentidades.ListaDespleglable
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable
Dim mo_cmbPersonaRecoje As New sighentidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim lbPrimeraVez As Boolean
Dim ml_IdTipoFinanciamiento As Long
Dim ml_idPaciente As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdFuenteFinanciamiento  As Long
Dim ml_IdServicioPaciente As Long
Dim ml_IdDiagnostico As Long
Dim oDOPaciente As New doPaciente
Dim oDOLabMovimiento As New DOLabMovimiento
Dim oDoLabMovimientoLaboratorio As New DoLabMovimientoLaboratorio
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim rsProductosCPT As Recordset
Dim rsProductos As Recordset
Dim oRsFormaPago As New Recordset
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_areaTrabajo As Long
Dim mo_cmbMedicos As New sighentidades.ListaDespleglable
Const lcConstanteMovimientoSalida As String = "S"
Dim lnIdReceta As Long
Dim ml_SeEligioGridBoleta As Boolean
Dim wxParametro509 As String
Dim lcMedicoDNI As String, lcCama As String, lcDxCodigo As String, lcDx As String, lcUPS As String
Dim lnEpsPorcentaje As Double
Dim lbCuentaDeEmergenciaCerrada As Boolean
Dim wxParametro578 As String

Property Let SeEligioGridBoleta(lValue As Boolean)
    ml_SeEligioGridBoleta = lValue
End Property
Property Get SeEligioGridBoleta() As Boolean
    SeEligioGridBoleta = ml_SeEligioGridBoleta
End Property
Property Let AreaTrabajo(lValue As Long)
    ml_areaTrabajo = lValue
End Property

Property Get AreaTrabajo() As Long
  AreaTrabajo = ml_areaTrabajo
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property

Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

Property Let IdMovimiento(lValue As Long)
    ml_IdMovimiento = lValue
End Property

Property Get IdMovimiento() As Long
    IdMovimiento = ml_IdMovimiento
End Property

Property Let puntoCarga(lValue As Long)
    ml_puntoCarga = lValue
End Property

Property Get puntoCarga() As Long
    IdPuntoCarga = ml_puntoCarga
End Property

Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyReturn
    Case vbKeyEscape
    Case vbKeyF10
    Case vbKeyF2
      btnAceptar_Click
    End Select
End Sub

Private Sub btnAceptar_Click()
  If btnAceptar.Enabled = False Then Exit Sub
  Dim lcMedico9 As String
  mo_reglasComunes.DevuelveCamaYdniMedico lcMedico9, lcMedicoDNI, lcCama, Val(mo_cmbMedicos.BoundText), 0, ml_idPaciente
  Select Case mi_Opcion
    Case sghAgregar
      If ValidarDatosObligatorios() Then
        CargaDatosAlObjetosDeDatos
        If ValidarReglas() Then
          If AgregarDatos() Then
            Me.txtNmovimiento = oDOLabMovimiento.IdMovimiento
            MsgBox "Se agregó correctamente el Movimiento N° " & oDOLabMovimiento.IdMovimiento & Chr(13) & _
                    lblOrdenPago.Caption, vbInformation, Me.Caption
            'ml_IdMovimiento = oDOLabMovimiento.IdMovimiento
            'btnImprimir_Click
            LimpiarFormulario
          Else
            MsgBox "No se pudo agregar la Órden de Laboratorio" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
          End If
        End If
      End If
    Case sghModificar
      If ValidarDatosObligatorios() Then
        CargaDatosAlObjetosDeDatos
        If ValidarReglas() Then
          If ModificarDatos() Then
            MsgBox "Se Modificó correctamente el Movimiento N° " & oDOLabMovimiento.IdMovimiento & Chr(13) & lblOrdenPago.Caption, vbExclamation, Me.Caption
            Me.Visible = False
          Else
            MsgBox "No se pudo modificar la Órden de Laboratorio." & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
          End If
        End If
      End If
    Case sghEliminar
      If Val(txtNboleta.Text) = 0 Then
        If mo_ReglasLaboratorio.LabOrdenConResultados(ucProductos.FacturacionProductos, Val(txtNmovimiento.Text)) = True Then
          MsgBox "La orden que desea anular, ya tiene resultados ingresados." & Chr(13) & Chr(13) & "Esta orden no puede ser anulada.", vbInformation, "SIGH "
          Exit Sub
        End If
      End If
      If MsgBox("¿Realmente desea anular la Órden de Laboratorio?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
      If ValidarReglas() Then
        CargaDatosAlObjetosDeDatos
        If EliminarDatos() Then
          MsgBox "La Órden de Laboratorio fue Anulada correctamente.", vbInformation, Me.Caption
          Me.Visible = False
        Else
          MsgBox "No se pudo anular la Órden de Laboratorio." & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
        End If
      End If
  End Select
End Sub

Sub LimpiarFormulario()
    lbCuentaDeEmergenciaCerrada = False
    lblOrdenPago.Caption = ""
    lnEpsPorcentaje = 0
    txtNcuenta.Text = ""
    txtDatosDeCuenta.Text = ""
    txtProcedencia.Text = ""
    txtPlan.Text = ""
    txtNserie.Text = ""
    txtNboleta.Text = ""
    txtNroOrden.Text = ""
    If txtNcuenta.Visible = True Then txtNcuenta.SetFocus
    If txtNserie.Visible = True Then txtNserie.SetFocus
    Me.ucProductos.LimpiarGrilla
'    Me.ucProductos.AgregaProducto
    UcPacienteDatos1.LimpiarDatosDePaciente
    ml_idPaciente = 0
    ml_IdTipoFinanciamiento = 0
    ml_IdFuenteFinanciamiento = 0
    ml_IdServicioPaciente = 0
    ml_IdComprobantePago = 0
    chkDxDefinitivo.Visible = False
    chkDxDefinitivo.Value = 1
    ml_IdDiagnostico = 0
    txtDx.Text = ""
    txtNombreDx.Text = ""
    txtResultadoFinal.Text = ""
    cmbPersonaRecoje.Text = ""
    txtZonaCuerpo.Text = ""
    txtMedico.Text = ""
    chkPlanNoCubre.Value = 0
    cmbMedicos.ListIndex = -1
    cmbFormaPago.Text = ""
    txtNreceta.Text = ""
    chkPlanNoCubre.Value = 0
    Me.txtFum.Text = sighentidades.FECHA_VACIA_DMY
    Me.txtEG.Text = ""
    If cmbResponsable.Locked = False Then cmbResponsable.ListIndex = -1
    Set grdConsumoPaciente.DataSource = Nothing
    ucProductos.PermiteAgregarItems = True
    txtFrealizaCpt.Text = lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
    lcMedicoDNI = "": lcCama = "": lcDxCodigo = "": lcDx = "": lcUPS = ""
    Me.txtColegiatura.Text = ""
End Sub

Function ValidarDatosObligatorios() As Boolean
    On Error Resume Next
    Dim lnTabError As Integer
    ValidarDatosObligatorios = False
    If txtNcuenta.Text = "" And txtNboleta.Text = "" Then Exit Function
    ms_MensajeError = ""
    UcPacienteDatos1.CargarDatosAlObjetoDatos oDOPaciente
    If txtDatosDeCuenta.Text = "" Then
       If oDOPaciente.ApellidoPaterno = "" Then
          ms_MensajeError = ms_MensajeError & "- Registre el Apellido Paterno" & Chr(13)
          lnTabError = 1
       End If
       If oDOPaciente.ApellidoMaterno = "" Then
          ms_MensajeError = ms_MensajeError & "- Registre el Apellido Materno" & Chr(13)
          lnTabError = 1
       End If
       If oDOPaciente.PrimerNombre = "" Then
          ms_MensajeError = ms_MensajeError & "- Registre el Primer Nombre" & Chr(13)
          lnTabError = 1
       End If
       If oDOPaciente.idTipoSexo = 0 And Me.ucProductos.TieneResultadoAutomatico = True Then
          ms_MensajeError = ms_MensajeError & "- elija SEXO" & Chr(13)
          lnTabError = 1
       End If
       If oDOPaciente.FechaNacimiento = 0 And Me.ucProductos.TieneResultadoAutomatico = True Then
          ms_MensajeError = ms_MensajeError & "- ingrese F.NACIMIENTO" & Chr(13)
          lnTabError = 1
       End If
    Else
    End If
    
    If Me.txtColegiatura.Text = "" And Me.ucProductos.TieneResultadoAutomatico = True Then
       ms_MensajeError = ms_MensajeError & "- ingrese COLEGIATURA" & Chr(13)
       lnTabError = 1
    End If
    If cmbResponsable.Text = "" Then
       ms_MensajeError = ms_MensajeError & "- Elija persona que recibe muestra " & Chr(13)
       lnTabError = 2
    End If
    If lcBuscaParametro.SeleccionaFilaParametro(583) <> "S" And Trim(txtMedico.Text) = "" Then
       txtMedico.Text = "."
       If Me.txtColegiatura.Text = "" Then
          Me.txtColegiatura.Text = "."
       End If
    End If
    If Trim(txtMedico.Text) = "" Then
       ms_MensajeError = ms_MensajeError & "- Falta Nombre de Médico que ordena análisis " & Chr(13)
       lnTabError = 3
    End If
    Select Case mi_Opcion
    Case sghAgregar, sghModificar

        'Cpt
        Set rsProductosCPT = Me.ucProductos.FacturacionProductos
        If Not (rsProductosCPT.EOF And rsProductosCPT.BOF) Then
            rsProductosCPT.MoveFirst
            txtNroOrden.Text = rsProductosCPT.Fields!idOrden
            Do While Not rsProductosCPT.EOF
                If rsProductosCPT!idProducto = 0 Then
                    
                   rsProductosCPT.Delete
                   rsProductosCPT.Update
                Else
                   If rsProductosCPT!Cantidad <= 0 Then ms_MensajeError = ms_MensajeError & "- El producto CPT: " & rsProductosCPT!Codigo & " " & Trim(rsProductosCPT!NombreProducto) & "   Tiene problemas con la Cantidad" & Chr(13)
                   If rsProductosCPT!PrecioUnitario <= 0 And rsProductosCPT!SeUsaSinPrecio = False Then
                      If Val(Me.txtNboleta.Text) = 0 Then  'debb-05/04/2011
                         ms_MensajeError = ms_MensajeError & "- El producto CPT: " & rsProductosCPT!Codigo & " " & Trim(rsProductosCPT!NombreProducto) & "   Tiene problemas con el Precio" & Chr(13)
                      End If
                   End If
                   If rsProductosCPT!Cantidad < rsProductosCPT!cantidadFallada Then ms_MensajeError = ms_MensajeError & "- El producto CPT: " & rsProductosCPT!Codigo & " " & Trim(rsProductosCPT!NombreProducto) & "   la CANTIDAD FALLADA debe ser menor a la CANTIDAD" & Chr(13)
                End If
                rsProductosCPT.MoveNext
            Loop
        End If
        'If Me.ucProductos.DevuelveTotalPagar <= 0 Then ms_MensajeError = ms_MensajeError & "- El Importe Total es S/ 0.00" & Chr(13)
        'Insumos
        Set rsProductos = Me.ucProductos.FacturacionInsumos
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                If rsProductos!idProducto = 0 Or rsProductos!idProductoCPT = 0 Then
                   rsProductos.Delete
                   rsProductos.Update
                Else
                   If rsProductos!Cantidad <= 0 Then ms_MensajeError = ms_MensajeError & "- El INSUMO: " & rsProductos!Codigo & " " & Trim(rsProductos!NombreProducto) & "   Tiene problemas con la Cantidad" & Chr(13)
                   If rsProductos!PrecioUnitario <= 0 Then ms_MensajeError = ms_MensajeError & "- El INSUMO: " & rsProductos!Codigo & " " & Trim(rsProductos!NombreProducto) & "   Tiene problemas con el Precio" & Chr(13)
                   If rsProductos!Cantidad < rsProductos!cantidadFallada Then ms_MensajeError = ms_MensajeError & "- El INSUMO: " & rsProductos!Codigo & " " & Trim(rsProductos!NombreProducto) & "   la CANTIDAD FALLADA debe ser menor a la CANTIDAD" & Chr(13)
                   rsProductosCPT.MoveFirst
                   rsProductosCPT.Find "idProducto=" & rsProductos!idProductoCPT
                   If rsProductosCPT.EOF Then ms_MensajeError = ms_MensajeError & "- El INSUMO: " & rsProductos!Codigo & " " & Trim(rsProductos!NombreProducto) & "   no tiene Código CPT" & Chr(13)
                   
                End If
                rsProductos.MoveNext
            Loop
        End If
    End Select
    
    If ms_MensajeError = "" Then
       ValidarDatosObligatorios = True
    Else
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       On Error Resume Next
       Select Case lnTabError
       Case 1
           UcPacienteDatos1.SetFocusOnApellidoPaterno
       Case 2
           cmbResponsable.SetFocus
       Case 3
       End Select
    End If
End Function

Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With oDOLabMovimiento
'            .fecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .fecha = lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
            .IdlabEstado = 2 'Registrado     '1 -> Atendido
            .IdPuntoCarga = ml_puntoCarga
            .idTipoConcepto = sghTipoConceptoImagen.sghImgTCsalida  'Salidas con Orden de Pago
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .MovTipo = lcConstanteMovimientoSalida
        End With
        With oDoLabMovimientoLaboratorio
            '.CorrelativoAnual
            .IdComprobantePago = ml_IdComprobantePago
            .idCuentaAtencion = Val(txtNcuenta.Text)
            .idOrden = Val(txtNroOrden.Text)
            .IdPersonaTomaLab = Val(mo_cmbResponsable.BoundText)
            .IdUsuarioAuditoria = ml_idUsuario
            .idPersonaRecoge = Val(mo_cmbPersonaRecoje.BoundText)
            .OrdenaPrueba = Trim(txtMedico.Text)
            .idDiagnostico = ml_IdDiagnostico
            If ml_IdDiagnostico > 0 Then
               .EsDiagnosticoDefinitivo = IIf(chkDxDefinitivo.Value = 1, sghTipoDx.sghTipoDxDefinitivo, sghTipoDx.sghTipoDxPresuntivo)    '1-definitivo, 2-presuntivo
            Else
               .EsDiagnosticoDefinitivo = sghTipoDx.sghTipoDxNINGUNO
            End If
            .Paciente = Trim(oDOPaciente.ApellidoPaterno) & " " & Trim(oDOPaciente.ApellidoMaterno) & " " & oDOPaciente.PrimerNombre
            .idTipoSexo = oDOPaciente.idTipoSexo
            .FechaNacimiento = oDOPaciente.FechaNacimiento
            If IsDate(Me.txtFum.Text) Then
                .eo_eg = Val(Me.txtEG.Text)
                .Eo_FUM = CDate(Me.txtFum.Text)
            End If
            .colegiatura = Me.txtColegiatura.Text
        End With
        With oDOPaciente  'ya lo cargo en Validacion de Datos
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoFactOrdenServ
            .FechaCreacion = oDOLabMovimiento.fecha
            .fechaDespacho = oDOLabMovimiento.fecha
            .idCuentaAtencion = Val(txtNcuenta.Text)
            .IdEstadoFacturacion = sghEstadoFacturacion.sghAtendido
            .IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
            .idPaciente = ml_idPaciente
            .IdPuntoCarga = ml_puntoCarga
            .IdServicioPaciente = ml_IdServicioPaciente
            .IdTipoFinanciamiento = ml_IdTipoFinanciamiento
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .IdUsuarioDespacho = ml_idUsuario
            .FechaHoraRealizaCpt = txtFrealizaCpt.Text
        End With
    Case sghModificar
        With oDOLabMovimiento
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoLabMovimientoLaboratorio
            '.CorrelativoAnual
            '.IdComprobantePago = ml_IdComprobantePago
            '.idCuentaAtencion = Val(txtNcuenta.Text)
            '.IdOrden = Val(txtNroOrden.Text)
            .IdPersonaTomaLab = Val(mo_cmbResponsable.BoundText)
            .IdUsuarioAuditoria = ml_idUsuario
            .idPersonaRecoge = Val(mo_cmbPersonaRecoje.BoundText)
            .idDiagnostico = ml_IdDiagnostico
            .OrdenaPrueba = Trim(txtMedico.Text)
            If ml_IdDiagnostico > 0 Then
               .EsDiagnosticoDefinitivo = IIf(chkDxDefinitivo.Value = 1, sghTipoDx.sghTipoDxDefinitivo, sghTipoDx.sghTipoDxPresuntivo)    '1-definitivo, 2-presuntivo
            Else
               .EsDiagnosticoDefinitivo = sghTipoDx.sghTipoDxNINGUNO
            End If
            .Paciente = Trim(oDOPaciente.ApellidoPaterno) & " " & Trim(oDOPaciente.ApellidoMaterno) & " " & oDOPaciente.PrimerNombre
            .idTipoSexo = oDOPaciente.idTipoSexo
            .FechaNacimiento = oDOPaciente.FechaNacimiento
            If IsDate(Me.txtFum.Text) Then
                .eo_eg = Val(Me.txtEG.Text)
                .Eo_FUM = CDate(Me.txtFum.Text)
            End If
            .colegiatura = Me.txtColegiatura.Text
        End With
        With oDOPaciente  'ya lo cargo en Validacion de Datos
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoFactOrdenServ
            .FechaHoraRealizaCpt = txtFrealizaCpt.Text
        End With
    Case sghEliminar
        With oDOLabMovimiento
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoLabMovimientoLaboratorio
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDOPaciente  'ya lo cargo en Validacion de Datos
            .IdUsuarioAuditoria = ml_idUsuario
        End With
    End Select
End Sub

Function ValidarReglas() As Boolean
  ValidarReglas = False
    If sighentidades.EsFecha(txtFum, "DD/MM/AAAA") And UcPacienteDatos1.idTipoSexo = 1 Then
       MsgBox "El Paciente es MASCULINO no debe tener fecha de última mestruación (FUM)", vbInformation, Me.Caption
       Exit Function
    ElseIf IsDate(txtFum.Text) And Val(Me.txtEG.Text) <= 0 Then
       MsgBox "Existe FUM pero no exite EDAD GESTACIONAL (EG), chequee el FUM", vbInformation, Me.Caption
       Exit Function
    End If
    
  ValidarReglas = True
End Function

Function AgregarDatos() As Boolean
    AgregarDatos = mo_ReglasLaboratorio.LabMovimientoLaboratorioAgregar(oDOLabMovimiento, oDoLabMovimientoLaboratorio, _
                   oDOPaciente, oDoFactOrdenServ, rsProductos, Val(txtNcuenta.Text), rsProductosCPT, _
                   mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta, _
                   IIf(Trim(Me.txtNserie.Text) = "", "", Me.txtNserie.Text & "-" & Me.txtNboleta.Text), _
                   lcMedicoDNI, IIf(cmbMedicos.Text = "", Me.txtMedico.Text, cmbMedicos.Text), lcCama, lcDx, lcDxCodigo, _
                   lcUPS, lnEpsPorcentaje, Val(Me.txtNcita.Text))
    If mo_ReglasLaboratorio.IdOrdenPago > 0 Then
       lblOrdenPago.Caption = "N° Orden de Pago: " & mo_ReglasLaboratorio.IdOrdenPago
    End If
    ms_MensajeError = mo_ReglasLaboratorio.MensajeError
    ml_IdMovimiento = oDOLabMovimiento.IdMovimiento
    If oDoLabMovimientoLaboratorio.idCuentaAtencion > 0 Then
       mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar oDoLabMovimientoLaboratorio.idCuentaAtencion, False, 0
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios oDoLabMovimientoLaboratorio.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
    End If
End Function

Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasLaboratorio.LabMovimientoLaboratorioModificar(oDOLabMovimiento, oDoLabMovimientoLaboratorio, _
                     oDOPaciente, oDoFactOrdenServ, rsProductos, Val(txtNcuenta.Text), rsProductosCPT, _
                     mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta, _
                   IIf(Trim(Me.txtNserie.Text) = "", "", Me.txtNserie.Text & "-" & Me.txtNboleta.Text), _
                   lcMedicoDNI, IIf(cmbMedicos.Text = "", Me.txtMedico.Text, cmbMedicos.Text), lcCama, lcDx, lcDxCodigo, lcUPS, _
                   lnEpsPorcentaje, Val(lblOrdenPago.Tag))
    If mo_ReglasLaboratorio.IdOrdenPago > 0 Then
       lblOrdenPago.Caption = "N° Orden de Pago: " & mo_ReglasLaboratorio.IdOrdenPago
    End If
    ms_MensajeError = mo_ReglasLaboratorio.MensajeError
    If oDoLabMovimientoLaboratorio.idCuentaAtencion > 0 Then
       mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar oDoLabMovimientoLaboratorio.idCuentaAtencion, False, 0
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios oDoLabMovimientoLaboratorio.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
    End If
End Function

Function EliminarDatos() As Boolean
    Set rsProductosCPT = Me.ucProductos.FacturacionProductos
    EliminarDatos = mo_ReglasLaboratorio.LabMovimientoLaboratorioAnular(oDOLabMovimiento, oDoLabMovimientoLaboratorio, oDOPaciente, _
                               oDoFactOrdenServ, rsProductos, Val(txtNcuenta.Text), rsProductosCPT, mo_lnIdTablaLISTBARITEMS, _
                               mo_lcNombrePc, lnIdReceta, _
                   IIf(Trim(Me.txtNserie.Text) = "", "", Me.txtNserie.Text & "-" & Me.txtNboleta.Text), _
                   lcMedicoDNI, IIf(cmbMedicos.Text = "", Me.txtMedico.Text, cmbMedicos.Text), lcCama, lcDx, lcDxCodigo, lcUPS, _
                   Val(lblOrdenPago.Tag))
    ms_MensajeError = mo_ReglasLaboratorio.MensajeError
    If oDoLabMovimientoLaboratorio.idCuentaAtencion > 0 Then
       mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar oDoLabMovimientoLaboratorio.idCuentaAtencion, False, 0
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios oDoLabMovimientoLaboratorio.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
    End If
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub btnTamizaje_Click()

'<Agregado por: WABG el: 11/29/2020-12:39:16 en el equipo: SISGALENPLUS-PC><CAMBIO 44>
   Dim rs As Recordset
   Dim oConexion As New ADODB.Connection
   oConexion.Open sighentidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   Set rs = mo_ReglasLaboratorio.SeleccionarProcedimientosTamizajeSegunTipoFinanciamiento(ml_IdTipoFinanciamiento)
   ucProductos.CargarItemsALaGrillaPaquete rs
   oConexion.Close
   Set oConexion = Nothing
   btnTamizaje.Visible = False
'</Agregado por: WABG el: 11/29/2020-12:39:16 en el equipo: SISGALENPLUS-PC><CAMBIO 44>


End Sub

Private Sub chkPlanNoCubre_Click()
    If chkPlanNoCubre.Value = 1 Then
       ml_IdTipoFinanciamiento = 1
       cmbFormaPago.BoundText = ml_IdTipoFinanciamiento
       ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
       ucProductos.LimpiarGrilla
    Else
       txtNcuenta_LostFocus
       ucProductos.LimpiarGrilla
    End If
    ucProductos.TabEnDescripcion
End Sub

Private Sub cmbMedicos_Click()
  txtMedico.Text = cmbMedicos.Text
End Sub

Private Sub cmbMedicos_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbMedicos
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbMedicos_LostFocus()
  Me.txtColegiatura.Text = mo_ReglasDeProgMedica.MedicoDevuelveColegiatura(Val(mo_cmbMedicos.BoundText))
  ucProductos.TabEnDescripcion
End Sub

Private Sub cmbPersonaRecoje_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbPersonaRecoje
End Sub

Private Sub cmbResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbResponsable
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdBuscaDx_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.SoloMuestraDxGalenHos = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtDx.Text = oDODiagnostico.CodigoCIE2004
            txtNombreDx.Text = oDODiagnostico.descripcion
            lcDxCodigo = Trim(txtDx.Text)
            lcDx = txtNombreDx.Text
            
            chkDxDefinitivo.Visible = True
        End If
    End If
    Set oBusqueda = Nothing
    Set oDODiagnostico = Nothing

End Sub

Private Sub cmdBuscarPago_Click()
  frmCaja.Show vbModal
End Sub

'Private Sub Form_Activate()
'    If ml_SeEligioGridBoleta = True And mi_Opcion = sghAgregar And ml_IdMovimiento > 0 Then
'       ml_SeEligioGridBoleta = False
'
'       On Error Resume Next
'       cmbResponsable.SetFocus
'    End If
'
'End Sub

Private Sub Form_Activate()
    If ml_SeEligioGridBoleta = True And mi_Opcion = sghAgregar And ml_IdMovimiento > 0 Then
       ml_SeEligioGridBoleta = False
       On Error Resume Next
       cmbResponsable.SetFocus
    Else
'        ssoptPaciente.Value = True 'Actualizado 29092014
        
        Frame1.Enabled = True
        If mi_Opcion = sghAgregar Then
        lblDx.Caption = ""
        End If
        If ssoptPaciente.Value = True Or ssoptPaciente.Value = 1 Then
          fraHospitalizado.Visible = True
          lblDx.Visible = True
          fraExterno.Visible = False
          cmbMedicos.Visible = True
          txtMedico.Visible = False
          cmbBuscaReceta.Visible = True
          txtNreceta.Visible = True
          'ucProductos.LimpiarGrilla 'Actualizado 28102014 yamill palomino
        End If
        On Error Resume Next
        Me.txtNcuenta.SetFocus
    End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbResponsable.MiComboBox = cmbResponsable
    Set mo_cmbMedicos.MiComboBox = cmbMedicos
    Set mo_cmbPersonaRecoje.MiComboBox = cmbPersonaRecoje
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    wxParametro578 = lcBuscaParametro.SeleccionaFilaParametro(578)
    lblOrdenPago.Caption = ""
    txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL
    txtEstado.Text = "Registrado"
'    txtFrealizaCpt.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    txtFrealizaCpt.Text = lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
    
    CargaDataCombos
    cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, sighentidades.NombreUsuario)
    'If cmbResponsable.Text <> "" Then cmbResponsable.Enabled = False
    
    Me.ucProductos.HabilitaIngresoDePrecio = False
    Me.ucProductos.PermiteVerColumnaCantidadFallada = True
    Me.ucProductos.idUsuario = ml_idUsuario
    Me.ucProductos.Inicializar
    Me.ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    Me.ucProductos.TipoProducto = sghServicio
    Me.ucProductos.IdPuntoCarga = ml_puntoCarga
    
    Select Case mi_Opcion
    Case sghAgregar
        If ml_puntoCarga = 2 Or ml_puntoCarga = 31 Or ml_puntoCarga = 33 Or ml_puntoCarga = 34 Or ml_puntoCarga = 35 Or ml_puntoCarga = 36 Or ml_puntoCarga = 37 Then
          Me.Caption = "Agregar Órdenes Patología Clínica"
            ElseIf ml_puntoCarga = 3 Or ml_puntoCarga = 32 Then
          Me.Caption = "Agregar Órdenes Anatomía Patológica"
        ElseIf ml_puntoCarga = 11 Or ml_puntoCarga = 38 Then
          Me.Caption = "Agregar Órdenes Banco de Sangre"
        End If
    Case sghModificar
        If ml_puntoCarga = 2 Or ml_puntoCarga = 31 Or ml_puntoCarga = 33 Or ml_puntoCarga = 34 Or ml_puntoCarga = 35 Or ml_puntoCarga = 36 Or ml_puntoCarga = 37 Then
          Me.Caption = "Modificar Órdenes Patología Clínica"
        ElseIf ml_puntoCarga = 3 Or ml_puntoCarga = 32 Then
          Me.Caption = "Modificar Órdenes Anatomía Patológica"
        ElseIf ml_puntoCarga = 11 Or ml_puntoCarga = 38 Then
          Me.Caption = "Modificar Órdenes Banco de Sangre"
        End If
    Case sghConsultar
        If ml_puntoCarga = 2 Or ml_puntoCarga = 31 Or ml_puntoCarga = 33 Or ml_puntoCarga = 34 Or ml_puntoCarga = 35 Or ml_puntoCarga = 36 Or ml_puntoCarga = 37 Then
          Me.Caption = "Consultar Órdenes Patología Clínica"
        ElseIf ml_puntoCarga = 3 Or ml_puntoCarga = 32 Then
          Me.Caption = "Consultar Órdenes Anatomía Patológica"
        ElseIf ml_puntoCarga = 11 Or ml_puntoCarga = 38 Then
          Me.Caption = "Consultar Órdenes Banco de Sangre"
        End If
    Case sghEliminar
        If ml_puntoCarga = 2 Or ml_puntoCarga = 31 Or ml_puntoCarga = 33 Or ml_puntoCarga = 34 Or ml_puntoCarga = 35 Or ml_puntoCarga = 36 Or ml_puntoCarga = 37 Then
          Me.Caption = "Eliminar Órdenes Patología Clínica"
        ElseIf ml_puntoCarga = 3 Or ml_puntoCarga = 32 Then
          Me.Caption = "Eliminar Órdenes Anatomía Patológica"
        ElseIf ml_puntoCarga = 11 Or ml_puntoCarga = 38 Then
          Me.Caption = "Eliminar Órdenes Banco de Sangre"
        End If
    End Select
    CargarDatosAlFormulario
End Sub

Sub CargaBoletaAutomaticamente()
    If ml_SeEligioGridBoleta = True And ml_IdMovimiento > 0 Then
        Dim oRsTmp1 As New Recordset
        Set oRsTmp1 = mo_AdminCaja.CajaComprobantesSeleccionarPorId(ml_IdMovimiento)
        If oRsTmp1.RecordCount > 0 Then
            ssoptExterno.Value = True
            ssoptExterno_Click 1
            Me.txtNserie.Text = Trim(oRsTmp1.Fields!NroSerie)
            Me.txtNboleta.Text = Trim(oRsTmp1.Fields!nroDocumento)
            txtNboleta_LostFocus
        End If
        Set oRsTmp1 = Nothing
    End If
End Sub

Sub CargarDatosAlFormulario()
 mo_Formulario.HabilitarDeshabilitar Me.txtFum, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNmovimiento, False
 mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
 mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
 mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
 mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNroOrden, False
 mo_Formulario.HabilitarDeshabilitar Me.txtProcedencia, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNombreDx, False
 mo_Formulario.HabilitarDeshabilitar cmbFormaPago, False
 wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
 wxParametro509 = lcBuscaParametro.SeleccionaFilaParametro(509)
 Me.UcPacienteDatos1.Inicializar

 Select Case mi_Opcion
     Case sghAgregar
        Me.ucProductos.idOrden = 0
        Me.ucProductos.CargaProductosPorIdOrden
        'Me.ucProductos.AgregaProducto
        CargaBoletaAutomaticamente
     Case sghModificar
        CargarDatosALosControles
     Case sghConsultar
        CargarDatosALosControles
     Case sghEliminar
        CargarDatosALosControles
 End Select
End Sub

Sub CargarDatosALosControles()
  On Error GoTo Fin
  btnImprimir.Visible = True
  mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
  mo_Formulario.HabilitarDeshabilitar Me.txtNserie, False
  mo_Formulario.HabilitarDeshabilitar Me.txtNboleta, False
  cmdBuscaCuentaPorApellidos.Enabled = False
  Me.chkPlanNoCubre.Visible = False: txtDatosDeCuenta.Width = txtDatosDeCuenta.Width + chkPlanNoCubre.Width + 190
        
  'Carga datos de la orden
  Dim oRsTmp As New Recordset
  Dim lbSigue As Boolean
  Dim oConexion As New Connection
  Dim oFactOrdenServicio As New FactOrdenServicio
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Set oRsTmp = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorIdMovimiento(ml_IdMovimiento)
  If oRsTmp.RecordCount > 0 Then
    With oDOLabMovimiento
      .IdMovimiento = ml_IdMovimiento
      .fecha = oRsTmp.Fields!fecha
      .IdlabEstado = oRsTmp.Fields!IdlabEstado
      .IdPuntoCarga = oRsTmp.Fields!IdPuntoCarga
      .idTipoConcepto = oRsTmp.Fields!idTipoConcepto
      .MovTipo = oRsTmp.Fields!MovTipo
      .idUsuario = oRsTmp.Fields!idUsuario
    End With
    With oDoLabMovimientoLaboratorio
      .IdMovimiento = ml_IdMovimiento
      .CorrelativoAnual = IIf(IsNull(oRsTmp.Fields!CorrelativoAnual), 0, oRsTmp.Fields!CorrelativoAnual)
      .IdComprobantePago = IIf(IsNull(oRsTmp.Fields!IdComprobantePago), 0, oRsTmp.Fields!IdComprobantePago)
      .idCuentaAtencion = IIf(IsNull(oRsTmp.Fields!idCuentaAtencion), 0, oRsTmp.Fields!idCuentaAtencion)
      .idOrden = oRsTmp.Fields!idOrden
      .IdPersonaTomaLab = IIf(IsNull(oRsTmp.Fields!IdPersonaTomaLab), 0, oRsTmp.Fields!IdPersonaTomaLab)
      .idPersonaRecoge = IIf(IsNull(oRsTmp.Fields!idPersonaRecoge), 0, oRsTmp.Fields!idPersonaRecoge)
      .idDiagnostico = IIf(IsNull(oRsTmp!idDiagnostico), 0, oRsTmp!idDiagnostico)
      .EsDiagnosticoDefinitivo = IIf(IsNull(oRsTmp!EsDiagnosticoDefinitivo), 0, oRsTmp!EsDiagnosticoDefinitivo)
      .OrdenaPrueba = IIf(IsNull(oRsTmp!OrdenaPrueba), 0, oRsTmp!OrdenaPrueba)
      If .idCuentaAtencion > 0 Then
         Dim lnFor As Integer
         For lnFor = 1 To cmbMedicos.ListCount
             If Trim(Me.cmbMedicos.List(lnFor - 1)) = Trim(.OrdenaPrueba) Then
                Me.cmbMedicos.ListIndex = lnFor - 1
                Exit For
             End If
         Next
      End If
      If Not IsNull(oRsTmp!Eo_FUM) Then
            Me.txtEG.Text = IIf(IsNull(oRsTmp!eo_eg), "", oRsTmp!eo_eg)
            Me.txtFum.Text = Format(oRsTmp!Eo_FUM, sighentidades.DevuelveFechaSoloFormato_DMY)
      End If
      .colegiatura = IIf(IsNull(oRsTmp!colegiatura), "", oRsTmp!colegiatura)
      Me.txtColegiatura.Text = .colegiatura
    End With
    '
    oDoFactOrdenServ.idOrden = oDoLabMovimientoLaboratorio.idOrden
    Set oFactOrdenServicio.Conexion = oConexion
    If oFactOrdenServicio.SeleccionarPorId(oDoFactOrdenServ) Then
       Me.txtFrealizaCpt.Text = Format(oDoFactOrdenServ.FechaHoraRealizaCpt, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
    End If
    '
    
    txtFregistro.Text = Format(oDOLabMovimiento.fecha, sighentidades.DevuelveFechaSoloFormato_DMY)
    txtEstado.Text = oRsTmp.Fields!destado
    txtNmovimiento.Text = ml_IdMovimiento
    txtNcuenta.Text = oDoLabMovimientoLaboratorio.idCuentaAtencion
    txtNroOrden.Text = oDoLabMovimientoLaboratorio.idOrden
    mo_cmbPersonaRecoje.BoundText = oDoLabMovimientoLaboratorio.idPersonaRecoge
    If Val(txtNcuenta.Text) <> 0 Then
      ssoptPaciente.Value = True
      ssoptExterno.Value = False
'      cmbMedicos.ListIndex = Ubica_En_Combo(cmbMedicos, oDoLabMovimientoLaboratorio.OrdenaPrueba)
    Else
      ssoptExterno.Value = True
      ssoptPaciente.Value = False
      txtMedico.Text = oDoLabMovimientoLaboratorio.OrdenaPrueba
    End If
    ssoptPaciente.Enabled = False
    ssoptExterno.Enabled = False
    
    'Dx
    Dim mo_Diagnostico As New DODiagnostico
    ml_IdDiagnostico = oDoLabMovimientoLaboratorio.idDiagnostico
    If ml_IdDiagnostico > 0 Then
      Set mo_Diagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(ml_IdDiagnostico)
      txtDx.Text = mo_Diagnostico.CodigoCIE2004
      txtNombreDx.Text = mo_Diagnostico.descripcion
      Me.lblDx.Caption = txtDx.Text & " " & txtNombreDx.Text
      Me.lblDx.Visible = True
      lcDxCodigo = Trim(txtDx.Text)
      lcDx = txtNombreDx.Text
      '
      chkDxDefinitivo.Visible = True
      chkDxDefinitivo.Value = IIf(oDoLabMovimientoLaboratorio.EsDiagnosticoDefinitivo = 1, 1, 0)
    End If
    '
    mo_cmbResponsable.BoundText = oDoLabMovimientoLaboratorio.IdPersonaTomaLab
    ml_IdServicioPaciente = IIf(IsNull(oRsTmp.Fields!IdServicioPaciente), 0, oRsTmp.Fields!IdServicioPaciente)
    ml_idPaciente = IIf(IsNull(oRsTmp.Fields!idPaciente), 0, oRsTmp.Fields!idPaciente)
    ml_IdFuenteFinanciamiento = IIf(IsNull(oRsTmp.Fields!IdFuenteFinanciamiento), 0, oRsTmp.Fields!IdFuenteFinanciamiento)
    ml_IdTipoFinanciamiento = oRsTmp.Fields!IdTipoFinanciamiento
    Set grdConsumoPaciente.DataSource = mo_ReglasLaboratorio.CptHistoricosPorPaciente(ml_idPaciente, ml_IdMovimiento)
    '
    UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
    If ml_idPaciente = 0 Then
        If Not IsNull(oRsTmp.Fields!FechaNacimiento) Then
           UcPacienteDatos1.FechaNacimiento = oRsTmp.Fields!FechaNacimiento
        End If
        If Not IsNull(oRsTmp.Fields!idTipoSexo) Then
           UcPacienteDatos1.idTipoSexo = oRsTmp.Fields!idTipoSexo
        End If
        UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta oRsTmp.Fields!Paciente
    Else
        UcPacienteDatos1.idPaciente = ml_idPaciente
        UcPacienteDatos1.CargarDatosDePacienteALosControles
    End If
    '
    If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
      Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
      lbSigue = True
      If oRsTmp.RecordCount > 0 Then
        If oRsTmp.Fields!idEstado <> 1 Then
          If mi_Opcion <> sghConsultar Then
            MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
            If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
              btnAceptar.Enabled = False
            Else
              lbSigue = False
            End If
          End If
        End If
        If lbSigue Then
            lnEpsPorcentaje = mo_ReporteUtil.DevuelveEpsPorcentaje(oRsTmp!EpsPorcentaje)
            mo_Formulario.HabilitarAlerta txtPlan, IIf(lnEpsPorcentaje > 0, True, False)
            If lnEpsPorcentaje > 0 Then
               Dim lcBoletaEPS As String
               lblOrdenPago.Tag = mo_ReglasFacturacion.DevuelveOrdenPago(oRsTmp!IdAtencion, sghPtoCargaCaja, oDOLabMovimiento.fecha, oConexion, lcBoletaEPS)
               lblOrdenPago.Caption = "N° Orden de Pago: " & lblOrdenPago.Tag
               If lcBoletaEPS <> "" Then
                    lblOrdenPago.Caption = lcBoletaEPS
                    MsgBox "El SEGURO tiene EPS, No podrá MODIFICAR/ELIMINAR porque ya pagó en CAJA" & Chr(13) & _
                           "Tendría que ANULAR (o NOTA DE CREDITO) la BOLETA para usar MODIFICAR/ELIMINAR", vbInformation, ""
                    Me.btnAceptar.Enabled = False
               End If
               
            End If
            lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
            txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!fechaingreso & " - " & IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
            txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento & mo_ReporteUtil.DevuelveEPScubre(lnEpsPorcentaje)
            'debb-14/04/2011
            If mi_Opcion = sghModificar And oRsTmp.Fields!IdFuenteFinanciamiento <> ml_IdFuenteFinanciamiento Then
               MsgBox "No se podrá modificar datos, porque el despacho tubo otro PRODUCTO/PLAN" & Chr(13) & "hubo RECALCULO", vbInformation, Me.Caption
               btnAceptar.Enabled = False
            End If
        End If
      End If
      '
      Set oRsTmp = mo_reglasComunes.RecetaCabeceraFiltraXcuentaYDocumentodespacho(txtNmovimiento.Text, Val(txtNcuenta.Text))
      lnIdReceta = 0
      ucProductos.PermiteAgregarItems = True
      If oRsTmp.RecordCount > 0 Then
         lnIdReceta = oRsTmp.Fields!IdReceta
         ucProductos.PermiteAgregarItems = False
      End If
      '
      
      BuscaServicioActualDelPaciente
      UcPacienteDatos1.DeshabilitarFrames True
    Else
      Dim oDOCajaComprobantesPago As New DOCajaComprobantesPago
      Set oDOCajaComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(oRsTmp.Fields!IdComprobantePago, oConexion)
      txtNserie.Text = oDOCajaComprobantesPago.NroSerie
      txtNboleta.Text = oDOCajaComprobantesPago.nroDocumento
      ucProductos.PermiteAgregarItems = False
      UcPacienteDatos1.DeshabilitarFrames False
    End If
'    UcPacienteDatos1.idPaciente = ml_idPaciente
'    UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
'    UcPacienteDatos1.CargarDatosDePacienteALosControles
    If oDOLabMovimiento.IdlabEstado = 0 Or mi_Opcion = sghConsultar Then btnAceptar.Enabled = False
    cmbFormaPago.BoundText = ml_IdTipoFinanciamiento
    mb_ExistenDatos = True
  Else
    mb_ExistenDatos = False
    Exit Sub
  End If
  
  'Cargar datos de los servicios
  Me.ucProductos.LimpiarGrilla
  Me.ucProductos.IdMovimiento = ml_IdMovimiento
  Me.ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
  Me.ucProductos.CargaProductosPorIdMovimiento
  Me.ucProductos.CargaObservacionesDeReceta lnIdReceta, oConexion
  
  Select Case mi_Opcion
    Case sghModificar
    Case sghEliminar
    Case sghConsultar
  End Select
  oConexion.Close
  Set oFactOrdenServicio = Nothing
  Set oConexion = Nothing
  Exit Sub
Fin:
  MsgBox Err.Description
 
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
  Dim oBusqueda As New SIGHNegocios.BuscaPacientes
  Dim oDOPaciente As New doPaciente
  Dim oConexion As New Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  oBusqueda.TipoFiltro = sghFiltrarTodos
  oBusqueda.MostrarFormulario
  If oBusqueda.BotonPresionado = sghAceptar Then
    Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
    If Not oDOPaciente Is Nothing Then
      ml_idPaciente = oDOPaciente.idPaciente
      Dim oRsTmp As New Recordset
      Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_idPaciente, oConexion, True)
      If oRsTmp.RecordCount > 0 Then txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
      oRsTmp.MoveFirst
      oRsTmp.Close
      Set oRsTmp = Nothing
      txtNcuenta_LostFocus
    End If
  End If
  oConexion.Close
  Set oConexion = Nothing
End Sub

Private Sub ssoptTipo_Click(Index As Integer, Value As Integer)
  Frame1.Enabled = True
  If Index = 0 Then
    
  ElseIf Index = 1 Then
    
  Else
    fraHospitalizado.Visible = False
    fraExterno.Visible = False
  End If
End Sub

Private Sub grdConsumoPaciente_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
     grdConsumoPaciente.Bands(0).Columns("Fecha").Width = 1500
     grdConsumoPaciente.Bands(0).Columns("idMovimiento").Width = 1000
     grdConsumoPaciente.Bands(0).Columns("Codigo").Width = 1000
     grdConsumoPaciente.Bands(0).Columns("Nombre").Width = 5000
     grdConsumoPaciente.Bands(0).Columns("Cantidad").Width = 500
End Sub

Private Sub ssoptExterno_Click(Value As Integer)
  Frame1.Enabled = True
  If Value = True Or Value = 1 Then
    fraExterno.Visible = True
    fraHospitalizado.Visible = False
    lblDx.Visible = False
    On Error Resume Next
    If txtNserie.Locked = False Then
      txtNserie.SetFocus
    Else
      'cmbResponsable.SetFocus
    End If
    txtMedico.Visible = True
    cmbMedicos.Visible = False
    cmbBuscaReceta.Visible = False
    txtNreceta.Visible = False
    lnEpsPorcentaje = 0
  End If
End Sub

Private Sub ssoptPaciente_Click(Value As Integer)
  Frame1.Enabled = True
  lblDx.Caption = ""
  If Value = True Or Value = 1 Then
    fraHospitalizado.Visible = True
    lblDx.Visible = True
    fraExterno.Visible = False
    If txtNcuenta.Locked = False Then
      txtNcuenta.SetFocus
    Else
      'cmbResponsable.SetFocus
    End If
    cmbMedicos.Visible = True
    txtMedico.Visible = False
    cmbBuscaReceta.Visible = True
    txtNreceta.Visible = True
  End If
End Sub







Private Sub txtEG_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtEG
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFrealizaCpt_LostFocus()
If Not IsDate(txtFrealizaCpt.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFrealizaCpt.Text = sighentidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub

Private Sub txtMedico_GotFocus()
  SeleccionaTexto txtMedico
End Sub

Private Sub txtMedico_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtMedico
End Sub

Private Sub txtMedico_LostFocus()
  ucProductos.TabEnDescripcion
End Sub

Private Sub txtNboleta_GotFocus()
  SeleccionaTexto txtNboleta
End Sub

Private Sub txtNboleta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNboleta
End Sub

Private Sub txtNboleta_KeyPress(KeyAscii As Integer)
  If Not (mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Or KeyAscii = 13 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtNboleta_LostFocus()
  If txtNboleta.Locked = True Then Exit Sub
  On Error Resume Next
'  If mo_Teclado.TextoEsSoloNumeros(txtNserie.Text) And mo_Teclado.TextoEsSoloNumeros(txtNboleta.Text) Then
  If mo_Teclado.TextoEsSoloNumeros(txtNboleta.Text) Then
    Dim rsBuscaBoleta As New Recordset
    Dim rsBuscaBoletaEnLaboratorio As New Recordset
    Dim oRsTmp1 As New Recordset, lcSql As String, lnIdCuenta111 As Long
    Dim oConexion As New Connection, lnIdPacienteBoleta As Long
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set rsBuscaBoleta = mo_AdminCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion(txtNserie.Text, Trim(txtNboleta.Text), oConexion)
    If rsBuscaBoleta.RecordCount > 0 Then
        lnIdPacienteBoleta = 0
        If rsBuscaBoleta.Fields!idPaciente > 0 Then
           lnIdPacienteBoleta = rsBuscaBoleta.Fields!idPaciente
           '
           CargaFUMdeHistoricos lnIdPacienteBoleta, oConexion
           '
        End If
        If rsBuscaBoleta.Fields!idEstadoComprobante <> sghEstadosComprobante.sighEstadosComprobantePagado Then
            MsgBox "Esa Boleta está ANULADA", vbInformation, Me.Caption
            txtNboleta.Text = ""
            txtNserie.Text = ""
            txtNserie.SetFocus
            ml_IdComprobantePago = 0
        Else
            Set rsBuscaBoletaEnLaboratorio = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorIdComprobantePago(rsBuscaBoleta.Fields!IdComprobantePago)
            If rsBuscaBoletaEnLaboratorio.RecordCount > 0 Then
                MsgBox "Esta Boleta ya fué DESPACHADA con Movimiento N° " & rsBuscaBoletaEnLaboratorio.Fields!IdMovimiento & " de fecha " & rsBuscaBoletaEnLaboratorio.Fields!fecha, vbInformation, Me.Caption
                txtNboleta.Text = ""
                txtNserie.Text = ""
                txtNserie.SetFocus
                ml_IdComprobantePago = 0
            Else
                ml_IdComprobantePago = rsBuscaBoleta.Fields!IdComprobantePago
                ml_IdTipoFinanciamiento = 1 'Contado
                ml_IdFuenteFinanciamiento = 1   'Contado
                txtNroOrden.Text = rsBuscaBoleta.Fields!idOrden
                If rsBuscaBoleta.Fields!idPaciente > 0 And rsBuscaBoleta.Fields!idCuentaAtencion > 0 Then
                   'Paciente contado, con cuenta (CE), pago en CAJA
                   lnIdCuenta111 = rsBuscaBoleta.Fields!idCuentaAtencion
                   ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(rsBuscaBoleta.Fields!idCuentaAtencion, CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
                   txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
                   UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                   UcPacienteDatos1.idPaciente = rsBuscaBoleta.Fields!idPaciente
                   UcPacienteDatos1.CargarDatosDePacienteALosControles
                   UcPacienteDatos1.DeshabilitarFrames True
                ElseIf rsBuscaBoleta.Fields!idPaciente > 0 Then
                   'Paciente con Nro Historia
                   UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                   UcPacienteDatos1.idPaciente = rsBuscaBoleta.Fields!idPaciente
                   UcPacienteDatos1.CargarDatosDePacienteALosControles
                   UcPacienteDatos1.DeshabilitarFrames True
                Else
                   'Paciente contado, EXTERNO
                   UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta (rsBuscaBoleta.Fields!razonSocial)
                   UcPacienteDatos1.DeshabilitarFrames False
                   UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                End If
                ucProductos.NoPermiteCargarCantidadFallada = True
                ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
                ucProductos.PermiteAgregarItems = False
                ucProductos.LimpiarGrilla
                ucProductos.CargarItemsALaGrillaCPT rsBuscaBoleta, True
                txtNcuenta.Text = ""
                txtDatosDeCuenta.Text = ""
                txtPlan.Text = ""
                '
                
                txtMedico.Text = ""
                If lnIdCuenta111 > 0 And oRsTmp1.Fields!idTipoServicio <> sghTipoServicio.sghHospitalizacion Then
                    Set oRsTmp1 = mo_reglasComunes.AtencionesSeleccionarMedicoPorCuenta(lnIdCuenta111)
                    If oRsTmp1.RecordCount > 0 Then
                       mo_cmbMedicos.BoundText = oRsTmp1.Fields!idEmpleado
                       Me.txtMedico.Text = cmbMedicos.Text
                       'mo_cmbMedicos.BoundText = ""
                    End If
                    oRsTmp1.Close
                End If
                '
                If lnIdPacienteBoleta > 0 And mi_Opcion = sghAgregar Then
                   Set grdConsumoPaciente.DataSource = mo_ReglasLaboratorio.CptHistoricosPorPaciente(lnIdPacienteBoleta, 0)
                End If
                '
                cmbResponsable.SetFocus
            End If
      End If
    End If
    Set rsBuscaBoleta = Nothing
    Set rsBuscaBoletaEnLaboratorio = Nothing
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
  End If
End Sub

Private Sub txtNcuenta_GotFocus()
  '
End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) = False Then
     cmdBuscaCuentaPorApellidos_Click
   Else
'    lblDx.Caption = mo_ReglasLaboratorio.SeleccionaDiagnosticoDeAtencion(txtNcuenta.Text)
'    lcDxCodigo = Trim(Left(lblDx.Caption, InStr("-->", lblDx.Caption) - 1))
'    lcDx = Trim(Mid(lblDx.Caption, InStr("-->", lblDx.Caption) + 3))
    If cmbResponsable.Enabled = True Then
      cmbResponsable.SetFocus
    Else
      If txtMedico.Visible = True Then
        txtMedico.SetFocus
      Else
        cmbMedicos.SetFocus
      End If
    End If
   End If
  End If
End Sub

Private Sub txtNcuenta_LostFocus()
  If txtNcuenta.Locked = True Then Exit Sub
  On Error Resume Next
  If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim lbSigue As Boolean
    Dim oConexion As New Connection
    lcMedicoDNI = "": lcCama = "": lcDxCodigo = "": lcDx = "": lcUPS = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
    lbSigue = True
    If oRsTmp.RecordCount > 0 Then
      If oRsTmp.Fields!idEstado <> 1 Then
        If mi_Opcion <> sghConsultar And lbCuentaDeEmergenciaCerrada = False Then
          MsgBox "Esta Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
          'txtNcuenta.SetFocus
          If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
            btnAceptar.Enabled = False
          Else
            lbSigue = False
          End If
        End If
      End If
      '
      If lbSigue = True And oRsTmp!esPacienteExterno <> True And wxParametro509 = "S" And mi_Opcion = sghAgregar Then
           If Val(txtNreceta.Text) = 0 Then
              MsgBox "No puede usar N° CUENTA, tiene que generar RECETA", vbInformation, Me.Caption
              lbSigue = False
           End If
      End If
      '
      
      If mi_Opcion = sghAgregar And _
           mo_AdminAdmision.AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS(Val(txtNcuenta.Text), wxParametro302, _
                                                                        oRsTmp.Fields!IdFuenteFinanciamiento) = False Then
                                                                     
           lbSigue = False
      End If
      
      If mi_Opcion = sghAgregar And oRsTmp.Fields!idTipoServicio = sghTipoServicio.sghConsultaExterna _
                                                                              And oRsTmp.Fields!IdFormaPago = 1 Then
              MsgBox "Es un Paciente PAGANTE y viene por CONSULTORIO EXTERNO" & Chr(13) & _
                      "Debe pagar antes en CAJA", vbInformation, "Laboratorio"
              lbSigue = False
      End If
      
      If mi_Opcion = sghAgregar And _
                              mo_AdminAdmision.LaFechaDespachoEsMenorAfechaCita(CDate(Format(oRsTmp!fechaingreso, _
                              sighentidades.DevuelveFechaSoloFormato_DMY) & " " & oRsTmp!horaIngreso)) = True Then
         lbSigue = False
      End If
          
      If lbSigue Then
        lnEpsPorcentaje = mo_ReporteUtil.DevuelveEpsPorcentaje(oRsTmp!EpsPorcentaje)
        mo_Formulario.HabilitarAlerta txtPlan, IIf(lnEpsPorcentaje > 0, True, False)
        lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
        txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!fechaingreso & " - " & _
                    IIf(oRsTmp!esPacienteExterno = True, "Externo", _
                    IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", _
                    IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia"))) & _
                    " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
        txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento & mo_ReporteUtil.DevuelveEPScubre(lnEpsPorcentaje)
        ml_idPaciente = oRsTmp.Fields!idPaciente
        ml_IdFuenteFinanciamiento = oRsTmp.Fields!IdFuenteFinanciamiento
        ml_IdTipoFinanciamiento = oRsTmp.Fields!IdFormaPago
        UcPacienteDatos1.idPaciente = ml_idPaciente
        UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
        UcPacienteDatos1.CargarDatosDePacienteALosControles
        UcPacienteDatos1.DeshabilitarFrames True
        ucProductos.IdTipoFinanciamiento = oRsTmp.Fields!IdFormaPago
        ucProductos.PermiteAgregarItems = True
        ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(Val(txtNcuenta.Text), CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
        BuscaServicioActualDelPaciente
        txtNserie.Text = ""
        txtNboleta.Text = ""
        txtNroOrden.Text = ""
        ml_IdComprobantePago = 0
        If cmbResponsable.Enabled = True Then cmbResponsable.SetFocus
        ucProductos.idCuentaAtencion = txtNcuenta.Text
        cmbFormaPago.BoundText = ml_IdTipoFinanciamiento
        '
        CargaFUMdeHistoricos ml_idPaciente, oConexion
        '
        lblDx.Caption = mo_ReglasLaboratorio.SeleccionaDiagnosticoDeAtencion(txtNcuenta.Text)
        If Len(lblDx) > 0 Then
            lcDxCodigo = Trim(Mid(lblDx.Caption, 13, InStr(lblDx.Caption, "-->") - 13))
            lcDx = Trim(Mid(lblDx.Caption, InStr(lblDx.Caption, "-->") + 3))
            
            Set oRsTmp2 = mo_reglasComunes.DiagnosticosSeleccionarXCodigo(lcDxCodigo)
            If oRsTmp2.RecordCount > 0 Then
               ml_IdDiagnostico = oRsTmp2!idDiagnostico
            End If
            oRsTmp2.Close
        ElseIf wxParametro578 = "S" Then
             MsgBox "No se puede despachar sino tiene Dx", vbInformation, ""
             txtDatosDeCuenta.Text = ""
             txtPlan.Text = ""
        End If
        '
        
        If oRsTmp.Fields!idTipoServicio <> sghTipoServicio.sghHospitalizacion Then
           mo_cmbMedicos.BoundText = oRsTmp.Fields!idEmpleado
           
           
        End If
        If mo_cmbMedicos.BoundText <> "" Then
           Me.txtColegiatura.Text = mo_ReglasDeProgMedica.MedicoDevuelveColegiatura(Val(mo_cmbMedicos.BoundText))
        End If
        '
        If mi_Opcion = sghAgregar Then
           Set grdConsumoPaciente.DataSource = mo_ReglasLaboratorio.CptHistoricosPorPaciente(ml_idPaciente, 0)
        End If
     Else
         txtNreceta.Text = ""
        
      End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oRsTmp2 = Nothing
'<Agregado por: WABG el: 11/30/2020-13:32:44 en el equipo: SISGALENPLUS-PC><CAMBIO 44>
     If Me.Caption = "Agregar Órdenes Banco de Sangre" Then
     btnTamizaje.Visible = True
     End If
'</Agregado por: WABG el: 11/30/2020-13:32:44 en el equipo: SISGALENPLUS-PC><CAMBIO 44>
  End If
End Sub

Sub BuscaServicioActualDelPaciente()
  Dim oBuscaNombreServicioPaciente As New SIGHNegocios.ReglasServiciosHosp
  Dim DOServicio As New DOServicio
  Dim oConexion As New Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Set DOServicio = oBuscaNombreServicioPaciente.ServiciosSeleccionarPorId(ml_IdServicioPaciente, oConexion)
  txtProcedencia.Text = DOServicio.Nombre
  lcUPS = DOServicio.codigoServicioHIS
  oConexion.Close
  Set oConexion = Nothing
  Set DOServicio = Nothing
  Set oBuscaNombreServicioPaciente = Nothing
End Sub

Private Sub txtNserie_GotFocus()
  'LimpiarFormulario
End Sub

Private Sub txtNserie_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtNserie
End Sub

Private Sub txtNserie_KeyPress(KeyAscii As Integer)
'  If Not (mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Or KeyAscii = 13 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Sub CargaDataCombos()
    '
    mo_cmbResponsable.BoundColumn = "idEmpleado"
    mo_cmbResponsable.ListField = "ApNom"
    Select Case ml_puntoCarga
    Case 2 'Patologia Clinica
       Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =10")
    Case 3  'Anatomia Patologica
       Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =11")
    Case 11   'Banco de Sangre
       Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =19")
    End Select
    'debb-09/08/2016
    If mo_reglasComunes.NOpuedeModificarResponsable(mi_Opcion, ml_idUsuario, mo_cmbResponsable.RowSource) Then
       If mi_Opcion = sghAgregar Then
          mo_cmbResponsable.BoundText = Trim(Str(ml_idUsuario))
       End If
       mo_Formulario.HabilitarDeshabilitar Me.cmbResponsable, False
    End If
    '
    mo_cmbPersonaRecoje.BoundColumn = "idRecojeExamen"
    mo_cmbPersonaRecoje.ListField = "RecojeExamen"
    Set mo_cmbPersonaRecoje.RowSource = mo_ReglasLaboratorio.LabRecojeExamenSeleccionarTodos
    mo_cmbMedicos.BoundColumn = "idEmpleado"
    mo_cmbMedicos.ListField = "ApNom"
    Set mo_cmbMedicos.RowSource = mo_ReglasLaboratorio.LabSeleccionaMedicos
    '
    Set oRsFormaPago = mo_ReglasFacturacion.TiposFinanciamientoSoloFarmacia
    Set cmbFormaPago.RowSource = oRsFormaPago
    cmbFormaPago.ListField = "Descripcion"
    cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
End Sub

Private Sub txtDx_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtDx
End Sub

Private Sub txtDx_LostFocus()
  Dim oDODiagnostico As DODiagnostico
  Set oDODiagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorCodigoCIE2004(txtDx.Text, True)
  If Not oDODiagnostico Is Nothing Then
    ml_IdDiagnostico = oDODiagnostico.idDiagnostico
    txtNombreDx.Text = oDODiagnostico.descripcion
    chkDxDefinitivo.Visible = True
  Else
    ml_IdDiagnostico = 0
    txtNombreDx.Text = ""
    chkDxDefinitivo.Visible = False
  End If
End Sub

Private Sub txtResultadoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtResultadoFinal
End Sub

Private Sub txtZonaCuerpo_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtZonaCuerpo
End Sub



Private Sub UcPacienteDatos1_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
     End If

End Sub

Private Sub ucProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
     End If
End Sub

Private Sub txtNreceta_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNreceta
       AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNreceta_LostFocus()
    If Val(txtNreceta.Text) > 0 Then
       Dim lcSql As String
       Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset
       Dim lnRecetaProcesada As Long, lnCuenta As Long
       lnRecetaProcesada = Val(txtNreceta.Text)
       '
       ucProductos.LimpiarGrilla
       Set oRsTmp1 = mo_reglasComunes.RecetasConCabeceraYdetalleSoloCpt(lnRecetaProcesada, sghRecetaEstados.sighRecetaRegistrada)
       If oRsTmp1.RecordCount > 0 Then
            If oRsTmp1.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                mo_reglasComunes.RecetaChequeaEstadoActual oRsTmp1.Fields!idCuentaAtencion, _
                                                           oRsTmp1.Fields!idEstado, _
                                                           0, oRsTmp1.Fields!DocumentoDespacho
                txtNreceta.Text = ""
            Else
                If oRsTmp1.Fields!IdPuntoCarga <> sghPtoCargaPatologiaClinica And _
                            oRsTmp1.Fields!IdPuntoCarga <> sghPtoCargaAnatomiaPatologica1 And _
                            oRsTmp1.Fields!IdPuntoCarga <> sghPtoCargaAnatomiaPatologica2 And _
                            oRsTmp1.Fields!IdPuntoCarga <> sghPtoCargaBancoSangre1 And _
                            oRsTmp1.Fields!IdPuntoCarga <> sghPtoCargaBancoSangre2 Then
                     MsgBox "Esa receta no es de LABORATORIO", vbInformation, "Imágenes"
                     txtNreceta.Text = ""
                Else
                     lbCuentaDeEmergenciaCerrada = mo_reglasComunes.CuentaDeEmergenciaCerrada(oRsTmp1!idCuentaAtencion, ml_puntoCarga)
                     If Not IsNull(oRsTmp1.Fields!idMedicoReceta) Then
                       Set oRsTmp2 = mo_ReglasDeProgMedica.MedicosSeleccionarXId(oRsTmp1.Fields!idMedicoReceta)
                       If oRsTmp2.RecordCount > 0 Then
                          mo_cmbMedicos.BoundText = oRsTmp2!idEmpleado
                          
                       End If
                       oRsTmp2.Close
                     End If
                     txtNcuenta.Text = oRsTmp1.Fields!idCuentaAtencion
                     txtNcuenta_LostFocus
                     ucProductos.CargaProductosPorIdReceta oRsTmp1
                     ucProductos.PermiteAgregarItems = False
                     lnIdReceta = lnRecetaProcesada
                     On Error Resume Next

                     Me.cmbResponsable.SetFocus
                End If
            End If
       Else
            MsgBox "Ese N° Receta NO EXISTE", vbInformation, "Caja"
            txtNreceta.Text = ""
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
    End If
End Sub


Private Sub cmbBuscaReceta_Click()
    Dim oBusqueda As New SIGHNegocios.clBuscaReceta
    oBusqueda.IdPuntoCarga = ml_puntoCarga
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       txtNreceta.Text = oBusqueda.IdRecetaSeleccionada
       txtNreceta_LostFocus
    End If
    Set oBusqueda = Nothing
End Sub

Private Sub txtFum_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFum
End Sub


'debb-09/09/2016
Private Sub txtFum_LostFocus()
    CargaEdadGestacional
End Sub

Sub CargaEdadGestacional()
    If sighentidades.EsFecha(txtFum, "DD/MM/AAAA") Then
       If UcPacienteDatos1.idTipoSexo = 1 Then
          MsgBox "El Paciente es MASCULINO no debe tener fecha de última mestruación (FUM)", vbInformation, Me.Caption
          txtFum.Text = sighentidades.FECHA_VACIA_DMY
          txtEG.Text = ""
       Else
          txtEG.Text = sighentidades.DevuelveEdadGestacional(CDate(txtFum.Text), CDate(txtFregistro.Text))
       End If
    Else
       txtFum.Text = sighentidades.FECHA_VACIA_DMY
       txtEG.Text = ""
    End If
End Sub

'debb-09/09/2016
Sub CargaFUMdeHistoricos(lnIdPaciente9 As Long, oConexion9 As Connection)
    If mi_Opcion = sghAgregar Then
        Me.txtFum.Text = mo_reglasComunes.DevuelveFUMenUltimaAtencion(lnIdPaciente9, CDate(txtFregistro.Text), oConexion9)
        CargaEdadGestacional
     End If
End Sub
'debb-09/09/2016
Private Sub txtEG_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub
'debb-09/09/2016
Private Sub txtEG_LostFocus()
   If txtEG.Text <> "" Then
        If UcPacienteDatos1.idTipoSexo = 1 Then
               MsgBox "El Paciente es MASCULINO no debe tener fecha de última mestruación (FUM)", vbInformation, Me.Caption
               txtFum.Text = sighentidades.FECHA_VACIA_DMY
               txtEG.Text = ""
        Else
             If Val(txtEG.Text) > 0 And Val(txtEG.Text) < 50 Then
                 txtFum.Text = sighentidades.DevuelveFUM(Val(txtEG.Text), CDate(txtFregistro.Text))
             Else
                 MsgBox "La Edad Gestacional(EG) no puede pasar de 50", vbInformation, Me.Caption
                 txtEG.Text = ""
                 txtFum.Text = sighentidades.FECHA_VACIA_DMY
             End If
        End If
   Else
        txtFum.Text = sighentidades.FECHA_VACIA_DMY
   End If
End Sub
Private Sub btnImprimir_Click()
    Dim oRep As New laboratorio
    oRep.ImpresionDeItems ml_IdMovimiento, Me.cmbResponsable.Text, Me.txtFrealizaCpt.Text, txtProcedencia.Text, Me.hwnd, _
                          IIf(cmbMedicos.Text = "", txtMedico.Text, cmbMedicos.Text)
    Set oRep = Nothing
End Sub



Private Sub txtNcita_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNcita
       AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNcita_LostFocus()
     If Val(txtNcita.Text) > 0 Then
        Dim oConexion As New Connection
        Dim oSiCitas As New SiCitas
        Dim DoSiCitas As New DoSiCitas
        Dim oDOCajaComprobantesPago As New DOCajaComprobantesPago
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 900
        oConexion.Open sighentidades.CadenaConexion
        Set oSiCitas.Conexion = oConexion
        DoSiCitas.IdUsuarioAuditoria = sighentidades.Usuario
        DoSiCitas.idCitaSI = Val(txtNcita.Text)
        Me.ucProductos.LimpiarGrilla
        If oSiCitas.SeleccionarPorId(DoSiCitas) = True Then
           If DoSiCitas.IdPuntoCarga <> ml_puntoCarga Then
                MsgBox "La Cita existe pero NO pertenece al PUNTO DE CARGA", vbInformation, ""
           ElseIf DoSiCitas.IdMovimiento > 0 Then
                MsgBox "La CITA ya tiene MOVIMIENTO N° " & DoSiCitas.IdMovimiento, vbInformation, ""
           Else
                If DoSiCitas.IdReceta > 0 Then
                   Me.ssoptPaciente.Value = True
                   txtNreceta.Text = DoSiCitas.IdReceta
                   txtNreceta_LostFocus
                ElseIf DoSiCitas.idCuentaAtencion > 0 Then
                   Me.ssoptPaciente.Value = True
                   txtNcuenta.Text = DoSiCitas.idCuentaAtencion
                   txtNcuenta_LostFocus
                   Me.ucProductos.CargaProductosPorIdCita Val(txtNcita.Text)
                ElseIf DoSiCitas.IdComprobantePago > 0 Then
                   Me.ssoptExterno.Value = True
                   Set oDOCajaComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(DoSiCitas.IdComprobantePago, oConexion)
                   txtNserie.Text = oDOCajaComprobantesPago.NroSerie
                   txtNboleta.Text = oDOCajaComprobantesPago.nroDocumento
                   txtNboleta_LostFocus
                   UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                   If DoSiCitas.idPaciente = 0 Then
                          If DoSiCitas.FechaNacimiento <> 0 Then
                             UcPacienteDatos1.FechaNacimiento = DoSiCitas.FechaNacimiento
                          End If
                          If DoSiCitas.idTipoSexo > 0 Then
                             UcPacienteDatos1.idTipoSexo = DoSiCitas.idTipoSexo
                          End If
                          UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta DoSiCitas.Paciente
                   Else
                          UcPacienteDatos1.idPaciente = DoSiCitas.idPaciente
                          UcPacienteDatos1.CargarDatosDePacienteALosControles
                   End If
                End If
           End If
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oSiCitas = Nothing
        Set DoSiCitas = Nothing
        Set oDOCajaComprobantesPago = Nothing
     End If
End Sub

