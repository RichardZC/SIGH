VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form HerrExportaSIS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporta/importa datos al SIS"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "HerrExportaSIS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6585
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   11615
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Importar Datos"
      TabPicture(0)   =   "HerrExportaSIS.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "Frame1(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Exportar datos"
      TabPicture(1)   =   "HerrExportaSIS.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Datos Generales"
      TabPicture(2)   =   "HerrExportaSIS.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "Frame10"
      Tab(2).Control(3)=   "FraFua2015"
      Tab(2).Control(4)=   "chkWeb"
      Tab(2).ControlCount=   5
      Begin VB.CheckBox chkWeb 
         Caption         =   "En este momento hay INTERNET ?"
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
         Left            =   -74895
         TabIndex        =   65
         Top             =   4890
         Width           =   3405
      End
      Begin VB.Frame FraFua2015 
         Caption         =   "Seleccionar el Formato Fua a usar en el EESS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   -74895
         TabIndex        =   64
         Top             =   3495
         Width           =   8100
         Begin VB.CommandButton btnVerDisenoFUA 
            Caption         =   "Vista Previa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   6720
            Picture         =   "HerrExportaSIS.frx":0D1E
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Visualizar el tipo de formato FUA configurado"
            Top             =   240
            Width           =   1275
         End
         Begin VB.ComboBox cmdFUAanexo 
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
            ItemData        =   "HerrExportaSIS.frx":1160
            Left            =   1800
            List            =   "HerrExportaSIS.frx":116A
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   675
            Width           =   4905
         End
         Begin VB.ComboBox cmbFUA 
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
            ItemData        =   "HerrExportaSIS.frx":1182
            Left            =   1800
            List            =   "HerrExportaSIS.frx":118C
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   285
            Width           =   4905
         End
         Begin VB.Label lblFuaAnexo 
            AutoSize        =   -1  'True
            Caption         =   "Anexo 2015"
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
            Left            =   135
            TabIndex        =   70
            Top             =   765
            Width           =   1005
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Formato FUA a usar"
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
            Left            =   135
            TabIndex        =   68
            Top             =   375
            Width           =   1605
         End
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   4440
         Width           =   8085
         Begin VB.TextBox txtTrama 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   60
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "(Parametro 332)"
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
            Left            =   5760
            TabIndex        =   62
            Top             =   330
            Width           =   1365
         End
         Begin VB.Label Label21 
            Caption         =   "Version de la guía de la trama de Interoperabilidad"
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
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Width           =   4245
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Consideraciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   -74880
         TabIndex        =   56
         Top             =   450
         Width           =   8085
         Begin VB.Label Label19 
            Caption         =   "* Chequear en tabla 'PARAMETROS' los valores para idParametros=301, 302, 303, 304, 305, 310, 322, 325, 327, 328, 360"
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
            Height          =   465
            Left            =   150
            TabIndex        =   57
            Top             =   330
            Width           =   7455
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Importa EESS/FF/MEDICAMENTOS/INSUMOS/MEDICOS/ nuevos y actualiza los existentes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   -74850
         TabIndex        =   47
         Top             =   3360
         Width           =   8085
         Begin VB.TextBox txtImportaESS 
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
            Left            =   1170
            MaxLength       =   30
            TabIndex        =   49
            Text            =   "c:\ACT13110009.zip"
            Top             =   330
            Width           =   6825
         End
         Begin VB.CommandButton btnActualizaEESS 
            Caption         =   "Actualiza TABLAS SIS"
            DisabledPicture =   "HerrExportaSIS.frx":11A4
            DownPicture     =   "HerrExportaSIS.frx":1604
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   6300
            Picture         =   "HerrExportaSIS.frx":1A79
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   990
            Width           =   1365
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar9 
            Height          =   300
            Left            =   1230
            TabIndex        =   50
            Top             =   1530
            Width           =   4950
            _ExtentX        =   6826
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   16744576
         End
         Begin VB.Label lblTabla 
            Caption         =   "...."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   180
            TabIndex        =   55
            Top             =   1530
            Width           =   1005
         End
         Begin VB.Label Label17 
            Caption         =   "(Agrega DATOS de: EESS/FF/MEDIC/INSUMOS)"
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
            Height          =   315
            Left            =   1200
            TabIndex        =   54
            Top             =   1230
            Width           =   5235
         End
         Begin VB.Label Label16 
            Caption         =   "(Agrega y actualiza DATOS de: MEDICOS)"
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
            Height          =   315
            Left            =   1200
            TabIndex        =   53
            Top             =   960
            Width           =   5235
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Archivo ZIP"
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
            Left            =   150
            TabIndex        =   52
            Top             =   390
            Width           =   930
         End
         Begin VB.Label Label14 
            Caption         =   "(previamente ha sido bajado de la página Web del SIS)"
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
            Height          =   315
            Left            =   1200
            TabIndex        =   51
            Top             =   720
            Width           =   5235
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "N° Formato FUA (al iniciar cada AÑO)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74880
         TabIndex        =   33
         Top             =   1590
         Width           =   8085
         Begin VB.ComboBox cmbTipo 
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
            ItemData        =   "HerrExportaSIS.frx":1EEE
            Left            =   1845
            List            =   "HerrExportaSIS.frx":1EF8
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1140
            Width           =   2220
         End
         Begin VB.Frame Frame8 
            Caption         =   "N° FUA"
            Height          =   1275
            Left            =   4230
            TabIndex        =   39
            Top             =   180
            Width           =   3435
            Begin VB.TextBox txtNumeroInicio 
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
               Left            =   1500
               MaxLength       =   8
               TabIndex        =   40
               Top             =   180
               Width           =   1785
            End
            Begin VB.TextBox txtNumeroFinal 
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
               Left            =   1500
               MaxLength       =   8
               TabIndex        =   41
               Top             =   540
               Width           =   1785
            End
            Begin VB.TextBox txtNumeroUltimo 
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
               Left            =   1500
               MaxLength       =   8
               TabIndex        =   43
               Top             =   900
               Width           =   1785
            End
            Begin VB.Label Label11 
               Caption         =   "Inicial"
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
               Left            =   120
               TabIndex        =   45
               Top             =   210
               Width           =   1215
            End
            Begin VB.Label Label12 
               Caption         =   "hasta"
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
               Left            =   120
               TabIndex        =   44
               Top             =   570
               Width           =   765
            End
            Begin VB.Label Label13 
               Caption         =   "último generado"
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
               Left            =   120
               TabIndex        =   42
               Top             =   930
               Width           =   1335
            End
         End
         Begin VB.TextBox txtLote 
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
            Left            =   1845
            MaxLength       =   2
            TabIndex        =   36
            Top             =   690
            Width           =   2190
         End
         Begin VB.TextBox txtDisa 
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
            Left            =   1845
            MaxLength       =   3
            TabIndex        =   34
            Top             =   300
            Width           =   2190
         End
         Begin VB.Label Label18 
            Caption         =   "Forma en que SisGalenPlus se usará el N° Formato FUA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   75
            TabIndex        =   67
            Top             =   1020
            Width           =   1770
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
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
            TabIndex        =   37
            Top             =   690
            Width           =   375
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Disa"
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
            TabIndex        =   35
            Top             =   300
            Width           =   315
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1215
         Left            =   -74910
         TabIndex        =   30
         Top             =   5220
         Width           =   8085
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "Actualizar "
            DisabledPicture =   "HerrExportaSIS.frx":1F10
            DownPicture     =   "HerrExportaSIS.frx":2370
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2820
            Picture         =   "HerrExportaSIS.frx":27E5
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   210
            Width           =   1365
         End
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "HerrExportaSIS.frx":2C5A
            DownPicture     =   "HerrExportaSIS.frx":311E
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4275
            Picture         =   "HerrExportaSIS.frx":360A
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   210
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   -74850
         TabIndex        =   26
         Top             =   5310
         Width           =   8085
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "HerrExportaSIS.frx":3AF6
            DownPicture     =   "HerrExportaSIS.frx":3FBA
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3960
            Picture         =   "HerrExportaSIS.frx":44A6
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Importa acreditados nuevos y actualiza los existentes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Index           =   0
         Left            =   -74850
         TabIndex        =   23
         Top             =   450
         Width           =   8085
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Actualiza Acreditados"
            DisabledPicture =   "HerrExportaSIS.frx":4992
            DownPicture     =   "HerrExportaSIS.frx":4DF2
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   6300
            Picture         =   "HerrExportaSIS.frx":5267
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1500
            Width           =   1365
         End
         Begin VB.CheckBox chkUsaPA 
            Caption         =   "Usar Procedimiento Almacenado del Servidor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            TabIndex        =   38
            Top             =   1350
            Width           =   5475
         End
         Begin VB.TextBox txtNacreditados 
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
            Left            =   2430
            MaxLength       =   30
            TabIndex        =   25
            Text            =   "c:\acreditadosEnero2013.zip"
            Top             =   360
            Width           =   5595
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar8 
            Height          =   300
            Left            =   210
            TabIndex        =   28
            Top             =   1770
            Width           =   5670
            _ExtentX        =   6826
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   16711680
         End
         Begin VB.Label Label20 
            Caption         =   "(eliminar contenido de carpeta      c:\DebbSIS\*.*     )"
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
            Height          =   315
            Left            =   2430
            TabIndex        =   58
            Top             =   1050
            Width           =   5235
         End
         Begin VB.Label Label8 
            Caption         =   "(previamente ha sido bajado de la página Web del SIS)"
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
            Height          =   315
            Left            =   2430
            TabIndex        =   29
            Top             =   750
            Width           =   5235
         End
         Begin VB.Label Label7 
            Caption         =   "Archivo ZIP de acreditados"
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
            Left            =   150
            TabIndex        =   24
            Top             =   390
            Width           =   2265
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   5250
         Width           =   8085
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
            Height          =   855
            Left            =   180
            Picture         =   "HerrExportaSIS.frx":56DC
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   210
            Width           =   1365
         End
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Exporta al SIS"
            DisabledPicture =   "HerrExportaSIS.frx":5BB5
            DownPicture     =   "HerrExportaSIS.frx":6015
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3060
            Picture         =   "HerrExportaSIS.frx":648A
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   210
            Width           =   1365
         End
         Begin VB.CommandButton btnCancelar 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "HerrExportaSIS.frx":68FF
            DownPicture     =   "HerrExportaSIS.frx":6DC3
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4515
            Picture         =   "HerrExportaSIS.frx":72AF
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   210
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Exporta datos de FUAs generadas en GalenHos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   1
         Top             =   450
         Width           =   8085
         Begin MSMask.MaskEdBox txtFechaDesde 
            Height          =   315
            Left            =   1830
            TabIndex        =   2
            Top             =   300
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MSMask.MaskEdBox txtFechaHasta 
            Height          =   315
            Left            =   3780
            TabIndex        =   3
            Top             =   300
            Width           =   1380
            _ExtentX        =   2434
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
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar1 
            Height          =   300
            Left            =   1830
            TabIndex        =   4
            Top             =   1500
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar2 
            Height          =   300
            Left            =   1830
            TabIndex        =   5
            Top             =   1890
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar3 
            Height          =   300
            Left            =   1830
            TabIndex        =   6
            Top             =   2310
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar4 
            Height          =   300
            Left            =   5970
            TabIndex        =   7
            Top             =   1470
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar5 
            Height          =   300
            Left            =   5970
            TabIndex        =   8
            Top             =   1890
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar6 
            Height          =   300
            Left            =   5970
            TabIndex        =   9
            Top             =   2340
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar7 
            Height          =   300
            Left            =   180
            TabIndex        =   10
            Top             =   1080
            Width           =   7500
            _ExtentX        =   13229
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin VB.Label Label23 
            Caption         =   "(eliminar la carpeta      c:\EnviosFua  y   c:\DebbSIS )"
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
            Height          =   315
            Left            =   150
            TabIndex        =   63
            Top             =   690
            Width           =   5235
         End
         Begin VB.Label Label6 
            Caption         =   "AtencionPRO.txt"
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
            Left            =   4140
            TabIndex        =   18
            Top             =   2340
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "AtencionMED.txt"
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
            Left            =   4140
            TabIndex        =   17
            Top             =   1920
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "AtencionSMI.txt"
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
            Left            =   4140
            TabIndex        =   16
            Top             =   1515
            Width           =   1605
         End
         Begin VB.Label Label3 
            Caption         =   "AtencionINS.txt"
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
            Left            =   150
            TabIndex        =   15
            Top             =   2340
            Width           =   1605
         End
         Begin VB.Label Label2 
            Caption         =   "AtencionDIA.txt"
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
            Left            =   150
            TabIndex        =   14
            Top             =   1920
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Atencion.txt"
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
            Left            =   150
            TabIndex        =   13
            Top             =   1515
            Width           =   1365
         End
         Begin VB.Label lblFechaRequerida 
            Caption         =   "hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3330
            TabIndex        =   12
            Top             =   330
            Width           =   765
         End
         Begin VB.Label lblFechaSolicitud 
            Caption         =   "F. Atención"
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
            Left            =   180
            TabIndex        =   11
            Top             =   300
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "HerrExportaSIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Exporta e Importa información del SIS
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim ml_IdUsuario As Long
Dim lcSql As String, lbSeImprime As Boolean
Dim mo_lcNombrePc  As String
Dim lnProgresBarElegido As Long
Private WithEvents oProcesos As SIGHSis.ReglasSISgalenhos
Attribute oProcesos.VB_VarHelpID = -1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let IdUsuario(lIdValue As Long)
    ml_IdUsuario = lIdValue
End Property

Private Sub btnAceptar_Click()
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 11
        oProcesos.IdUsuario = ml_IdUsuario
        oProcesos.lcNombrePc = mo_lcNombrePc
        oProcesos.ExportaSIS Me.txtFechaDesde.Text, Me.txtFechaHasta.Text, Me.hwnd
        Set oProcesos = Nothing
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub btnActualizaEESS_Click()
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       lnProgresBarElegido = 9
       Dim ldHoraInicial As Date, ldHoraFInal As Date
       ldHoraInicial = lcBuscaParametro.RetornaHoraServidorSQL1
       oProcesos.IdUsuario = ml_IdUsuario
       oProcesos.lcNombrePc = mo_lcNombrePc
       oProcesos.ImportaEESSsis txtImportaESS.Text
       Me.MousePointer = 1
       ldHoraFInal = lcBuscaParametro.RetornaHoraServidorSQL1
       MsgBox "Terminó el proceso en: " & Round(DateDiff("s", ldHoraInicial, ldHoraFInal) / 60, 2) & "  seg", vbInformation, Me.Caption
       If oProcesos.MensajeError <> "" Then
           MsgBox "Errores: " & Chr(13) & Left(oProcesos.MensajeError, 500)
       Else
          Me.Visible = False
       End If
       Set oProcesos = Nothing
    End If
End Sub

Private Sub btnVerDisenoFUA_Click()
        If Left(cmbFUA.Text, 1) = "B" And Left(cmdFUAanexo.Text, 1) = "3" Then
            MsgBox "Seleccione un tipo de Anexo 2015 para el FUA", vbInformation, Me.Caption
            Exit Sub
        End If
        Dim Ruta As String
        Dim lcFormatoFua As String
        Dim lcTipoAnexo As String
        lcFormatoFua = Left(cmbFUA.Text, 1)
        lcTipoAnexo = IIf(cmdFUAanexo.Visible = True, Left(cmdFUAanexo.Text, 1), "")
        Ruta = App.Path + "\Imagenes\FUA\" + lcFormatoFua + lcTipoAnexo + "\" + lcFormatoFua + lcTipoAnexo + "-1.png"
        Dim ret As Long
        ret = ShellExecute(Me.hwnd, "Open", Ruta, "", "", 1)
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub btnImprimir_Click()
        Dim oRsTmp1 As New Recordset
        Dim mrs_Tmp As New Recordset
        Dim lnTotalReg As Long
        Dim oConexionExterna As New Connection
        Dim oProcesos1 As New Procesos
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        oConexionExterna.CursorLocation = adUseClient
        Set oProcesos1.ProgressRpt1 = Me.XP_ProgressBar1
        Set oProcesos1.progressRpt2 = Me.XP_ProgressBar2
        Set oProcesos1.progressRpt3 = Me.XP_ProgressBar3
        Set oProcesos1.progressRpt4 = Me.XP_ProgressBar4
        Set oProcesos1.progressRpt5 = Me.XP_ProgressBar5
        Set oProcesos1.progressRpt6 = Me.XP_ProgressBar6
        Set oProcesos1.progressRpt7 = Me.XP_ProgressBar7
        Set oRsTmp1 = oProcesos1.ChequeaCuentasSinFUAgrabado(mrs_Tmp, oConexionExterna, Me.txtFechaDesde.Text, Me.txtFechaHasta.Text)
        If mrs_Tmp.RecordCount > 0 Then
            oProcesos1.CrearReporteObservaciones_excel mrs_Tmp, Me.hwnd
        Else
           MsgBox "No se encontrol FUA con problemas", vbInformation, Me.Caption
        End If
        Set oRsTmp1 = Nothing
        Set mrs_Tmp = Nothing
        Set oConexionExterna = Nothing
End Sub





Private Sub cmbFUA_Click()
    If Left(cmbFUA.Text, 1) = "B" Then
       lblFuaAnexo.Visible = True
       cmdFUAanexo.Visible = True
    Else
       lblFuaAnexo.Visible = False
       cmdFUAanexo.Visible = False
    End If
End Sub

Private Sub cmbTipo_Click()
'    If Me.cmbTipo.ListIndex = 0 Then '0 = Manual
        mo_Formulario.HabilitarDeshabilitar Me.txtNumeroInicio, True
        mo_Formulario.HabilitarDeshabilitar Me.txtNumeroFinal, True
        mo_Formulario.HabilitarDeshabilitar Me.txtNumeroUltimo, False
        If Trim(Me.txtNumeroFinal.Text) <> "" Then
            BuscarNumeroUltimo_DatosGenerales
        End If
'    Else ' 1 = Automatico
'        Me.txtNumeroInicio.Text = "": mo_Formulario.HabilitarDeshabilitar Me.txtNumeroInicio, False
'        Me.txtNumeroFinal.Text = "": mo_Formulario.HabilitarDeshabilitar Me.txtNumeroFinal, False
'        Me.txtNumeroUltimo.Text = "": mo_Formulario.HabilitarDeshabilitar Me.txtNumeroUltimo, False
'    End If
End Sub

Private Sub cmdAceptar_Click()
    If MsgBox("¿Esta seguro?", vbQuestion + vbYesNo, "Importación de Acreditados SIS") = vbYes Then
       Me.MousePointer = 11
       lnProgresBarElegido = 8
       Dim ldHoraInicial As Date, ldHoraFInal As Date
       ldHoraInicial = lcBuscaParametro.RetornaHoraServidorSQL1
       oProcesos.IdUsuario = ml_IdUsuario
       oProcesos.lcNombrePc = mo_lcNombrePc
       If chkUsaPA.Value = 1 Then
          oProcesos.ImportaAcreditadosSIScursor txtNacreditados.Text
       Else
          oProcesos.ImportaAcreditadosSIS txtNacreditados.Text
       End If
       Me.MousePointer = 1
       ldHoraFInal = lcBuscaParametro.RetornaHoraServidorSQL1
       MsgBox "Terminó el proceso en: " & Round(DateDiff("s", ldHoraInicial, ldHoraFInal) / 60, 2) & "  seg", vbInformation, Me.Caption
       If oProcesos.MensajeError <> "" Then
           MsgBox "Errores: " & Chr(13) & Left(oProcesos.MensajeError, 500)
       Else
          Me.Visible = False
       End If
       Set oProcesos = Nothing
    End If
End Sub

Private Sub cmdActualizar_Click()
    If Me.cmbTipo.ListIndex = 0 Then '0 = Manual
        If Val(txtNumeroInicio.Text) <= 0 Then
           MsgBox "El número INICIAL debe ser mayor a CERO", vbInformation, Me.Caption
           Exit Sub
        End If
        If Val(Me.txtNumeroFinal.Text) <= 0 Then
           MsgBox "El número FINAL debe ser mayor a CERO", vbInformation, Me.Caption
           Exit Sub
        ElseIf Val(Me.txtNumeroFinal.Text) <= Val(Me.txtNumeroInicio.Text) Then
           MsgBox "El número FINAL debe ser mayor al número INICIAL", vbInformation, Me.Caption
           Exit Sub
        End If
        If Me.txtNumeroUltimo.Text <> "" Then
            If Not (Val(Me.txtNumeroUltimo.Text) >= Val(Me.txtNumeroInicio.Text) And Val(Me.txtNumeroUltimo.Text) <= Val(Me.txtNumeroFinal.Text)) Then
               MsgBox "El ULTIMO NUMERO GENERADO debe estar entre el INICIAL y FINAL", vbInformation, Me.Caption
               Exit Sub
            End If
        End If
        'BuscarNumeroUltimo_DatosGenerales
    End If
    If mo_ReglasSISgalenhos.SisFuaActualizaDatos(Me.txtDisa.Text, Me.txtLote.Text, Me.txtNumeroInicio.Text, _
                       Me.txtNumeroFinal.Text, Me.txtNumeroUltimo.Tag, Me.cmbTipo.ListIndex + 1, _
                       IIf(Me.chkWeb.Value = 1, "S", "n"), Left(cmbFUA.Text, 1), Left(cmdFUAanexo.Text, 1)) = True Then
       Me.Visible = False
    Else
       MsgBox mo_ReglasSISgalenhos.MensajeError, vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdCancelar_Click()
    Me.Visible = False
End Sub



Private Sub cmdSalir_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
    
    Set oProcesos = New SIGHSis.ReglasSISgalenhos
    XP_ProgressBar8.ShowText = True
    XP_ProgressBar9.ShowText = True
    '
    txtFechaDesde.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    txtFechaHasta.Text = Date
    '
    CargaNumeroFuaDelAnioActual
    Me.cmbTipo.ListIndex = Val(lcBuscaParametro.SeleccionaFilaParametro(320)) - 1
    Me.chkWeb.Value = IIf(lcBuscaParametro.SeleccionaFilaParametro(322) = "S", 1, 0)
    '
    Me.txtTrama.Text = Format(lcBuscaParametro.SeleccionaFilaParametro(332), "0.00")
    mo_Formulario.HabilitarDeshabilitar txtTrama, False
    '
    lcBuscaParametro.DevuelveComboLlenoSegunDescripcion cmbFUA, 358
    lcBuscaParametro.DevuelveComboLlenoSegunDescripcion cmdFUAanexo, 359
    cmbFUA_Click
    '
End Sub

Sub CargaNumeroFuaDelAnioActual()
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaSeleccionarTodos
    If oRsTmp1.RecordCount = 0 Then
       Me.txtDisa.Text = lcBuscaParametro.SeleccionaFilaParametro(310)
       Me.txtLote.Text = Right(Trim(Str(Year(CDate(lcBuscaParametro.RetornaFechaServidorSQL)))), 2)
       Me.txtNumeroInicio.Text = ""
       Me.txtNumeroFinal.Text = ""
       Me.txtNumeroUltimo.Text = ""
    Else
       Me.txtDisa.Text = oRsTmp1.Fields!fuaDisa
       Me.txtLote.Text = oRsTmp1.Fields!fuaLote
       Me.txtNumeroInicio.Text = IIf(IsNull(oRsTmp1.Fields!FuaNumeroInicial), "", oRsTmp1.Fields!FuaNumeroInicial)
       Me.txtNumeroFinal.Text = IIf(IsNull(oRsTmp1.Fields!FuaNumeroFinal), "", oRsTmp1.Fields!FuaNumeroFinal)
       If IsNull(oRsTmp1.Fields!FuaNumeroInicial) = False And IsNull(oRsTmp1.Fields!FuaUltimoGenerado) = False Then
            If Val(oRsTmp1.Fields!FuaNumeroInicial) > Val(oRsTmp1.Fields!FuaUltimoGenerado) Then
                Me.txtNumeroUltimo.Text = ""
                Me.txtNumeroUltimo.Tag = oRsTmp1.Fields!FuaUltimoGenerado
            Else
                If Val(oRsTmp1.Fields!FuaNumeroFinal) >= Val(oRsTmp1.Fields!FuaUltimoGenerado) Then
                    Me.txtNumeroUltimo.Text = IIf(IsNull(oRsTmp1.Fields!FuaUltimoGenerado), "", oRsTmp1.Fields!FuaUltimoGenerado)
                    Me.txtNumeroUltimo.Tag = IIf(IsNull(oRsTmp1.Fields!FuaUltimoGenerado), "", oRsTmp1.Fields!FuaUltimoGenerado)
                Else
                    Me.txtNumeroUltimo.Text = ""
                    Me.txtNumeroUltimo.Tag = oRsTmp1.Fields!FuaUltimoGenerado
                End If
            End If
       End If
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub oProcesos_ProgressActualizaValor(lnValorActual As Long, lnValorTotal As Long, lcTablaActualizar As String)
    Me.Refresh
    Select Case lnProgresBarElegido
    Case 8
        XP_ProgressBar8.Max = lnValorTotal
        XP_ProgressBar8.Min = 0
        XP_ProgressBar8.Value = lnValorActual
    Case 9
        XP_ProgressBar9.Max = lnValorTotal
        XP_ProgressBar9.Min = 0
        XP_ProgressBar9.Value = lnValorActual
        lblTabla.Caption = lcTablaActualizar
    End Select
    DoEvents
End Sub

Private Sub txtFechaDesde_LostFocus()
If Not EsFecha(txtFechaDesde.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaDesde.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtfechaHasta_LostFocus()
If Not EsFecha(txtFechaHasta.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaHasta.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtNumeroFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNumeroFinal
End Sub

Private Sub txtNumeroFinal_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNumeroInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNumeroInicio
End Sub

Private Sub txtNumeroInicio_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNumeroUltimo_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNumeroInicio_LostFocus()
    BuscarNumeroUltimo_DatosGenerales
End Sub

Private Sub txtNumeroFinal_LostFocus()
    BuscarNumeroUltimo_DatosGenerales
End Sub

Sub BuscarNumeroUltimo_DatosGenerales()
    If txtNumeroInicio.Text = "" Then Exit Sub
    If txtNumeroFinal.Text = "" Then Exit Sub
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionConsultarUltimoNumero(txtNumeroInicio.Text, txtNumeroFinal.Text)
    If oRsTmp1.RecordCount = 0 Then
        Me.txtNumeroUltimo.Text = Me.txtNumeroInicio.Text
        Me.txtNumeroUltimo.Tag = Trim(CStr(Val(txtNumeroInicio.Text) - 1))
    Else
        oRsTmp1.MoveFirst
        Me.txtNumeroUltimo.Text = oRsTmp1.Fields!FuaNumero
        Me.txtNumeroUltimo.Tag = oRsTmp1.Fields!FuaNumero
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub




