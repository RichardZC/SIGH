VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form frmBancoSangre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BANCO DE SANGRE"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13335
   Icon            =   "frmBancoSangre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   60
      TabIndex        =   253
      Top             =   60
      Width           =   13215
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
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   254
         Top             =   150
         Width           =   6060
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   11010
         TabIndex        =   255
         Top             =   150
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   150
         TabIndex        =   257
         Top             =   210
         Width           =   1155
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
         Left            =   10050
         TabIndex        =   256
         Top             =   210
         Width           =   945
      End
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   60
      TabIndex        =   136
      Top             =   8340
      Width           =   13215
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmBancoSangre.frx":0CCA
         DownPicture     =   "frmBancoSangre.frx":118E
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
         Left            =   6765
         Picture         =   "frmBancoSangre.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   135
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
         Height          =   700
         Left            =   60
         Picture         =   "frmBancoSangre.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   150
         Width           =   1365
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmBancoSangre.frx":203F
         DownPicture     =   "frmBancoSangre.frx":249F
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
         Left            =   5205
         Picture         =   "frmBancoSangre.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Frame BAS001 
      ForeColor       =   &H00000000&
      Height          =   7815
      Left            =   60
      TabIndex        =   137
      Top             =   540
      Width           =   13215
      Begin VB.TextBox BS000_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   5700
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox BS000_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1740
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox BS000_03 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         ItemData        =   "frmBancoSangre.frx":2D89
         Left            =   9480
         List            =   "frmBancoSangre.frx":2D99
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   2145
      End
      Begin VB.ComboBox BS000_03 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         ItemData        =   "frmBancoSangre.frx":2DCB
         Left            =   5700
         List            =   "frmBancoSangre.frx":2DD5
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   570
         Width           =   1575
      End
      Begin VB.ComboBox BS000_03 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         ItemData        =   "frmBancoSangre.frx":2DED
         Left            =   1740
         List            =   "frmBancoSangre.frx":2DFD
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   570
         Width           =   1575
      End
      Begin TabDlg.SSTab BS000_04 
         Height          =   6735
         Left            =   120
         TabIndex        =   252
         Top             =   1020
         Width           =   13005
         _ExtentX        =   22939
         _ExtentY        =   11880
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
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
         TabCaption(0)   =   "Datos Personales"
         TabPicture(0)   =   "frmBancoSangre.frx":2E0E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "BS000_06(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "BS000_05(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "BS000_05(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "BS000_07"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Protocolo de Selección"
         TabPicture(1)   =   "frmBancoSangre.frx":2E2A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "BS000_05(2)"
         Tab(1).Control(1)=   "BS000_06(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Exámenes"
         TabPicture(2)   =   "frmBancoSangre.frx":2E46
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "BS000_05(8)"
         Tab(2).Control(1)=   "BS000_05(7)"
         Tab(2).Control(2)=   "BS000_06(2)"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Calificación del Donante"
         TabPicture(3)   =   "frmBancoSangre.frx":2E62
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "BS000_05(9)"
         Tab(3).Control(1)=   "BS000_06(3)"
         Tab(3).ControlCount=   2
         Begin VB.Frame BS000_05 
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   1335
            Index           =   9
            Left            =   -70560
            TabIndex        =   208
            Top             =   2520
            Width           =   3615
            Begin Threed.SSOption optApto 
               Height          =   285
               Index           =   2
               Left            =   465
               TabIndex        =   132
               Top             =   840
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   503
               _Version        =   262144
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3.- NO APTO PERMANENTEMENTE"
            End
            Begin Threed.SSOption optApto 
               Height          =   255
               Index           =   1
               Left            =   465
               TabIndex        =   131
               Top             =   480
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   450
               _Version        =   262144
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2.- NO APTO TEMPORALMENTE"
            End
            Begin Threed.SSOption optApto 
               Height          =   285
               Index           =   0
               Left            =   465
               TabIndex        =   130
               Top             =   90
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   262144
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1.- APTO"
            End
         End
         Begin VB.Frame BS000_05 
            Caption         =   "4.- Exámenes Complementarios"
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
            Height          =   2930
            Index           =   8
            Left            =   -74760
            TabIndex        =   192
            Top             =   3060
            Width           =   12525
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   29
               Left            =   1920
               TabIndex        =   117
               Top             =   900
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   25
               Left            =   1920
               TabIndex        =   115
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   27
               Left            =   1920
               TabIndex        =   116
               Top             =   570
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   31
               Left            =   1920
               TabIndex        =   118
               Top             =   1230
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   37
               Left            =   1920
               TabIndex        =   121
               Top             =   2220
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   33
               Left            =   1920
               TabIndex        =   119
               Top             =   1560
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   35
               Left            =   1920
               TabIndex        =   120
               Top             =   1890
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   39
               Left            =   1920
               TabIndex        =   122
               Top             =   2550
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   30
               Left            =   10770
               TabIndex        =   125
               Top             =   930
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   26
               Left            =   10770
               TabIndex        =   123
               Top             =   270
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   28
               Left            =   10770
               TabIndex        =   124
               Top             =   600
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   32
               Left            =   10770
               TabIndex        =   126
               Top             =   1260
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   38
               Left            =   10770
               TabIndex        =   129
               Top             =   2250
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   34
               Left            =   10770
               TabIndex        =   127
               Top             =   1590
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   36
               Left            =   10770
               TabIndex        =   128
               Top             =   1920
               Width           =   1515
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Hematocrito"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   40
               Left            =   180
               TabIndex        =   207
               Top             =   270
               Width           =   1005
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "HbsAg"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   44
               Left            =   180
               TabIndex        =   206
               Top             =   930
               Width           =   525
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Hemoglobina"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   42
               Left            =   180
               TabIndex        =   205
               Top             =   600
               Width           =   1050
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "VDRL/RPR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   46
               Left            =   180
               TabIndex        =   204
               Top             =   1260
               Width           =   825
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo Sanguíneo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   48
               Left            =   180
               TabIndex        =   203
               Top             =   1590
               Width           =   1410
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Fenotipo RH"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   52
               Left            =   180
               TabIndex        =   202
               Top             =   2250
               Width           =   1005
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Factor RH"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   50
               Left            =   180
               TabIndex        =   201
               Top             =   1920
               Width           =   795
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Anti HTLV"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   54
               Left            =   180
               TabIndex        =   200
               Top             =   2580
               Width           =   840
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Anti Core VHB"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   41
               Left            =   9555
               TabIndex        =   199
               Top             =   300
               Width           =   1170
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Anti VIH"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   45
               Left            =   10035
               TabIndex        =   198
               Top             =   960
               Width           =   690
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Anti Chagas"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   43
               Left            =   9765
               TabIndex        =   197
               Top             =   630
               Width           =   960
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Anti VHC"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   47
               Left            =   9990
               TabIndex        =   196
               Top             =   1290
               Width           =   735
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Malaria"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   49
               Left            =   10155
               TabIndex        =   195
               Top             =   1620
               Width           =   525
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Variante Du"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   53
               Left            =   9720
               TabIndex        =   194
               Top             =   2280
               Width           =   960
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Bartonella"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   51
               Left            =   9885
               TabIndex        =   193
               Top             =   1950
               Width           =   795
            End
         End
         Begin VB.Frame BS000_05 
            Caption         =   "2.- Protocolo de Selección al Donante de Sangre"
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
            Height          =   6060
            Index           =   2
            Left            =   -74880
            TabIndex        =   165
            Top             =   480
            Width           =   12735
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   25
               Left            =   5520
               TabIndex        =   248
               Top             =   3840
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   18
                  Left            =   0
                  TabIndex        =   59
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   18
                  Left            =   570
                  TabIndex        =   60
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   24
               Left            =   5520
               TabIndex        =   247
               Top             =   3600
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   17
                  Left            =   0
                  TabIndex        =   57
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   17
                  Left            =   570
                  TabIndex        =   58
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   23
               Left            =   5520
               TabIndex        =   246
               Top             =   3360
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   16
                  Left            =   0
                  TabIndex        =   55
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   16
                  Left            =   570
                  TabIndex        =   56
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   22
               Left            =   5520
               TabIndex        =   245
               Top             =   3120
               Width           =   1095
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   13
                  Left            =   570
                  TabIndex        =   54
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   13
                  Left            =   0
                  TabIndex        =   53
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   21
               Left            =   5520
               TabIndex        =   244
               Top             =   2880
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   12
                  Left            =   0
                  TabIndex        =   51
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   12
                  Left            =   570
                  TabIndex        =   52
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   20
               Left            =   5520
               TabIndex        =   243
               Top             =   2640
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   10
                  Left            =   0
                  TabIndex        =   49
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   10
                  Left            =   570
                  TabIndex        =   50
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   19
               Left            =   5520
               TabIndex        =   242
               Top             =   2400
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   9
                  Left            =   0
                  TabIndex        =   47
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   9
                  Left            =   570
                  TabIndex        =   48
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   18
               Left            =   5520
               TabIndex        =   241
               Top             =   2160
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   8
                  Left            =   0
                  TabIndex        =   45
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   8
                  Left            =   570
                  TabIndex        =   46
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   17
               Left            =   5520
               TabIndex        =   240
               Top             =   1920
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   7
                  Left            =   0
                  TabIndex        =   43
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   7
                  Left            =   570
                  TabIndex        =   44
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   16
               Left            =   5520
               TabIndex        =   239
               Top             =   1680
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   6
                  Left            =   0
                  TabIndex        =   41
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   6
                  Left            =   570
                  TabIndex        =   42
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   15
               Left            =   5520
               TabIndex        =   238
               Top             =   1440
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   5
                  Left            =   0
                  TabIndex        =   39
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   5
                  Left            =   600
                  TabIndex        =   40
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   14
               Left            =   5520
               TabIndex        =   237
               Top             =   1200
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   4
                  Left            =   0
                  TabIndex        =   37
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   4
                  Left            =   570
                  TabIndex        =   38
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   13
               Left            =   5520
               TabIndex        =   236
               Top             =   960
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   3
                  Left            =   0
                  TabIndex        =   35
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   3
                  Left            =   570
                  TabIndex        =   36
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   12
               Left            =   5520
               TabIndex        =   235
               Top             =   720
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   2
                  Left            =   0
                  TabIndex        =   33
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   2
                  Left            =   570
                  TabIndex        =   34
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   11
               Left            =   5520
               TabIndex        =   234
               Top             =   480
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   31
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   1
                  Left            =   570
                  TabIndex        =   32
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   10
               Left            =   5520
               TabIndex        =   233
               Top             =   240
               Width           =   1095
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   29
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   0
                  Left            =   570
                  TabIndex        =   30
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               Caption         =   "16.- ¿Ha tenido contacto sexual con algún grupo de riesgo?"
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
               Height          =   1215
               Index           =   6
               Left            =   6960
               TabIndex        =   175
               Top             =   4080
               Width           =   5655
               Begin VB.CheckBox chkCGR 
                  Caption         =   "Homosexual"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   60
                  TabIndex        =   103
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.CheckBox chkCGR 
                  Caption         =   "Bisexual"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   60
                  TabIndex        =   104
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.CheckBox chkCGR 
                  Caption         =   "Promiscuo"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   2
                  Left            =   60
                  TabIndex        =   105
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.CheckBox chkCGR 
                  Caption         =   "Prostituta"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   3
                  Left            =   60
                  TabIndex        =   106
                  Top             =   960
                  Width           =   1455
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   15
                  Left            =   3570
                  TabIndex        =   108
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Enabled         =   0   'False
                  Caption         =   "No"
               End
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   15
                  Left            =   3000
                  TabIndex        =   107
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Enabled         =   0   'False
                  Caption         =   "Si"
               End
            End
            Begin VB.Frame BS000_05 
               Caption         =   "12.- ¿Ha tenido o tiene alguna(s) de las siguientes enfermedades?"
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
               Height          =   2415
               Index           =   3
               Left            =   6960
               TabIndex        =   174
               Top             =   240
               Width           =   5655
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Glomerulonefritis"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   23
                  Left            =   3960
                  TabIndex        =   88
                  Top             =   1440
                  Width           =   1575
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Osteomielitis (5a)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   22
                  Left            =   3960
                  TabIndex        =   85
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Mononucleosis"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   21
                  Left            =   3960
                  TabIndex        =   82
                  Top             =   960
                  Width           =   1575
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Amebiasis (1a)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   20
                  Left            =   3960
                  TabIndex        =   79
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Hipertiroidismo"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   19
                  Left            =   3960
                  TabIndex        =   76
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Dengue (1a)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   18
                  Left            =   3960
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Transtornos de Coagulación"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   17
                  Left            =   1920
                  TabIndex        =   94
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Enfermedades Venéreas (3a)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   16
                  Left            =   1920
                  TabIndex        =   92
                  Top             =   1920
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Fiebre Reumática (Rp)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   15
                  Left            =   1920
                  TabIndex        =   90
                  Top             =   1680
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Asma"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   14
                  Left            =   1920
                  TabIndex        =   87
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Diabetes (Rp)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   13
                  Left            =   1920
                  TabIndex        =   84
                  Top             =   1200
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Cáncer (Rp)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   12
                  Left            =   1920
                  TabIndex        =   81
                  Top             =   960
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Hemorragias"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   11
                  Left            =   1920
                  TabIndex        =   78
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Convulsiones (Rp)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   10
                  Left            =   1920
                  TabIndex        =   75
                  Top             =   480
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Hipertensión Arterial"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   9
                  Left            =   1920
                  TabIndex        =   72
                  Top             =   240
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Cardiopatías (Rp)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   8
                  Left            =   90
                  TabIndex        =   93
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Bartolenosis"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   7
                  Left            =   90
                  TabIndex        =   91
                  Top             =   1920
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Chagas (Rp)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   6
                  Left            =   90
                  TabIndex        =   89
                  Top             =   1680
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Paludismo"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   5
                  Left            =   90
                  TabIndex        =   86
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Fiebre Amarilla (1a)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   4
                  Left            =   90
                  TabIndex        =   83
                  Top             =   1200
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Fiebre Malta (3a)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   3
                  Left            =   90
                  TabIndex        =   80
                  Top             =   960
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Fiebre Tifoidea (2a)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   2
                  Left            =   90
                  TabIndex        =   77
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Tuberculosis (5a)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   90
                  TabIndex        =   74
                  Top             =   480
                  Width           =   2415
               End
               Begin VB.CheckBox chkEnf 
                  Caption         =   "Hepatitis"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   90
                  TabIndex        =   71
                  Top             =   240
                  Width           =   2415
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   11
                  Left            =   5010
                  TabIndex        =   96
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Enabled         =   0   'False
                  Caption         =   "No"
               End
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   11
                  Left            =   4440
                  TabIndex        =   95
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Enabled         =   0   'False
                  Caption         =   "Si"
               End
            End
            Begin VB.Frame BS000_05 
               Caption         =   "15.- ¿Pertenece a algún grupo de riesgo?"
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
               Height          =   1215
               Index           =   4
               Left            =   6960
               TabIndex        =   173
               Top             =   2760
               Width           =   5655
               Begin VB.CheckBox chkPGR 
                  Caption         =   "Prostituta"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   3
                  Left            =   60
                  TabIndex        =   100
                  Top             =   960
                  Width           =   1455
               End
               Begin VB.CheckBox chkPGR 
                  Caption         =   "Promiscuo"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   2
                  Left            =   60
                  TabIndex        =   99
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.CheckBox chkPGR 
                  Caption         =   "Bisexual"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   60
                  TabIndex        =   98
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.CheckBox chkPGR 
                  Caption         =   "Homosexual"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   60
                  TabIndex        =   97
                  Top             =   240
                  Width           =   1455
               End
               Begin Threed.SSOption optSi 
                  Height          =   195
                  Index           =   14
                  Left            =   1800
                  TabIndex        =   101
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  _Version        =   262144
                  Enabled         =   0   'False
                  Caption         =   "Si"
               End
               Begin Threed.SSOption optNo 
                  Height          =   195
                  Index           =   14
                  Left            =   2370
                  TabIndex        =   102
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _Version        =   262144
                  Enabled         =   0   'False
                  Caption         =   "No"
               End
            End
            Begin VB.Frame BS000_05 
               Caption         =   "Mujeres"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1935
               Index           =   5
               Left            =   240
               TabIndex        =   166
               Top             =   4080
               Visible         =   0   'False
               Width           =   6495
               Begin VB.Frame BS000_05 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Index           =   28
                  Left            =   3150
                  TabIndex        =   251
                  Top             =   1650
                  Width           =   1095
                  Begin Threed.SSOption optNo 
                     Height          =   195
                     Index           =   20
                     Left            =   585
                     TabIndex        =   70
                     Top             =   0
                     Width           =   615
                     _ExtentX        =   1085
                     _ExtentY        =   344
                     _Version        =   262144
                     Caption         =   "No"
                  End
                  Begin Threed.SSOption optSi 
                     Height          =   195
                     Index           =   20
                     Left            =   0
                     TabIndex        =   69
                     Top             =   0
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   344
                     _Version        =   262144
                     Caption         =   "Si"
                  End
               End
               Begin VB.Frame BS000_05 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Index           =   27
                  Left            =   3150
                  TabIndex        =   250
                  Top             =   1080
                  Width           =   1095
                  Begin Threed.SSOption optNo 
                     Height          =   195
                     Index           =   19
                     Left            =   600
                     TabIndex        =   67
                     Top             =   0
                     Width           =   615
                     _ExtentX        =   1085
                     _ExtentY        =   344
                     _Version        =   262144
                     Caption         =   "No"
                  End
                  Begin Threed.SSOption optSi 
                     Height          =   195
                     Index           =   19
                     Left            =   0
                     TabIndex        =   66
                     Top             =   0
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   344
                     _Version        =   262144
                     Caption         =   "Si"
                  End
               End
               Begin VB.Frame BS000_05 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Index           =   26
                  Left            =   3150
                  TabIndex        =   249
                  Top             =   840
                  Width           =   3255
                  Begin Threed.SSOption optMens 
                     Height          =   195
                     Index           =   2
                     Left            =   2385
                     TabIndex        =   65
                     Top             =   0
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   344
                     _Version        =   262144
                     Caption         =   "Escaso"
                  End
                  Begin Threed.SSOption optMens 
                     Height          =   195
                     Index           =   1
                     Left            =   1200
                     TabIndex        =   64
                     Top             =   0
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   344
                     _Version        =   262144
                     Caption         =   "Moderado"
                  End
                  Begin Threed.SSOption optMens 
                     Height          =   195
                     Index           =   0
                     Left            =   0
                     TabIndex        =   63
                     Top             =   0
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   344
                     _Version        =   262144
                     Caption         =   "Abundante"
                  End
               End
               Begin VB.TextBox BS000_01 
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
                  Height          =   285
                  Index           =   17
                  Left            =   3135
                  TabIndex        =   62
                  Top             =   480
                  Width           =   495
               End
               Begin VB.TextBox BS000_01 
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
                  Height          =   285
                  Index           =   16
                  Left            =   3135
                  TabIndex        =   61
                  Top             =   150
                  Width           =   1335
               End
               Begin VB.TextBox BS000_01 
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
                  Height          =   285
                  Index           =   18
                  Left            =   3135
                  TabIndex        =   68
                  Top             =   1335
                  Width           =   975
               End
               Begin VB.Label BS000_00 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "25.- ¿Está dando de lactar?"
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
                  Index           =   29
                  Left            =   1125
                  TabIndex        =   172
                  Top             =   1680
                  Width           =   1980
               End
               Begin VB.Label BS000_00 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "22.- En su menstruación, el sangrado es "
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
                  Index           =   26
                  Left            =   60
                  TabIndex        =   171
                  Top             =   840
                  Width           =   2940
               End
               Begin VB.Label BS000_00 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "20.- ¿Cuando fué su última regla?"
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
                  Index           =   24
                  Left            =   660
                  TabIndex        =   170
                  Top             =   180
                  Width           =   2415
               End
               Begin VB.Label BS000_00 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "23.- ¿Está gestando?"
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
                  Index           =   27
                  Left            =   1560
                  TabIndex        =   169
                  Top             =   1080
                  Width           =   1530
               End
               Begin VB.Label BS000_00 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "21.- ¿Cuántos días menstrúa?"
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
                  Index           =   25
                  Left            =   945
                  TabIndex        =   168
                  Top             =   510
                  Width           =   2145
               End
               Begin VB.Label BS000_00 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "24.- Fecha del último parto"
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
                  Index           =   28
                  Left            =   1155
                  TabIndex        =   167
                  Top             =   1365
                  Width           =   1935
               End
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "19.- ¿Ha sido excluido como donante anteriormente?"
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
               Index           =   23
               Left            =   1650
               TabIndex        =   191
               Top             =   3840
               Width           =   3780
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "14.- ¿Ha tenido contacto directo con personas que tengan ictericia?"
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
               Index           =   20
               Left            =   555
               TabIndex        =   190
               Top             =   3120
               Width           =   4860
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "1.- ¿Ha donado sangre alguna vez?"
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
               Left            =   2835
               TabIndex        =   189
               Top             =   240
               Width           =   2550
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "3.- ¿Se puso nervioso cuando donó sangre?"
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
               Left            =   2235
               TabIndex        =   188
               Top             =   720
               Width           =   3150
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "2.- ¿Donó sangre los últimos 3 meses?"
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
               Left            =   2655
               TabIndex        =   187
               Top             =   480
               Width           =   2730
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "10.- ¿Ha sido tatuado?"
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
               Index           =   17
               Left            =   3765
               TabIndex        =   186
               Top             =   2400
               Width           =   1635
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "5.- ¿Ha recibido sangre, tranplante de órganos ó tejidos?"
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
               Left            =   1290
               TabIndex        =   185
               Top             =   1200
               Width           =   4110
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "4.- ¿Ha sido operado en los últimos 6 meses?"
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
               Left            =   2190
               TabIndex        =   184
               Top             =   960
               Width           =   3210
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "11.- ¿Ha sido sometido a punción de piel (aretes, acupunturas)?"
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
               Left            =   825
               TabIndex        =   183
               Top             =   2640
               Width           =   4590
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "9.- ¿Está tomando alguna medicina?"
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
               Index           =   16
               Left            =   2835
               TabIndex        =   182
               Top             =   2160
               Width           =   2580
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "8.- ¿Ha usado drogas ilegales?"
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
               Index           =   15
               Left            =   3225
               TabIndex        =   181
               Top             =   1920
               Width           =   2190
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "13.- ¿Ha tenido contacto directo con personas que tengan hepatitis?"
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
               Left            =   480
               TabIndex        =   180
               Top             =   2880
               Width           =   4935
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "7.- ¿Viajó fuera del país en los últimos años?"
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
               Index           =   14
               Left            =   2250
               TabIndex        =   179
               Top             =   1680
               Width           =   3165
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "6.- ¿Ha viajado a zona endémica de paludismo?"
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
               Left            =   2025
               TabIndex        =   178
               Top             =   1440
               Width           =   3390
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "17.- ¿Tuvo contacto sexual con más de una persona en los últimos 3 años?"
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
               Index           =   21
               Left            =   60
               TabIndex        =   177
               Top             =   3360
               Width           =   5370
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "18.- ¿Tiene SIDA o ha tenido alguna prueba positiva de SIDA?"
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
               Index           =   22
               Left            =   975
               TabIndex        =   176
               Top             =   3600
               Width           =   4455
            End
         End
         Begin VB.Frame BS000_05 
            Caption         =   "3.- Examen Clínico"
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
            Height          =   2535
            Index           =   7
            Left            =   -74760
            TabIndex        =   154
            Top             =   480
            Width           =   12525
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   23
               Left            =   1950
               TabIndex        =   113
               Top             =   1560
               Width           =   10365
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   21
               Left            =   1950
               TabIndex        =   111
               Top             =   900
               Width           =   1875
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   19
               Left            =   1950
               TabIndex        =   109
               Top             =   240
               Width           =   1875
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   20
               Left            =   1950
               TabIndex        =   110
               Top             =   570
               Width           =   1875
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   555
               Index           =   24
               Left            =   1950
               MultiLine       =   -1  'True
               TabIndex        =   114
               Top             =   1890
               Width           =   10365
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   22
               Left            =   1950
               TabIndex        =   112
               Top             =   1230
               Width           =   1875
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Kg."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   31
               Left            =   3900
               TabIndex        =   164
               Top             =   270
               Width           =   270
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "mmHG"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   35
               Left            =   3900
               TabIndex        =   163
               Top             =   930
               Width           =   540
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "mt."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   33
               Left            =   3900
               TabIndex        =   162
               Top             =   600
               Width           =   285
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "pul/min"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   37
               Left            =   3900
               TabIndex        =   161
               Top             =   1260
               Width           =   600
            End
            Begin VB.Label BS000_00 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Estado de accesos venosos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   38
               Left            =   180
               TabIndex        =   160
               Top             =   1590
               Width           =   2040
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Peso"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   30
               Left            =   180
               TabIndex        =   159
               Top             =   270
               Width           =   390
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Presión Arterial"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   34
               Left            =   180
               TabIndex        =   158
               Top             =   930
               Width           =   1215
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Talla"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   32
               Left            =   180
               TabIndex        =   157
               Top             =   600
               Width           =   360
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Observaciones"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   39
               Left            =   180
               TabIndex        =   156
               Top             =   1905
               Width           =   1170
            End
            Begin VB.Label BS000_00 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Pulso"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   36
               Left            =   180
               TabIndex        =   155
               Top             =   1260
               Width           =   420
            End
         End
         Begin VB.CheckBox BS000_07 
            Caption         =   "Es para Donación por reposición"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   3600
            Visible         =   0   'False
            Width           =   3045
         End
         Begin VB.Frame BS000_05 
            Caption         =   "1. Datos del Donador"
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
            Height          =   2925
            Index           =   0
            Left            =   240
            TabIndex        =   138
            Top             =   510
            Width           =   12465
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   41
               Left            =   8370
               TabIndex        =   18
               Top             =   1290
               Width           =   3915
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   40
               Left            =   1920
               TabIndex        =   13
               Top             =   2220
               Width           =   4155
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   705
               Index           =   9
               Left            =   8370
               MultiLine       =   -1  'True
               TabIndex        =   20
               Top             =   1950
               Width           =   3885
            End
            Begin VB.ComboBox BS000_03 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   3
               Left            =   8370
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   645
               Width           =   3915
            End
            Begin VB.ComboBox BS000_03 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   7
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   1890
               Width           =   2265
            End
            Begin VB.ComboBox BS000_03 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   6
               ItemData        =   "frmBancoSangre.frx":2E7E
               Left            =   1920
               List            =   "frmBancoSangre.frx":2E91
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   1530
               Width           =   2265
            End
            Begin VB.ComboBox BS000_03 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   4
               ItemData        =   "frmBancoSangre.frx":2ED6
               Left            =   1920
               List            =   "frmBancoSangre.frx":2EE0
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   1200
               Width           =   2265
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   7
               Left            =   8370
               TabIndex        =   19
               Top             =   1620
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   3
               Left            =   1920
               TabIndex        =   7
               Top             =   585
               Width           =   4155
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   2
               Left            =   1920
               TabIndex        =   6
               Top             =   255
               Width           =   4155
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   4
               Left            =   1920
               TabIndex        =   8
               Top             =   915
               Width           =   4155
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   4590
               TabIndex        =   12
               Top             =   1890
               Width           =   1485
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   5
               Left            =   8370
               TabIndex        =   17
               Top             =   975
               Width           =   3915
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   8
               Left            =   1920
               TabIndex        =   14
               Top             =   2550
               Width           =   4155
            End
            Begin MSMask.MaskEdBox BS000_02 
               Height          =   315
               Index           =   1
               Left            =   8370
               TabIndex        =   15
               Top             =   315
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
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
               PromptChar      =   " "
            End
            Begin VB.Label BS000_00 
               Caption         =   "Nº"
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
               Index           =   74
               Left            =   4365
               TabIndex        =   153
               Top             =   1950
               Width           =   195
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Observaciones"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   64
               Left            =   7020
               TabIndex        =   152
               Top             =   1980
               Width           =   1320
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ocupación"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   56
               Left            =   7020
               TabIndex        =   151
               Top             =   675
               Width           =   1320
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Lugar Procedencia"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   60
               Left            =   7020
               TabIndex        =   150
               Top             =   1320
               Width           =   1320
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Documento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   63
               Left            =   180
               TabIndex        =   149
               Top             =   1890
               Width           =   960
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sexo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   59
               Left            =   180
               TabIndex        =   148
               Top             =   1230
               Width           =   405
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Teléfono"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   62
               Left            =   7020
               TabIndex        =   147
               Top             =   1650
               Width           =   1320
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Lugar Nacimiento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   65
               Left            =   180
               TabIndex        =   146
               Top             =   2220
               Width           =   1410
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Estado Civil"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   61
               Left            =   180
               TabIndex        =   145
               Top             =   1560
               Width           =   900
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Nacimiento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   7020
               TabIndex        =   144
               Top             =   345
               Width           =   1320
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Apellido Materno"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   55
               Left            =   180
               TabIndex        =   143
               Top             =   585
               Width           =   1365
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nombres"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   57
               Left            =   180
               TabIndex        =   142
               Top             =   915
               Width           =   720
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Apellido Paterno"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   6
               Left            =   180
               TabIndex        =   141
               Top             =   255
               Width           =   1335
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Centro de Trabajo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   58
               Left            =   7020
               TabIndex        =   140
               Top             =   1005
               Width           =   1320
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Dirección"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   66
               Left            =   180
               TabIndex        =   139
               Top             =   2550
               Width           =   735
            End
         End
         Begin VB.Frame BS000_05 
            Caption         =   "Datos Personales del Postulante"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1600
            Index           =   1
            Left            =   240
            TabIndex        =   209
            Top             =   3630
            Visible         =   0   'False
            Width           =   12435
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   42
               Left            =   8370
               TabIndex        =   28
               Top             =   915
               Width           =   2475
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   15
               Left            =   1920
               TabIndex        =   25
               Top             =   1260
               Width           =   4155
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   13
               Left            =   8370
               TabIndex        =   27
               Top             =   585
               Width           =   1515
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   12
               Left            =   1920
               TabIndex        =   23
               Top             =   615
               Width           =   4155
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   10
               Left            =   1920
               TabIndex        =   22
               Top             =   285
               Width           =   4155
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   14
               Left            =   1920
               TabIndex        =   24
               Top             =   945
               Width           =   4155
            End
            Begin VB.TextBox BS000_01 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   11
               Left            =   8370
               TabIndex        =   26
               Top             =   255
               Width           =   1515
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Grado de Parentesco"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   73
               Left            =   180
               TabIndex        =   216
               Top             =   1290
               Width           =   1725
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Cama"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   70
               Left            =   7875
               TabIndex        =   215
               Top             =   615
               Width           =   435
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Atención"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   72
               Left            =   7020
               TabIndex        =   214
               Top             =   945
               Width           =   1320
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Sala Hospitalización"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   69
               Left            =   180
               TabIndex        =   213
               Top             =   645
               Width           =   1530
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   71
               Left            =   180
               TabIndex        =   212
               Top             =   975
               Width           =   930
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre de Receptor"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   67
               Left            =   180
               TabIndex        =   211
               Top             =   315
               Width           =   1725
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Historia Clínica"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   68
               Left            =   7020
               TabIndex        =   210
               Top             =   285
               Width           =   1320
            End
         End
         Begin VB.Shape BS000_06 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   6300
            Index           =   1
            Left            =   -74925
            Top             =   360
            Width           =   12855
         End
         Begin VB.Shape BS000_06 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   6300
            Index           =   0
            Left            =   75
            Top             =   360
            Width           =   12855
         End
         Begin VB.Shape BS000_06 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   6300
            Index           =   2
            Left            =   -74925
            Top             =   360
            Width           =   12855
         End
         Begin VB.Shape BS000_06 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   6300
            Index           =   3
            Left            =   -74925
            Top             =   360
            Width           =   12855
         End
      End
      Begin MSMask.MaskEdBox BS000_02 
         Height          =   315
         Index           =   0
         Left            =   9480
         TabIndex        =   2
         Top             =   240
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         HideSelection   =   0   'False
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
         PromptChar      =   " "
      End
      Begin VB.Label BS000_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Postulante"
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
         Index           =   1
         Left            =   3840
         TabIndex        =   222
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label BS000_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Registro"
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
         Index           =   2
         Left            =   7875
         TabIndex        =   221
         Top             =   270
         Width           =   1470
      End
      Begin VB.Label BS000_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Donante"
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
         Index           =   0
         Left            =   120
         TabIndex        =   220
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label BS000_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo postulante"
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
         Index           =   5
         Left            =   8055
         TabIndex        =   219
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label BS000_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factor RH"
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
         Index           =   4
         Left            =   4800
         TabIndex        =   218
         Top             =   600
         Width           =   795
      End
      Begin VB.Label BS000_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Sanguíneo"
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
         Index           =   3
         Left            =   120
         TabIndex        =   217
         Top             =   600
         Width           =   1410
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1695
      Left            =   0
      TabIndex        =   232
      Top             =   600
      Visible         =   0   'False
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2990
   End
   Begin VB.Frame INM006 
      Caption         =   "ELISA HBsAg (Antígeno Australiano)"
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
      Height          =   1110
      Left            =   0
      TabIndex        =   223
      Top             =   600
      Visible         =   0   'False
      Width           =   7220
      Begin VB.TextBox INM006_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   480
         TabIndex        =   227
         Top             =   435
         Width           =   1215
      End
      Begin VB.TextBox INM006_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5940
         TabIndex        =   226
         Text            =   "E.I.A."
         Top             =   435
         Width           =   1095
      End
      Begin VB.TextBox INM006_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   225
         Text            =   "Positivo si es > de 0.254"
         Top             =   435
         Width           =   2535
      End
      Begin VB.TextBox INM006_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1365
         TabIndex        =   224
         Top             =   750
         Width           =   5670
      End
      Begin VB.Label INM006_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Height          =   255
         Index           =   2
         Left            =   5760
         TabIndex        =   231
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label INM006_00 
         Alignment       =   2  'Center
         Caption         =   "Valor de Referencia"
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
         Height          =   255
         Index           =   1
         Left            =   3060
         TabIndex        =   230
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label INM006_00 
         Caption         =   "Observaciones"
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
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   229
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM006_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   228
         Top             =   225
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmBancoSangre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultados de Banco de Sangre
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable

Dim mo_BS000_03_6 As New sighentidades.ListaDespleglable
Dim mo_BS000_03_7 As New sighentidades.ListaDespleglable
Dim mo_BS000_03_3 As New sighentidades.ListaDespleglable

Dim ml_idUsuario As Long
Dim ml_idOrden As Long
Dim ml_nombrePrueba As String
Dim ml_idAnalisis As Long
Dim ml_idPaciente As Long
Dim ml_resultado As String
Dim ml_observacion As String
Dim ml_IdMovimiento As Long
Dim ms_MensajeError As String
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

Property Let DetalleOrden(lValue As ADODB.Recordset)
  Set ml_DetalleOrden = lValue
End Property

Property Let CodigoPruebaSeleccionada(lValue As String)
   ml_CodigoPruebaSeleccionada = lValue
End Property

Sub CargaDataCombos()
  mo_cmbResponsable.BoundColumn = "idEmpleado"
  mo_cmbResponsable.ListField = "ApNom"
  'Set mo_cmbResponsable.RowSource = mo_ReglasLaboratorio.EmpleadosDeLab(ml_areaTrabajo)
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =19")

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

Sub SeleccionaTexto(T As TextBox)
  T.SelStart = 0
  T.SelLength = Len(T.Text)
  'T.BackColor = &HC0FFFF
  'T.BackColor = &HFFFFFF'blanco
End Sub

Sub SeleccionaMask(M As MaskEdBox)
  M.SelStart = 0
  M.SelLength = Len(M.Text)
  'm.BackColor = &HC0FFFF
  'm.BackColor = &HFFFFFF'blanco
End Sub

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

Public Function Inicializar()
  Set mo_BS000_03_6.MiComboBox = BS000_03(6)
  Set mo_BS000_03_7.MiComboBox = BS000_03(7)
  Set mo_BS000_03_3.MiComboBox = BS000_03(3)
End Function

Public Sub ConfigurarComboBoxes()
  mo_BS000_03_3.BoundColumn = "IdTipoOcupacion"
  mo_BS000_03_3.ListField = "DescripcionLarga"
  Set mo_BS000_03_3.RowSource = mo_AdminServiciosComunes.TiposOcupacionSeleccionarTodos()
  
  mo_BS000_03_6.BoundColumn = "IdEstadoCivil"
  mo_BS000_03_6.ListField = "DescripcionLarga"
  Set mo_BS000_03_6.RowSource = mo_AdminServiciosComunes.TiposEstadoCivilSeleccionarTodos()
  
  mo_BS000_03_7.BoundColumn = "IdDocIdentidad"
  mo_BS000_03_7.ListField = "DescripcionLarga"
  Set mo_BS000_03_7.RowSource = mo_AdminServiciosComunes.TiposDocIdentidadSeleccionarTodos()
        
End Sub

Private Function Ubica_En_Combo(C As ComboBox, Co As String) As Integer
  Ubica_En_Combo = -1
  Dim Z, Y As Integer
  Z = C.ListCount
  For Y = 0 To Z - 1
    If C.List(Y) = Co Then
      Ubica_En_Combo = Y
      Exit For
    End If
  Next Y
End Function

Private Sub BS000_01_GotFocus(Index As Integer)
  SeleccionaTexto BS000_01(Index)
End Sub

Private Sub BS000_01_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  'mo_Teclado.RealizarNavegacion  KeyCode, BS000_01(Index)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BS000_02_GotFocus(Index As Integer)
  SeleccionaMask BS000_02(Index)
End Sub

Private Sub BS000_02_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
  'mo_Teclado.RealizarNavegacion KeyCode, BS000_02(Index)
End Sub

Private Sub BS000_03_Click(Index As Integer)
  If Index = 2 Then
    If BS000_03(2).Text = "Reposición" Then
      BS000_07.Visible = True
    Else
      BS000_07.Visible = False
    End If
  End If
  If Index = 4 Then
    If BS000_03(4).Text = "Femenino" Then
      BS000_05(5).Visible = True
    Else
      BS000_05(5).Visible = False
    End If
  End If
End Sub

Private Sub BS000_03_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BS000_07_Click()
  BS000_05(1).Visible = BS000_07.Value
End Sub

Private Sub BS000_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub chkCGR_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub chkEnf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub chkPGR_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
  If Me.txtFresultado.Text = sighentidades.FECHA_VACIA_DMY Then
    MsgBox "Por favor ingresar la Fecha del Resultado", vbInformation, "SIGH "
    Exit Sub
  End If
  ml_nombreRealiza = mo_cmbResponsable.BoundText
  If ml_CodigoPruebaSeleccionada = "BSA001" Then
    ml_resultado = BS000_01(0).Text & "\" & BS000_01(1).Text & "\" & BS000_02(0).Text & "\" & BS000_03(0).Text & "\" & BS000_03(1).Text & "\" & BS000_03(2).Text & "\" & _
                   BS000_01(2).Text & "\" & BS000_01(3).Text & "\" & BS000_01(4).Text & "\" & BS000_03(4).Text & "\" & BS000_03(6).Text & "\" & BS000_03(7).Text & "\" & BS000_01(6).Text & "\" & BS000_01(40).Text & "\" & BS000_01(8).Text & "\" & BS000_02(1).Text & "\" & BS000_03(3).Text & "\" & BS000_01(5).Text & "\" & BS000_01(41).Text & "\" & BS000_01(7).Text & "\" & BS000_01(9).Text & "\" & _
                   BS000_07.Value & "\" & BS000_01(10).Text & "\" & BS000_01(12).Text & "\" & BS000_01(14).Text & "\" & BS000_01(15).Text & "\" & BS000_01(11).Text & "\" & BS000_01(13).Text & "\" & BS000_01(42).Text & "\" & _
                   optSi(0).Value & "\" & optNo(0).Value & "\" & optSi(1).Value & "\" & optNo(1).Value & "\" & optSi(2).Value & "\" & optNo(2).Value & "\" & optSi(3).Value & "\" & optNo(3).Value & "\" & optSi(4).Value & "\" & optNo(4).Value & "\" & optSi(5).Value & "\" & optNo(5).Value & "\" & optSi(6).Value & "\" & optNo(6).Value & "\" & optSi(7).Value & "\" & optNo(7).Value & "\" & optSi(8).Value & "\" & optNo(8).Value & "\" & optSi(9).Value & "\" & optNo(9).Value & "\" & optSi(10).Value & "\" & optNo(10).Value & "\" & optSi(12).Value & "\" & optNo(12).Value & "\" & optSi(13).Value & "\" & optNo(13).Value & "\" & optSi(16).Value & "\" & optNo(16).Value & "\" & optSi(17).Value & "\" & optNo(17).Value & "\" & optSi(18).Value & "\" & optNo(18).Value & "\" & _
                   chkEnf(0).Value & "\" & chkEnf(1).Value & "\" & chkEnf(2).Value & "\" & chkEnf(3).Value & "\" & chkEnf(4).Value & "\" & chkEnf(5).Value & "\" & chkEnf(6).Value & "\" & chkEnf(7).Value & "\" & chkEnf(8).Value & "\" & chkEnf(9).Value & "\" & chkEnf(10).Value & "\" & chkEnf(11).Value & "\" & chkEnf(12).Value & "\" & chkEnf(13).Value & "\" & chkEnf(14).Value & "\" & chkEnf(15).Value & "\" & chkEnf(16).Value & "\" & chkEnf(17).Value & "\" & chkEnf(18).Value & "\" & chkEnf(19).Value & "\" & chkEnf(20).Value & "\" & chkEnf(21).Value & "\" & chkEnf(22).Value & "\" & chkEnf(23).Value & "\" & optSi(11).Value & "\" & optNo(11).Value & "\" & _
                   chkPGR(0).Value & "\" & chkPGR(1).Value & "\" & chkPGR(2).Value & "\" & chkPGR(3).Value & "\" & optSi(14).Value & "\" & optNo(14).Value & "\" & _
                   chkCGR(0).Value & "\" & chkCGR(1).Value & "\" & chkCGR(2).Value & "\" & chkCGR(3).Value & "\" & optSi(15).Value & "\" & optNo(15).Value & "\" & _
                   BS000_01(16).Text & "\" & BS000_01(17).Text & "\" & optMens(0).Value & "\" & optMens(1).Value & "\" & optMens(2).Value & "\" & optSi(19).Value & "\" & optNo(19).Value & "\" & BS000_01(18).Text & "\" & optSi(20).Value & "\" & optNo(20).Value & "\" & _
                   BS000_01(19).Text & "\" & BS000_01(20).Text & "\" & BS000_01(21).Text & "\" & BS000_01(22).Text & "\" & BS000_01(23).Text & "\" & BS000_01(24).Text & "\" & _
                   BS000_01(25).Text & "\" & BS000_01(27).Text & "\" & BS000_01(29).Text & "\" & BS000_01(31).Text & "\" & BS000_01(33).Text & "\" & BS000_01(35).Text & "\" & BS000_01(37).Text & "\" & BS000_01(39).Text & "\" & BS000_01(26).Text & "\" & BS000_01(28).Text & "\" & BS000_01(30).Text & "\" & BS000_01(32).Text & "\" & BS000_01(34).Text & "\" & BS000_01(36).Text & "\" & BS000_01(38).Text & "\" & _
                   optApto(0).Value & "\" & optApto(1).Value & "\" & optApto(2).Value
    ml_observacion = ""
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  mo_ReglasLaboratorio.LabIngresaResultados idPrueba, idOrden, ml_resultado, ml_observacion, idUsuario, ml_nombreRealiza, ml_DetalleOrden, ml_idOrdenLab, BS000_03(0).Text, BS000_03(1).Text, ml_idPaciente, CDate(Me.txtFresultado.Text), mo_lcNombrePc, mo_lnIdTablaLISTBARITEMS, Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre, Me.Caption
End Sub


Private Sub cmdImprimir_Click()
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_CodigoPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
    'mo_ReglasLaboratorio.LabImprimeCabeceraResultados UcPacienteDatos1.idPaciente, nombrePaciente, UcPacienteDatos1.NroHistoriaClinica, ldFechaResultado, nombreMedico
    mo_ReglasLaboratorio.LabImprimeResultadosBS ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
    mo_ReglasLaboratorio.LabImprimePieResultados
  Else
    MsgBox "Debe grabar los resultados antes de poder imprimirlos", vbInformation, ""
  End If
End Sub

Private Sub Form_Initialize()
  Set mo_cmbResponsable.MiComboBox = cmbResponsable
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
  Me.txtFresultado.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  Dim oRsTmp As New Recordset
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
  
  If ml_CodigoPruebaSeleccionada = "BSA001" Then  'Banco de Sangre
    'TopBoton BQM001
    BS000_01(2).Text = UcPacienteDatos1.APat
    BS000_01(3).Text = UcPacienteDatos1.AMat
    BS000_01(4).Text = UcPacienteDatos1.Nombre
    BS000_02(1).Text = UcPacienteDatos1.FechaNacimiento
    BS000_03(4).ListIndex = Ubica_En_Combo(BS000_03(4), UcPacienteDatos1.Sexo)
    
    BS000_04.Tab = 0
    Call Inicializar
    Call ConfigurarComboBoxes
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  'Recupera información si es que ya esta grabado
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_resultado = "" Or Val(ml_nombreRealiza) = 0 Then
     Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(ml_idPaciente)
     If oRsTmp.RecordCount > 0 Then
        If Not IsNull(oRsTmp.Fields!grupoSanguineo) Then
           BS000_03(0).Text = oRsTmp.Fields!grupoSanguineo
        End If
        If Not IsNull(oRsTmp.Fields!factorRh) Then
          BS000_03(1).Text = oRsTmp.Fields!factorRh
        End If
     End If
     oRsTmp.Close
     Set oRsTmp = Nothing
     Exit Sub
  End If
  cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, mo_ReglasLaboratorio.LabEmpleado(ml_nombreRealiza))
  Me.txtFresultado.Text = Format(IIf(ldFechaResultado = 0, Now, ldFechaResultado), sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  'If cmbResponsable.Text <> "" Then cmbResponsable.Enabled = False
  Dim Temp As String
  'Asigna la información recuperada en el formulario
  If ml_CodigoPruebaSeleccionada = "BSA001" Then 'Banco de Sangre
    'MsgBox 1
    On Error Resume Next
    BS000_01(0).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(1).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_02(0).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_03(0).ListIndex = Ubica_En_Combo(BS000_03(0), Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_03(1).ListIndex = Ubica_En_Combo(BS000_03(1), Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_03(2).ListIndex = Ubica_En_Combo(BS000_03(2), Temp)
    BS000_01(2).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(3).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(4).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_03(4).ListIndex = Ubica_En_Combo(BS000_03(4), Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_03(6).ListIndex = Ubica_En_Combo(BS000_03(6), Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_03(7).ListIndex = Ubica_En_Combo(BS000_03(7), Temp)
    BS000_01(6).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(40).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(8).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_02(1).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_03(3).ListIndex = Ubica_En_Combo(BS000_03(3), Temp)
    BS000_01(5).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(41).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(7).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(9).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_07.Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(10).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(12).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(14).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(15).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(11).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(13).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(42).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(6).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(6).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(7).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(7).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(8).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(8).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(9).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(9).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(10).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(10).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(12).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(12).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(13).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(13).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(16).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(16).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(17).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(17).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(18).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(18).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(6).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(7).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(8).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(9).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(10).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(11).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(12).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(13).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(14).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(15).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(16).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(17).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(18).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(19).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(20).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(21).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(22).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkEnf(23).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(11).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(11).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkPGR(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkPGR(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkPGR(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkPGR(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(14).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(14).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkCGR(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkCGR(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkCGR(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    chkCGR(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(15).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(15).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(16).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(17).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optMens(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optMens(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optMens(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(19).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(19).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(18).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optSi(20).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optNo(20).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(19).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(20).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(21).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(22).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(23).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(24).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(25).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(27).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(29).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(31).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(33).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(35).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(37).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(39).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(26).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(28).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(30).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(32).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(34).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(36).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BS000_01(38).Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optApto(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optApto(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    optApto(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    '
    Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(ml_idPaciente)
    If oRsTmp.RecordCount > 0 Then
       If Not IsNull(oRsTmp.Fields!grupoSanguineo) Then
          BS000_03(0).Text = oRsTmp.Fields!grupoSanguineo
       End If
       If Not IsNull(oRsTmp.Fields!factorRh) Then
          BS000_03(1).Text = oRsTmp.Fields!factorRh
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    '
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
  End If
End Sub

Private Sub optApto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub optMens_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub optNo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub optSi_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub
