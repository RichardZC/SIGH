VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPatologia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PATOLOGÍA QUIRÚRGICA"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "frmPatologia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   60
      TabIndex        =   45
      Top             =   1680
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
         TabIndex        =   46
         Top             =   180
         Width           =   3120
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   5580
         TabIndex        =   47
         Top             =   210
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
         Left            =   120
         TabIndex        =   49
         Top             =   240
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
         Left            =   4590
         TabIndex        =   48
         Top             =   255
         Width           =   945
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1815
      Left            =   60
      TabIndex        =   38
      Top             =   15
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   90
      TabIndex        =   39
      Top             =   7380
      Width           =   7095
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmPatologia.frx":0CCA
         DownPicture     =   "frmPatologia.frx":118E
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
         Left            =   3690
         Picture         =   "frmPatologia.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   42
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
         Left            =   90
         Picture         =   "frmPatologia.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmPatologia.frx":203F
         DownPicture     =   "frmPatologia.frx":249F
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
         Left            =   2220
         Picture         =   "frmPatologia.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Frame PAQ002 
      Caption         =   "Líquidos con Block Cell"
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
      TabIndex        =   26
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin TabDlg.SSTab PAQ002_00 
         Height          =   4575
         Left            =   60
         TabIndex        =   44
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
         TabPicture(0)   =   "frmPatologia.frx":2D89
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "CPA001_01(16)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "PAQ002_01(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "PAQ002_01(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "PAQ002_01(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "PAQ002_01(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "PAQ002_05"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "PAQ002_04"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "PAQ002_03"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "PAQ002_02"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Resultado"
         TabPicture(1)   =   "frmPatologia.frx":2DA5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "PAQ002_06"
         Tab(1).Control(1)=   "PAQ002_07"
         Tab(1).Control(2)=   "PAQ002_08"
         Tab(1).Control(3)=   "PAQ002_01(4)"
         Tab(1).Control(4)=   "PAQ002_01(5)"
         Tab(1).Control(5)=   "PAQ002_01(6)"
         Tab(1).Control(6)=   "CPA001_01(0)"
         Tab(1).ControlCount=   7
         Begin VB.TextBox PAQ002_02 
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
            Left            =   1260
            MaxLength       =   35
            TabIndex        =   9
            Top             =   420
            Width           =   2595
         End
         Begin VB.TextBox PAQ002_03 
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
            TabIndex        =   10
            Top             =   1050
            Width           =   6675
         End
         Begin VB.TextBox PAQ002_04 
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
            TabIndex        =   11
            Top             =   2250
            Width           =   6675
         End
         Begin VB.TextBox PAQ002_05 
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
            TabIndex        =   12
            Top             =   3580
            Width           =   6675
         End
         Begin VB.TextBox PAQ002_06 
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
            MaxLength       =   35
            TabIndex        =   13
            Top             =   420
            Width           =   2595
         End
         Begin VB.TextBox PAQ002_07 
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
            TabIndex        =   14
            Top             =   990
            Width           =   6675
         End
         Begin VB.TextBox PAQ002_08 
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
            TabIndex        =   15
            Top             =   2160
            Width           =   6675
         End
         Begin VB.Label PAQ002_01 
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
            TabIndex        =   33
            Top             =   450
            Width           =   1065
         End
         Begin VB.Label PAQ002_01 
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
            TabIndex        =   32
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label PAQ002_01 
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
            TabIndex        =   31
            Top             =   2040
            Width           =   2130
         End
         Begin VB.Label PAQ002_01 
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
            TabIndex        =   30
            Top             =   3360
            Width           =   1005
         End
         Begin VB.Label PAQ002_01 
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
            TabIndex        =   29
            Top             =   450
            Width           =   2760
         End
         Begin VB.Label PAQ002_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Impresión Diagnóstica"
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
            TabIndex        =   28
            Top             =   780
            Width           =   1575
         End
         Begin VB.Label PAQ002_01 
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
            TabIndex        =   27
            Top             =   1950
            Width           =   1695
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
            TabIndex        =   34
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
            Index           =   0
            Left            =   -74940
            TabIndex        =   35
            Top             =   360
            Width           =   6825
         End
      End
   End
   Begin VB.Frame PAQ001 
      Caption         =   "Pieza Operatoria"
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
      TabIndex        =   16
      Top             =   2430
      Visible         =   0   'False
      Width           =   7185
      Begin TabDlg.SSTab PAQ001_00 
         Height          =   4575
         Left            =   60
         TabIndex        =   43
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
         TabPicture(0)   =   "frmPatologia.frx":2DC1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "CPA001_01(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "PAQ001_01(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "PAQ001_01(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "PAQ001_01(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "PAQ001_01(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "PAQ001_05"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "PAQ001_04"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "PAQ001_03"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "PAQ001_02"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Resultado"
         TabPicture(1)   =   "frmPatologia.frx":2DDD
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "PAQ001_06"
         Tab(1).Control(1)=   "PAQ001_07"
         Tab(1).Control(2)=   "PAQ001_08"
         Tab(1).Control(3)=   "PAQ001_09"
         Tab(1).Control(4)=   "PAQ001_10"
         Tab(1).Control(5)=   "PAQ001_01(4)"
         Tab(1).Control(6)=   "PAQ001_01(5)"
         Tab(1).Control(7)=   "PAQ001_01(6)"
         Tab(1).Control(8)=   "PAQ001_01(7)"
         Tab(1).Control(9)=   "PAQ001_01(8)"
         Tab(1).Control(10)=   "CPA001_01(2)"
         Tab(1).ControlCount=   11
         Begin VB.TextBox PAQ001_02 
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
            Left            =   1260
            MaxLength       =   35
            TabIndex        =   0
            Top             =   420
            Width           =   2595
         End
         Begin VB.TextBox PAQ001_03 
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
            TabIndex        =   1
            Top             =   1050
            Width           =   6675
         End
         Begin VB.TextBox PAQ001_04 
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
            TabIndex        =   2
            Top             =   2250
            Width           =   6675
         End
         Begin VB.TextBox PAQ001_05 
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
            TabIndex        =   3
            Top             =   3580
            Width           =   6675
         End
         Begin VB.TextBox PAQ001_06 
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
            MaxLength       =   35
            TabIndex        =   4
            Top             =   420
            Width           =   2595
         End
         Begin VB.TextBox PAQ001_07 
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
            TabIndex        =   5
            Top             =   870
            Width           =   6675
         End
         Begin VB.TextBox PAQ001_08 
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
            TabIndex        =   6
            Top             =   1920
            Width           =   6675
         End
         Begin VB.TextBox PAQ001_09 
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
            TabIndex        =   7
            Top             =   3030
            Width           =   6675
         End
         Begin VB.TextBox PAQ001_10 
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
            Left            =   -74850
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   4080
            Width           =   6675
         End
         Begin VB.Label PAQ001_01 
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
            TabIndex        =   25
            Top             =   450
            Width           =   1065
         End
         Begin VB.Label PAQ001_01 
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
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   2130
         End
         Begin VB.Label PAQ001_01 
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
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label PAQ001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Muestra o Especimen motivo de Estudio"
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
            TabIndex        =   22
            Top             =   3360
            Width           =   2835
         End
         Begin VB.Label PAQ001_01 
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
            TabIndex        =   21
            Top             =   450
            Width           =   2760
         End
         Begin VB.Label PAQ001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Éxamen Clínico"
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
            TabIndex        =   20
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label PAQ001_01 
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
            TabIndex        =   19
            Top             =   1710
            Width           =   1695
         End
         Begin VB.Label PAQ001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hallazgos operatorios"
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
            Left            =   -74970
            TabIndex        =   18
            Top             =   2790
            Width           =   1545
         End
         Begin VB.Label PAQ001_01 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Impresión Diagnóstica"
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
            Left            =   -74940
            TabIndex        =   17
            Top             =   3870
            Width           =   1575
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
            Index           =   1
            Left            =   60
            TabIndex        =   36
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
            Index           =   2
            Left            =   -74940
            TabIndex        =   37
            Top             =   360
            Width           =   6825
         End
      End
   End
End
Attribute VB_Name = "frmPatologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultado para patología Quirúrgica
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
  'Set mo_cmbResponsable.RowSource = mo_ReglasLaboratorio.EmpleadosDeLab(ml_areaTrabajo)
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =18")
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
  Fra.Top = UcPacienteDatos1.Top + UcPacienteDatos1.Height + 600 '350
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
  If Me.txtFresultado.Text = sighentidades.FECHA_VACIA_DMY Then
    MsgBox "Por favor ingresar la Fecha del Resultado", vbInformation, "SIGH "
    Exit Sub
  End If
  ml_nombreRealiza = mo_cmbResponsable.BoundText
  If ml_CodigoPruebaSeleccionada = "PAQ001" Then  'Pieza operatoria mayor, Biopsia Quirurgica, Pieza operatoria Mediana, Pieza Operatoria pequeña
    'PAQ001
    ml_resultado = PAQ001_02 & "\" & PAQ001_03 & "\" & PAQ001_04 & "\" & PAQ001_05 & "\" & PAQ001_06 & "\" & PAQ001_07 & "\" & PAQ001_08 & "\" & PAQ001_09 & "\" & PAQ001_10
  ElseIf ml_CodigoPruebaSeleccionada = "PAQ002" Then 'Líquidos con BlockCell
    'PAQ002
    ml_resultado = PAQ002_02 & "\" & PAQ002_03 & "\" & PAQ002_04 & "\" & PAQ002_05 & "\" & PAQ002_06 & "\" & PAQ002_07 & "\" & PAQ002_08
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  mo_ReglasLaboratorio.LabIngresaResultados idPrueba, idOrden, ml_resultado, ml_observacion, idUsuario, ml_nombreRealiza, ml_DetalleOrden, ml_idOrdenLab, "", "", 0, CDate(Me.txtFresultado.Text), mo_lcNombrePc, mo_lnIdTablaLISTBARITEMS, Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre, Me.Caption
End Sub

Private Sub cmdImprimir_Click()
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_CodigoPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
    mo_ReglasLaboratorio.LabImprimeCabeceraResultados UcPacienteDatos1.idPaciente, nombrePaciente, UcPacienteDatos1.NroHistoriaClinica, ldFechaResultado, _
                         nombreMedico + mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(ml_idOrden)
    mo_ReglasLaboratorio.LabImprimeResultadosPAQ ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
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
  If ml_CodigoPruebaSeleccionada = "PAQ001" Then  'Pieza operatoria mayor, Biopsia Quirurgica, Pieza operatoria Mediana, Pieza Operatoria pequeña
    TopBoton PAQ001
  ElseIf ml_CodigoPruebaSeleccionada = "PAQ002" Then 'Líquidos con BlockCell
    TopBoton PAQ002
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
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
  If ml_CodigoPruebaSeleccionada = "PAQ001" Then  'Pieza operatoria mayor, Biopsia Quirurgica, Pieza operatoria Mediana, Pieza Operatoria pequeña
    'PAQ001
    PAQ001_02 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ001_03 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ001_04 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ001_05 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ001_06 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ001_07 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ001_08 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ001_09 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ001_10 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
  ElseIf ml_CodigoPruebaSeleccionada = "PAQ002" Then 'Líquidos con BlockCell
    'PAQ002
    PAQ002_02 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ002_03 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ002_04 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ002_05 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ002_06 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ002_07 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    PAQ002_08 = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
End Sub


Private Sub PAQ001_02_GotFocus()
  SeleccionaTexto PAQ001_02
End Sub

Private Sub PAQ001_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ001_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ001_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ001_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ001_06_GotFocus()
  SeleccionaTexto PAQ001_06
End Sub

Private Sub PAQ001_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ001_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ001_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ001_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ001_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ002_02_GotFocus()
  SeleccionaTexto PAQ002_02
End Sub

Private Sub PAQ002_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ002_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ002_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ002_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ002_06_GotFocus()
  SeleccionaTexto PAQ002_06
End Sub

Private Sub PAQ002_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ002_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub PAQ002_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

