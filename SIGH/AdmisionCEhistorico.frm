VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AdmisionCEhistorico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Históricos en Consultorios Externos"
   ClientHeight    =   9720
   ClientLeft      =   1125
   ClientTop       =   4545
   ClientWidth     =   13170
   Icon            =   "AdmisionCEhistorico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   8940
      Width           =   13125
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar "
         DisabledPicture =   "AdmisionCEhistorico.frx":000C
         DownPicture     =   "AdmisionCEhistorico.frx":04D0
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   6656
         Picture         =   "AdmisionCEhistorico.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton btnImprime 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5299
         Picture         =   "AdmisionCEhistorico.frx":0EA8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13125
      Begin TabDlg.SSTab tabModulos 
         Height          =   8745
         Left            =   60
         TabIndex        =   6
         Top             =   135
         Width           =   13005
         _ExtentX        =   22939
         _ExtentY        =   15425
         _Version        =   393216
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
         TabCaption(0)   =   "HC Consulta Externa"
         TabPicture(0)   =   "AdmisionCEhistorico.frx":1381
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSTabResultados"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Atención integral Niño Sano"
         TabPicture(1)   =   "AdmisionCEhistorico.frx":139D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ucPerinatalAS1"
         Tab(1).Control(1)=   "grdControles"
         Tab(1).Control(2)=   "lblControl"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Historico de atenciones"
         TabPicture(2)   =   "AdmisionCEhistorico.frx":13B9
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ucHCelectronicaLista1"
         Tab(2).ControlCount=   1
         Begin TabDlg.SSTab SSTabResultados 
            Height          =   8295
            Left            =   60
            TabIndex        =   11
            Top             =   345
            Width           =   12915
            _ExtentX        =   22781
            _ExtentY        =   14631
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
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
            TabCaption(0)   =   "Todos"
            TabPicture(0)   =   "AdmisionCEhistorico.frx":13D5
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grdAnteriores"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "TabHistoricos"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Solo con Resultados (cpt de Laboratorio/Imágenes)"
            TabPicture(1)   =   "AdmisionCEhistorico.frx":13F1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label4"
            Tab(1).Control(1)=   "grdApoyoDx"
            Tab(1).Control(2)=   "chkSoloConResultados"
            Tab(1).ControlCount=   3
            Begin VB.CheckBox chkSoloConResultados 
               Caption         =   "Solo con Resultados"
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
               Left            =   -64545
               TabIndex        =   34
               Top             =   480
               Value           =   1  'Checked
               Width           =   2220
            End
            Begin TabDlg.SSTab TabHistoricos 
               Height          =   5745
               Left            =   30
               TabIndex        =   12
               Top             =   2430
               Width           =   12840
               _ExtentX        =   22648
               _ExtentY        =   10134
               _Version        =   393216
               Tabs            =   5
               TabsPerRow      =   5
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
               TabCaption(0)   =   "Cita"
               TabPicture(0)   =   "AdmisionCEhistorico.frx":140D
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "lblObservaciones"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "txtCitaObservaciones"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "lblExamenes"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "txtCitaExClinicos"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "lblTratamiento"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "TxtCitaTratamiento"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "lblDx"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "txtDx"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "lblExamen"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "txtCitaExamenClinico"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "lblMotivo"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "txtCitaMotivo"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).ControlCount=   12
               TabCaption(1)   =   "Imágenes (exámenes)"
               TabPicture(1)   =   "AdmisionCEhistorico.frx":1429
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Label3"
               Tab(1).Control(1)=   "grdImagenes"
               Tab(1).ControlCount=   2
               TabCaption(2)   =   "Farmacia (despachos)"
               TabPicture(2)   =   "AdmisionCEhistorico.frx":1445
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "grdFarmacia"
               Tab(2).ControlCount=   1
               TabCaption(3)   =   "Laboratorio (exámenes)"
               TabPicture(3)   =   "AdmisionCEhistorico.frx":1461
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Label2"
               Tab(3).Control(1)=   "grdLaboratorio"
               Tab(3).ControlCount=   2
               TabCaption(4)   =   "Otros CPT"
               TabPicture(4)   =   "AdmisionCEhistorico.frx":147D
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "grdOtrosCpt"
               Tab(4).ControlCount=   1
               Begin VB.TextBox txtCitaMotivo 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1770
                  Left            =   90
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   24
                  Top             =   750
                  Width           =   4170
               End
               Begin VB.TextBox lblMotivo 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   90
                  TabIndex        =   23
                  Text            =   "Motivo de la Consulta"
                  Top             =   390
                  Width           =   4155
               End
               Begin VB.TextBox txtCitaExamenClinico 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1770
                  Left            =   4260
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   22
                  Top             =   780
                  Width           =   4170
               End
               Begin VB.TextBox lblExamen 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   4275
                  TabIndex        =   21
                  Text            =   "Exámen Clínico"
                  Top             =   390
                  Width           =   4155
               End
               Begin VB.TextBox txtDx 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1770
                  Left            =   8475
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   20
                  Top             =   780
                  Width           =   4170
               End
               Begin VB.TextBox lblDx 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   8460
                  TabIndex        =   19
                  Text            =   "Diagnóstico"
                  Top             =   390
                  Width           =   4155
               End
               Begin VB.TextBox TxtCitaTratamiento 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1920
                  Left            =   75
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   18
                  Top             =   2910
                  Width           =   4170
               End
               Begin VB.TextBox lblTratamiento 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   75
                  TabIndex        =   17
                  Text            =   "Tratamiento"
                  Top             =   2565
                  Width           =   4155
               End
               Begin VB.TextBox txtCitaExClinicos 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1920
                  Left            =   4260
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   16
                  Top             =   2910
                  Width           =   4170
               End
               Begin VB.TextBox lblExamenes 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   4275
                  TabIndex        =   15
                  Text            =   "Servicios de apoyo al Dx"
                  Top             =   2565
                  Width           =   4155
               End
               Begin VB.TextBox txtCitaObservaciones 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1920
                  Left            =   8475
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   14
                  Top             =   2910
                  Width           =   4170
               End
               Begin VB.TextBox lblObservaciones 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   8460
                  TabIndex        =   13
                  Text            =   "Observaciones"
                  Top             =   2565
                  Width           =   4155
               End
               Begin UltraGrid.SSUltraGrid grdImagenes 
                  Height          =   3945
                  Left            =   -74880
                  TabIndex        =   25
                  Top             =   450
                  Width           =   12540
                  _ExtentX        =   22119
                  _ExtentY        =   6959
                  _Version        =   131072
                  GridFlags       =   17040384
                  LayoutFlags     =   67108884
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "grdImagenes"
               End
               Begin UltraGrid.SSUltraGrid grdFarmacia 
                  Height          =   5175
                  Left            =   -74820
                  TabIndex        =   26
                  Top             =   465
                  Width           =   12465
                  _ExtentX        =   21987
                  _ExtentY        =   9128
                  _Version        =   131072
                  GridFlags       =   17040384
                  LayoutFlags     =   67108884
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "grdFarmacia"
               End
               Begin UltraGrid.SSUltraGrid grdLaboratorio 
                  Height          =   3945
                  Left            =   -74805
                  TabIndex        =   27
                  Top             =   480
                  Width           =   12480
                  _ExtentX        =   22013
                  _ExtentY        =   6959
                  _Version        =   131072
                  GridFlags       =   17040384
                  LayoutFlags     =   67108884
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "grdLaboratorio"
               End
               Begin UltraGrid.SSUltraGrid grdOtrosCpt 
                  Height          =   4125
                  Left            =   -74880
                  TabIndex        =   28
                  Top             =   450
                  Width           =   12540
                  _ExtentX        =   22119
                  _ExtentY        =   7276
                  _Version        =   131072
                  GridFlags       =   17040384
                  LayoutFlags     =   67108884
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Otros CPT"
               End
               Begin VB.Label Label3 
                  Caption         =   "* Pulsar ENTER para ver RESULTADO"
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
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   30
                  Top             =   4425
                  Width           =   8805
               End
               Begin VB.Label Label2 
                  Caption         =   "* Pulsar ENTER para ver RESULTADO"
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
                  Height          =   255
                  Left            =   -74805
                  TabIndex        =   29
                  Top             =   4425
                  Width           =   8805
               End
            End
            Begin UltraGrid.SSUltraGrid grdAnteriores 
               Height          =   1950
               Left            =   45
               TabIndex        =   31
               Top             =   450
               Width           =   12750
               _ExtentX        =   22490
               _ExtentY        =   3440
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Relación de Citas "
            End
            Begin UltraGrid.SSUltraGrid grdApoyoDx 
               Height          =   6885
               Left            =   -74895
               TabIndex        =   33
               Top             =   900
               Width           =   12675
               _ExtentX        =   22357
               _ExtentY        =   12144
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "grdApoyoDx"
            End
            Begin VB.Label Label4 
               Caption         =   "* Pulsar ENTER para ver RESULTADO"
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
               Height          =   255
               Left            =   -74955
               TabIndex        =   32
               Top             =   7950
               Width           =   8805
            End
         End
         Begin SISGalenPlus.ucPerinatalAS ucPerinatalAS1 
            Height          =   7035
            Left            =   -74955
            TabIndex        =   7
            Top             =   1065
            Width           =   11730
            _ExtentX        =   20690
            _ExtentY        =   12409
         End
         Begin UltraGrid.SSUltraGrid grdControles 
            Height          =   6915
            Left            =   -63180
            TabIndex        =   8
            Top             =   1140
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   12197
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "N°"
         End
         Begin SISGalenPlus.ucHCelectronicaLista ucHCelectronicaLista1 
            Height          =   8295
            Left            =   -74835
            TabIndex        =   10
            Top             =   360
            Width           =   12660
            _ExtentX        =   22331
            _ExtentY        =   14631
         End
         Begin VB.Label lblControl 
            AutoSize        =   -1  'True
            Caption         =   "......"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   -74850
            TabIndex        =   9
            Top             =   750
            Width           =   270
         End
      End
      Begin VB.TextBox txtPaciente 
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
         Left            =   990
         TabIndex        =   1
         Top             =   180
         Width           =   10035
      End
      Begin VB.Label Label1 
         Caption         =   "Paciente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Width           =   825
      End
   End
End
Attribute VB_Name = "AdmisionCEhistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Historicos de Atenciones de un paciente
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHDatos.CatalogoServicios
Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad

Dim ml_IdPruebaSeleccionada As String
Dim ml_NombrePruebaSeleccionada As String
Dim ml_nombrePaciente As String
Dim ml_idOrden As Long
Dim ml_IdProducto As Long

Dim ml_CodigoPruebaSeleccionada As String
Dim ml_NombreMedico As String
Dim ml_areaTrabajo As Long
Dim ml_idOrdenLab As Long
Dim ldFechaNacimiento As Date
Dim mo_Formulario As New sighEntidades.Formulario
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsCitasAnteriores As New Recordset
Dim oRsServiciosIntermedios As New Recordset
Dim oRsControlesNS As New Recordset
Dim oConexion As New Connection
Dim ml_idCuentaAtencion As Long
Dim lcSql As String
Dim lnIdProducto As Long
Const lcLineaChar As String = "¨"
Dim ml_NroHistoriaClinica As Long
Dim ml_Paciente As String
Dim ml_IdPaciente As Long
Dim ml_idTipoSexo As Long
Dim lbCargaUnaSolaVez  As Boolean
Dim lbMuestraTab As Integer     '1-Niño Sano     2-General
Dim ml_Medico As String
Dim oRsSoloResultadosImagenes As New Recordset
Dim oRsSoloResultadosLaboratorio As New Recordset
 
Property Let MuestraTab(lValue As Integer)
  lbMuestraTab = lValue
End Property

Property Let idPaciente(lValue As Long)
  ml_IdPaciente = lValue
End Property
Property Let Paciente(lValue As String)
  ml_Paciente = lValue
End Property
Property Let NroHistoriaClinica(lValue As Long)
  ml_NroHistoriaClinica = lValue
End Property
Property Let idTipoSexo(lValue As Long)
    ml_idTipoSexo = lValue
End Property
Property Let Medico(lValue As String)
  ml_Medico = lValue
End Property

Private Sub btnCancelar_Click()
    Unload Me
End Sub


Private Sub btnImprime_Click()
    On Error GoTo ErrImp
    Dim oRptHistoriaClinicaCE As New RptHistoriaClinicaCE
    Dim oRsTmp As New Recordset
    Dim lnIdAntencion As Long, lnIdCuenta As Long
    Dim oReglasAdmision As New ReglasAdmision
    Dim lcDxMedico As String, lcDx As String
    lcDx = txtDx.Text
    lcDxMedico = ""
    lnIdAntencion = 0
    lnIdCuenta = 0
    If Not IsNull(oRsCitasAnteriores.Fields!idAtencion) And oRsCitasAnteriores.Fields!idAtencion > 0 Then
        If Trim(txtDx.Text) = "" Then
            MsgBox "Pulse Doble Click en una cita de la relación de citas", vbInformation, "Informacion"
            Exit Sub
        End If
       lcDx = Left(txtDx.Text, InStr(txtDx.Text, Chr(13)) - 1)
       lcDxMedico = Mid(txtDx.Text, InStr(txtDx.Text, lcLineaChar) + 1, 1000)
       Set oRsTmp = oReglasAdmision.AtencionesSeleccionarPorIdAtencion(oRsCitasAnteriores.Fields!idAtencion)
       lnIdAntencion = 0
       If oRsTmp.RecordCount > 0 Then
          lnIdCuenta = oRsTmp.Fields!idCuentaAtencion
       End If
    End If
    oRptHistoriaClinicaCE.CrearReporteCeAtencionPaciente Me.hwnd, lnIdAntencion, _
                     Trim(txtPaciente.Text) & _
                     " (Edad: " & Trim(Str(oRsCitasAnteriores.Fields!TriajeEdad)) & ")", lnIdCuenta, _
                     oRsCitasAnteriores.Fields!CitaServicioJamo, oRsCitasAnteriores.Fields!CitaFecha, _
                     oRsCitasAnteriores.Fields!CitaMedico, IIf(IsNull(oRsCitasAnteriores.Fields!TriajePresion), "", _
                     oRsCitasAnteriores.Fields!TriajePresion), IIf(IsNull(oRsCitasAnteriores.Fields!triajeTalla), "", _
                     oRsCitasAnteriores.Fields!triajeTalla), IIf(IsNull(oRsCitasAnteriores.Fields!TriajeTemperatura), "", _
                     oRsCitasAnteriores.Fields!TriajeTemperatura), IIf(IsNull(oRsCitasAnteriores.Fields!triajePeso), "", _
                     oRsCitasAnteriores.Fields!triajePeso), txtCitaMotivo.Text, txtCitaExamenClinico.Text, lcDxMedico, _
                     lcDx, TxtCitaTratamiento.Text, txtCitaExClinicos.Text, txtCitaObservaciones.Text, False, 0, False
    Set oRsTmp = Nothing
    Set oReglasAdmision = Nothing
    Set oRptHistoriaClinicaCE = Nothing
ErrImp:

End Sub

Sub CargaSoloResultadosImagenesLaboratorio()
 On Error GoTo errCargResult
 Set grdApoyoDx.DataSource = mo_AdminAdmision.ServiciosIntermediosSeleccionarPorPaciente(ml_IdPaciente, False, True, 0, _
                                                                   True, IIf(chkSoloConResultados.Value = 1, True, False))
errCargResult:
End Sub

Private Sub chkSoloConResultados_Click()
    Me.MousePointer = 11
    CargaSoloResultadosImagenesLaboratorio
    mo_Apariencia.ConfigurarFilasBiColores grdApoyoDx, sighEntidades.GrillaConFilasBicolor
    Me.MousePointer = 1
End Sub

Private Sub Form_Load()
   On Error GoTo ErrLoad
   
   mo_Formulario.HabilitarDeshabilitar Me.lblDx, False
   mo_Formulario.HabilitarDeshabilitar Me.lblExamen, False
   mo_Formulario.HabilitarDeshabilitar Me.lblExamenes, False
   mo_Formulario.HabilitarDeshabilitar Me.lblMotivo, False
   mo_Formulario.HabilitarDeshabilitar Me.lblObservaciones, False
   mo_Formulario.HabilitarDeshabilitar Me.lblTratamiento, False
   

   Me.txtPaciente.Locked = True
   Me.txtCitaExamenClinico.Locked = True
   Me.txtCitaExClinicos.Locked = True
   Me.txtCitaMotivo.Locked = True
   Me.txtCitaObservaciones.Locked = True
   Me.TxtCitaTratamiento.Locked = True
   Me.txtDx.Locked = True
   Me.Caption = ml_Paciente
   txtPaciente.Text = ml_Paciente
   oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
   oConexion.CursorLocation = adUseClient
   oConexion.CommandTimeout = 150
   Set oRsCitasAnteriores = mo_AdminAdmision.AtencionesCeXnrohistoria(ml_NroHistoriaClinica, oConexion)
   Set Me.grdAnteriores.DataSource = oRsCitasAnteriores
   grdFarmacia.Caption = "": grdImagenes.Caption = "": grdLaboratorio.Caption = ""
   mo_Apariencia.ConfigurarFilasBiColores grdAnteriores, sighEntidades.GrillaConFilasBicolor
   mo_Apariencia.ConfigurarFilasBiColores grdFarmacia, sighEntidades.GrillaConFilasBicolor
   mo_Apariencia.ConfigurarFilasBiColores grdImagenes, sighEntidades.GrillaConFilasBicolor
   mo_Apariencia.ConfigurarFilasBiColores grdLaboratorio, sighEntidades.GrillaConFilasBicolor
   mo_Apariencia.ConfigurarFilasBiColores grdOtrosCpt, sighEntidades.GrillaConFilasBicolor
   '
   Set oRsControlesNS = mo_ReglasComunes.PerinatalAtencionCredSeleccionarControles(ml_IdPaciente)
   Set grdControles.DataSource = oRsControlesNS
   mo_Apariencia.ConfigurarFilasBiColores grdControles, sighEntidades.GrillaConFilasBicolor
   '
   TieneAccesoAlBotonImprimir
   '
   grdControles_DblClick   'CARGA niño sano
   '
   ucHCelectronicaLista1.Inicializar
   ucHCelectronicaLista1.DesdeHistorico ml_NroHistoriaClinica   'CARGA general
   '
   
   'CargaSoloResultadosImagenesLaboratorio
   '
   Select Case lbMuestraTab
   Case 1   'niño sano
     tabModulos.Tab = 1
   Case 2   'General
     tabModulos.Tab = 2
   End Select
   '
ErrLoad:
End Sub

Sub TieneAccesoAlBotonImprimir()
    Dim oRsPermisos As New Recordset
    Set oRsPermisos = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(sighEntidades.Usuario)
    oRsPermisos.Filter = "IdPermiso=407"
    If oRsPermisos.RecordCount = 0 Then
       btnImprime.Enabled = False
    End If
    Set oRsPermisos = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    oRsCitasAnteriores.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oRsCitasAnteriores = Nothing
End Sub

Private Sub grdAnteriores_DblClick()
    On Error GoTo ErrAnt
    lblMotivo.Text = "Cita: " & Format(oRsCitasAnteriores.Fields!CitaFecha, "dd/mm/yyyy") & "  (Motivo)"
    txtCitaMotivo.Text = IIf(IsNull(oRsCitasAnteriores.Fields!CitaMotivo), "", oRsCitasAnteriores.Fields!CitaMotivo)
    txtCitaExamenClinico.Text = IIf(IsNull(oRsCitasAnteriores.Fields!CitaExamenClinico), "", oRsCitasAnteriores.Fields!CitaExamenClinico)
    txtCitaExClinicos.Text = IIf(IsNull(oRsCitasAnteriores.Fields!CitaExClinicos), "", oRsCitasAnteriores.Fields!CitaExClinicos)
    TxtCitaTratamiento.Text = IIf(IsNull(oRsCitasAnteriores.Fields!CitaTratamiento), "", oRsCitasAnteriores.Fields!CitaTratamiento)
    txtCitaObservaciones.Text = IIf(IsNull(oRsCitasAnteriores.Fields!CitaObservaciones), "", oRsCitasAnteriores.Fields!CitaObservaciones)
    txtDx.Text = IIf(IsNull(oRsCitasAnteriores.Fields!CitaDiagMed), "", oRsCitasAnteriores.Fields!CitaDiagMed)
    '
    ml_idCuentaAtencion = 0
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorIdAtencion(oRsCitasAnteriores.Fields!idAtencion)
    If oRsTmp1.RecordCount > 0 Then
       ml_idCuentaAtencion = oRsTmp1.Fields!idCuentaAtencion
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    '
    TabHistoricos.Tab = 0
ErrAnt:
End Sub

Private Sub grdAnteriores_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        grdAnteriores.Bands(0).Columns("IdAtencion").Hidden = True
        grdAnteriores.Bands(0).Columns("NroHistoriaClinica").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaDniMedicoJamo").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaIdServicio").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaMotivo").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaExamenClinico").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaDiagMed").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaExClinicos").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaTratamiento").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaObservaciones").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaFechaAtencion").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaIdUsuario").Hidden = True
        grdAnteriores.Bands(0).Columns("TriajeIdUsuario").Hidden = True
        grdAnteriores.Bands(0).Columns("TriajeFecha").Hidden = True
        grdAnteriores.Bands(0).Columns("CitaFecha").Header.Caption = "Fecha"
        grdAnteriores.Bands(0).Columns("CitaFecha").Width = 1200
        grdAnteriores.Bands(0).Columns("CitaMedico").Header.Caption = "Médico"
        grdAnteriores.Bands(0).Columns("CitaMedico").Width = 2000
        grdAnteriores.Bands(0).Columns("CitaServicioJamo").Header.Caption = "Servicio"
        grdAnteriores.Bands(0).Columns("CitaServicioJamo").Width = 2500
        grdAnteriores.Bands(0).Columns("TriajeEdad").Header.Caption = "Edad"
        grdAnteriores.Bands(0).Columns("TriajeEdad").Width = 500
        grdAnteriores.Bands(0).Columns("TriajePresion").Header.Caption = "Presión"
        grdAnteriores.Bands(0).Columns("TriajePresion").Width = 1000
        grdAnteriores.Bands(0).Columns("TriajeTalla").Header.Caption = "Talla"
        grdAnteriores.Bands(0).Columns("TriajeTalla").Width = 1000
        grdAnteriores.Bands(0).Columns("TriajeTemperatura").Header.Caption = "Temperatura"
        grdAnteriores.Bands(0).Columns("TriajeTemperatura").Width = 1000
        grdAnteriores.Bands(0).Columns("TriajePeso").Header.Caption = "Peso"
        grdAnteriores.Bands(0).Columns("TriajePeso").Width = 1000
        
End Sub

Private Sub grdAnteriores_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdAnteriores_DblClick
    End If
End Sub

Private Sub grdApoyoDx_DblClick()
  On Error GoTo Fin
  
  
  Dim ml_IdPruebaSeleccionada As String
  Dim ml_NombrePruebaSeleccionada As String
  Dim ml_nombrePaciente As String
  Dim ml_idOrden As Long
  Dim ml_IdProducto As Long
  Dim ml_NombreMedico As String
  Dim ml_areaTrabajo As Long
  Dim ml_idOrdenLab As Long
  Dim oRsTmp As New Recordset
  Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
  

  
  
  'Cargar los formularios para el resultado
  Set oRsTmp = grdApoyoDx.DataSource
  If oRsTmp.Fields!resultado <> "SI" Then
     Set oRsTmp = Nothing
     Set mo_ReglasLaboratorio = Nothing
     Exit Sub
  End If
  If Len(oRsTmp!resultado1) > 0 And oRsTmp!resultado1 <> "SI" Then
     MsgBox oRsTmp!resultado1, vbInformation, "Resultado"
     Set oRsTmp = Nothing
     Set mo_ReglasLaboratorio = Nothing
     Exit Sub
  End If
  '*********************Imagen********************
  If Left(oRsTmp!ServicioApDx, 1) = "I" Then
    Dim mo_reglasImagen As New SIGHNegocios.ReglasImagenes
    Dim oResultadosImg As New SIGHImagen.ResultadosImg
    Dim rsResultados As New Recordset
    Dim oRsTmp9 As New Recordset
    Set oRsTmp9 = mo_ReglasComunes.ImagMovimientoSeleccionarIdOrden(oRsTmp!IdOrden)
    If oRsTmp9.RecordCount > 0 Then
        Set rsResultados = mo_reglasImagen.ImagMovimientoResultadosSeleccionarPorId(oRsTmp9!IdMovimiento)
        oResultadosImg.Producto = oRsTmp!Codigo & " " & oRsTmp!Item
        oResultadosImg.SoloEsConsulta = True
        oResultadosImg.idProductoCpt = oRsTmp!idProducto
        oResultadosImg.IdMovimiento = oRsTmp9!IdMovimiento
        Set oResultadosImg.rsResultados = rsResultados
        oResultadosImg.Paciente = txtPaciente.Text
        oResultadosImg.PuntoCarga = oRsTmp9!idPuntoCarga
        oResultadosImg.MostrarFormulario
    End If
    Set mo_ReglasLaboratorio = Nothing
    Set oRsTmp = Nothing
    Set mo_reglasImagen = Nothing
    Set oResultadosImg = Nothing
    Set rsResultados = Nothing
    Set oRsTmp9 = Nothing
    Exit Sub
  End If
  
  
  ml_IdPruebaSeleccionada = oRsTmp("codigo")
  ml_NombrePruebaSeleccionada = oRsTmp("item")
  ml_nombrePaciente = txtPaciente.Text
  ml_idOrden = oRsTmp("idOrden")
  ml_IdProducto = oRsTmp("idProducto")
  
  'debb-10/07/2018
  Dim ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
  ReglasArchivoClinico.ActualizaIDenTablaLabResultadoPorItems Nothing, ml_idOrden
  Set ReglasArchivoClinico = Nothing
  
  If mo_ReglasLaboratorio.UsaNuevaVentanaResultadosLaboratorio(oRsTmp!IdOrden, oRsTmp!Codigo) = True Then
      '************(inicio) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
    
      Dim oRsTmp1 As New Recordset
      Set oRsTmp1 = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.FactCatalogoServicios.Codigo='" & ml_IdPruebaSeleccionada & "'")
      If oRsTmp1.RecordCount > 0 Then
            Dim lcHistoria As String, lcEdadEnAtencion As String, lcServicioActualPaciente As String
            mo_ReglasLaboratorio.LlenaItemsConResultadosParaImpresion oRsTmp1, ml_idOrden, lcEdadEnAtencion, _
                                                                           lcHistoria, lcServicioActualPaciente, ml_IdPruebaSeleccionada
            mo_ReglasLaboratorio.Imprimir_LabResultadosItems oRsTmp1, lcEdadEnAtencion, lcHistoria, lcServicioActualPaciente
            oRsTmp1.Close
            Set oRsTmp1 = Nothing
            Exit Sub
      End If
      oRsTmp1.Close
      Set oRsTmp1 = Nothing
      '************(fin) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
  Else
    Dim oMuestraResultado As New SIGHLaboratorio.Ingresos
    oMuestraResultado.MuestraResultadoDelExamen ml_IdPruebaSeleccionada, ml_NombrePruebaSeleccionada, _
                                                ml_nombrePaciente, ml_idOrden, ml_IdPaciente, ml_NombreMedico, _
                                                ml_areaTrabajo, ml_idOrdenLab, ml_idTipoSexo, True
    Set oMuestraResultado = Nothing
  End If
  Set oRsTmp = Nothing
  Set mo_ReglasLaboratorio = Nothing
Fin:

End Sub

Private Sub grdApoyoDx_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        grdApoyoDx.Bands(0).Columns("idCuentaAtencion").Hidden = True
        grdApoyoDx.Bands(0).Columns("resultado1").Hidden = True
        grdApoyoDx.Bands(0).Columns("IdOrden").Hidden = True
        grdApoyoDx.Bands(0).Columns("idProducto").Hidden = True
        grdApoyoDx.Bands(0).Columns("Codigo").Hidden = True
        grdApoyoDx.Bands(0).Columns("Fecha").Width = 800
        grdApoyoDx.Bands(0).Columns("hora").Width = 500
        grdApoyoDx.Bands(0).Columns("servicioApDx").Width = 1000
        grdApoyoDx.Bands(0).Columns("item").Width = 3100
        grdApoyoDx.Bands(0).Columns("cantidad").Width = 400
        grdApoyoDx.Bands(0).Columns("resultado").Width = 1000
        grdApoyoDx.Bands(0).Columns("especialista").Width = 1500
        grdApoyoDx.Bands(0).Columns("resultado").CellMultiLine = ssCellMultiLineTrue
        grdApoyoDx.Bands(0).Columns("item").CellMultiLine = ssCellMultiLineTrue
        grdApoyoDx.Bands(0).Columns("Receta").Width = 800
        grdApoyoDx.Bands(0).Columns("UsuarioDespacho").Width = 1200
        
        grdApoyoDx.Caption = ""

End Sub

Private Sub grdApoyoDx_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
     If KeyAscii = 13 Then
        grdApoyoDx_DblClick
     End If
End Sub

Private Sub grdControles_DblClick()
       On Error GoTo eerr1
       Dim lnEdadEnDias As Integer, lnIdTipoEdad As Integer, lnPeso9 As Double, lnTalla9 As Long
       Dim oConexion1 As New Connection
       oConexion1.CommandTimeout = 900
       oConexion1.CursorLocation = adUseClient
       oConexion1.Open sighEntidades.CadenaConexion
       
       lnPeso9 = 0: lnTalla9 = 0
       If oRsCitasAnteriores.RecordCount > 0 Then
            oRsCitasAnteriores.MoveFirst
            oRsCitasAnteriores.Find "idatencion=" & oRsControlesNS!idAtencion
            If Not oRsCitasAnteriores.EOF Then
                 If Not IsNull(oRsCitasAnteriores!triajePeso) Then
                    lnPeso9 = Val(oRsCitasAnteriores!triajePeso)
                 End If
                 If Not IsNull(oRsCitasAnteriores!triajeTalla) Then
                    lnTalla9 = Val(oRsCitasAnteriores!triajeTalla)
                 End If
            End If
       End If
       
       lnEdadEnDias = oRsControlesNS!Edad
       lnIdTipoEdad = oRsControlesNS!idTipoEdad
       
           
        Me.ucPerinatalAS1.FechaAtencion = oRsControlesNS!fechaEgreso
        Me.ucPerinatalAS1.FechaNacimiento = oRsControlesNS!FechaNacimiento
       'If lbCargaUnaSolaVez = False Then
       '   lbCargaUnaSolaVez = True
           Me.ucPerinatalAS1.idUsuario = sighEntidades.Usuario
           Me.ucPerinatalAS1.Inicializar
       'End If
       Me.ucPerinatalAS1.idPaciente = ml_IdPaciente
       Me.ucPerinatalAS1.idAtencion = oRsControlesNS!idAtencion
       Me.ucPerinatalAS1.idTipoSexo = oRsControlesNS!idTipoSexo
       Me.ucPerinatalAS1.EdadEnMeses = sighEntidades.DevuelveEdadEnMeses(oRsControlesNS!FechaNacimiento, oRsControlesNS!fechaEgreso)
       Me.ucPerinatalAS1.CargaDatosAcontroles lnEdadEnDias, lnIdTipoEdad, lnPeso9, lnTalla9, oConexion1
       lblControl.Caption = "N° " & Trim(Str(oRsControlesNS!CredN)) & _
                            "  F.Atención: " & oRsControlesNS!fechaEgreso & _
                            "  N°Cuenta: " & Trim(Str(oRsControlesNS!idCuentaAtencion)) & _
                            "  Peso: " & Trim(Str(lnPeso9)) & _
                            "  Talla: " & Trim(Str(lnTalla9)) & _
                            "  " & IIf(lnPeso9 = 0 Or lnTalla9 = 0, "(sin GRAFICO)", "") & _
                            "  Edad: " & Trim(Str(lnEdadEnDias)) & " " & IIf(lnIdTipoEdad = 1, "Años", IIf(lnIdTipoEdad = 2, "Meses", IIf(lnIdTipoEdad = 3, "Días", "Horas"))) & _
                            "  Médico: " & oRsControlesNS!Medico
eerr1:
       oConexion1.Close
       Set oConexion1 = Nothing
End Sub


Private Sub grdControles_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        grdControles.Bands(0).Columns("credN").Header.Caption = "N°"
        grdControles.Bands(0).Columns("credN").Width = 500
End Sub

Private Sub grdFarmacia_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        grdFarmacia.Bands(0).Columns("Codigo").Width = 1200
        grdFarmacia.Bands(0).Columns("nombre").Width = 7200
        grdFarmacia.Bands(0).Columns("cantidad").Width = 1200

End Sub

Private Sub grdImagenes_DblClick()
    On Error GoTo ErrImg
    If oRsServiciosIntermedios.Fields!resultadoFinal <> "" Then
        MsgBox oRsServiciosIntermedios.Fields!resultadoFinal, vbInformation, "Resultado"
    ElseIf oRsServiciosIntermedios!resultado = "SI" Then
        '*********************Por ITEM********************
          Dim mo_reglasImagen As New SIGHNegocios.ReglasImagenes
          Dim oResultadosImg As New SIGHImagen.ResultadosImg
          Dim rsResultados As New Recordset
            Set rsResultados = mo_reglasImagen.ImagMovimientoResultadosSeleccionarPorId(oRsServiciosIntermedios!IdMovimiento)
            oResultadosImg.Producto = oRsServiciosIntermedios!Codigo & " " & oRsServiciosIntermedios!nombre
            oResultadosImg.SoloEsConsulta = True
            oResultadosImg.idProductoCpt = oRsServiciosIntermedios!idProducto
            oResultadosImg.IdMovimiento = oRsServiciosIntermedios!IdMovimiento
            Set oResultadosImg.rsResultados = rsResultados
            oResultadosImg.Paciente = ml_Paciente
            oResultadosImg.PuntoCarga = oRsServiciosIntermedios!idPuntoCarga
            oResultadosImg.MostrarFormulario
          Set mo_reglasImagen = Nothing
          Set oResultadosImg = Nothing
          Set rsResultados = Nothing
    End If
      
ErrImg:
End Sub

Private Sub grdImagenes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        grdImagenes.Bands(0).Columns("idOrden").Hidden = True
        grdImagenes.Bands(0).Columns("idProducto").Hidden = True
        grdImagenes.Bands(0).Columns("idPuntoCarga").Hidden = True
        grdImagenes.Bands(0).Columns("idMovimiento").Hidden = True
        grdImagenes.Bands(0).Columns("Codigo").Width = 1200
        grdImagenes.Bands(0).Columns("nombre").Width = 6500
        grdImagenes.Bands(0).Columns("cantidad").Width = 500
        grdImagenes.Bands(0).Columns("resultado").Width = 600
End Sub

Private Sub grdImagenes_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdImagenes_DblClick
    End If
End Sub

Private Sub grdImagenesResul_DblClick()
    On Error GoTo ErrImg
    If oRsSoloResultadosImagenes.Fields!resultadoFinal <> "" Then
        MsgBox oRsSoloResultadosImagenes.Fields!resultadoFinal, vbInformation, "Resultado"
    ElseIf oRsSoloResultadosImagenes!resultado = "SI" Then
        '*********************Por ITEM********************
          Dim mo_reglasImagen As New SIGHNegocios.ReglasImagenes
          Dim oResultadosImg As New SIGHImagen.ResultadosImg
          Dim rsResultados As New Recordset
            Set rsResultados = mo_reglasImagen.ImagMovimientoResultadosSeleccionarPorId(oRsSoloResultadosImagenes!IdMovimiento)
            oResultadosImg.Producto = oRsSoloResultadosImagenes!Codigo & " " & oRsSoloResultadosImagenes!nombre
            oResultadosImg.SoloEsConsulta = True
            oResultadosImg.idProductoCpt = oRsSoloResultadosImagenes!idProducto
            oResultadosImg.IdMovimiento = oRsSoloResultadosImagenes!IdMovimiento
            Set oResultadosImg.rsResultados = rsResultados
            oResultadosImg.Paciente = ml_Paciente
            oResultadosImg.PuntoCarga = oRsSoloResultadosImagenes!idPuntoCarga
            oResultadosImg.MostrarFormulario
          Set mo_reglasImagen = Nothing
          Set oResultadosImg = Nothing
          Set rsResultados = Nothing
    End If
      
ErrImg:

End Sub



Private Sub grdLaboratorio_DblClick()
  On Error GoTo Fin
  'Cargar los formularios para el resultado
  ml_IdPruebaSeleccionada = oRsServiciosIntermedios("codigo")
  ml_NombrePruebaSeleccionada = oRsServiciosIntermedios("nombre")
  ml_nombrePaciente = txtPaciente.Text
  ml_idOrden = oRsServiciosIntermedios("idOrden")
  ml_IdProducto = oRsServiciosIntermedios("idProducto")
  
  'debb-10/07/2018
  Dim ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
  ReglasArchivoClinico.ActualizaIDenTablaLabResultadoPorItems Nothing, ml_idOrden
  Set ReglasArchivoClinico = Nothing
  If mo_ReglasLaboratorio.UsaNuevaVentanaResultadosLaboratorio(oRsServiciosIntermedios!IdOrden, oRsServiciosIntermedios!Codigo) = True Then
      '************(inicio) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
      Dim oRsTmp1 As New Recordset
      Set oRsTmp1 = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.FactCatalogoServicios.Codigo='" & ml_IdPruebaSeleccionada & "'")
      'If oRsTmp1.RecordCount = 0 Then
            Dim lcHistoria As String, lcEdadEnAtencion As String, lcServicioActualPaciente As String
            mo_ReglasLaboratorio.LlenaItemsConResultadosParaImpresion oRsTmp1, ml_idOrden, lcEdadEnAtencion, _
                                                                           lcHistoria, lcServicioActualPaciente, ml_IdPruebaSeleccionada
            mo_ReglasLaboratorio.Imprimir_LabResultadosItems oRsTmp1, lcEdadEnAtencion, lcHistoria, lcServicioActualPaciente
            oRsTmp1.Close
            Set oRsTmp1 = Nothing
            Exit Sub
      'End If
      oRsTmp1.Close
      Set oRsTmp1 = Nothing
      
      '************(fin) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
  Else
    Dim oMuestraResultado As New SIGHLaboratorio.Ingresos
    oMuestraResultado.MuestraResultadoDelExamen ml_IdPruebaSeleccionada, ml_NombrePruebaSeleccionada, _
                                                ml_nombrePaciente, ml_idOrden, ml_IdPaciente, ml_NombreMedico, _
                                                ml_areaTrabajo, ml_idOrdenLab, ml_idTipoSexo, True
    Set oMuestraResultado = Nothing
  End If
  Exit Sub
  
Fin:
  Exit Sub
End Sub

Private Sub grdLaboratorio_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        grdLaboratorio.Bands(0).Columns("idOrden").Hidden = True
        grdLaboratorio.Bands(0).Columns("Codigo").Width = 1200
        grdLaboratorio.Bands(0).Columns("nombre").Width = 6500
        grdLaboratorio.Bands(0).Columns("cantidad").Width = 1200
        grdLaboratorio.Bands(0).Columns("resultado").Width = 1200
End Sub

Private Sub grdLaboratorio_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdLaboratorio_DblClick
    End If
End Sub

Private Sub grdLaboratorioSoloResultado_DblClick()
  On Error GoTo Fin
  'Cargar los formularios para el resultado
  ml_IdPruebaSeleccionada = oRsSoloResultadosLaboratorio("codigo")
  ml_NombrePruebaSeleccionada = oRsSoloResultadosLaboratorio("nombre")
  ml_nombrePaciente = txtPaciente.Text
  ml_idOrden = oRsSoloResultadosLaboratorio("idOrden")
  ml_IdProducto = oRsSoloResultadosLaboratorio("idProducto")
  
  'debb-10/07/2018
  Dim ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
  ReglasArchivoClinico.ActualizaIDenTablaLabResultadoPorItems Nothing, ml_idOrden
  Set ReglasArchivoClinico = Nothing
  If mo_ReglasLaboratorio.UsaNuevaVentanaResultadosLaboratorio(oRsSoloResultadosLaboratorio!IdOrden, oRsSoloResultadosLaboratorio!Codigo) = True Then
      '************(inicio) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
      Dim oRsTmp1 As New Recordset
      Set oRsTmp1 = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.FactCatalogoServicios.Codigo='" & ml_IdPruebaSeleccionada & "'")
      If oRsTmp1.RecordCount > 0 Then
            Dim lcHistoria As String, lcEdadEnAtencion As String, lcServicioActualPaciente As String
            mo_ReglasLaboratorio.LlenaItemsConResultadosParaImpresion oRsTmp1, ml_idOrden, lcEdadEnAtencion, _
                                                                           lcHistoria, lcServicioActualPaciente, ml_IdPruebaSeleccionada
            mo_ReglasLaboratorio.Imprimir_LabResultadosItems oRsTmp1, lcEdadEnAtencion, lcHistoria, lcServicioActualPaciente
            oRsTmp1.Close
            Set oRsTmp1 = Nothing
            Exit Sub
      End If
      oRsTmp1.Close
      Set oRsTmp1 = Nothing
      
      '************(fin) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
  Else
    Dim oMuestraResultado As New SIGHLaboratorio.Ingresos
    oMuestraResultado.MuestraResultadoDelExamen ml_IdPruebaSeleccionada, ml_NombrePruebaSeleccionada, _
                                                ml_nombrePaciente, ml_idOrden, ml_IdPaciente, ml_NombreMedico, _
                                                ml_areaTrabajo, ml_idOrdenLab, ml_idTipoSexo, True
    Set oMuestraResultado = Nothing
  End If
  Exit Sub
  
Fin:
  Exit Sub

End Sub



Private Sub grdOtrosCpt_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        grdOtrosCpt.Bands(0).Columns("IdPuntoCarga").Hidden = True
        grdOtrosCpt.Bands(0).Columns("IdCuentaAtencion").Hidden = True
        grdOtrosCpt.Bands(0).Columns("codigo").Width = 1000
        grdOtrosCpt.Bands(0).Columns("nombre").Width = 6300
        grdOtrosCpt.Bands(0).Columns("cantidad").Width = 600
        grdOtrosCpt.Bands(0).Columns("precio").Width = 600
        grdOtrosCpt.Bands(0).Columns("total").Width = 1000

End Sub

Private Sub SSTabResultados_Click(PreviousTab As Integer)
        
        Select Case SSTabResultados.Tab
        Case 0   '
        Case 1   'Imagenes/laboratorio
             Me.MousePointer = 11
             CargaSoloResultadosImagenesLaboratorio
             mo_Apariencia.ConfigurarFilasBiColores grdApoyoDx, sighEntidades.GrillaConFilasBicolor
             Me.MousePointer = 1
        End Select
End Sub



Private Sub TabHistoricos_Click(PreviousTab As Integer)
   On Error GoTo errTab
   If oRsCitasAnteriores.Fields!idAtencion > 0 Then
        Dim oRsTmp1 As New Recordset
        Dim mo_reglasImagen As New SIGHNegocios.ReglasImagenes
        Dim lcResultado As String
        Select Case TabHistoricos.Tab
        Case 0   'Citas
        Case 1   'Imagenes
             If oRsServiciosIntermedios.State = adStateOpen Then
                 Set oRsServiciosIntermedios = Nothing
             End If
             With oRsServiciosIntermedios
                .Fields.Append "idOrden", adDouble
                .Fields.Append "Codigo", adVarChar, 20
                .Fields.Append "Nombre", adVarChar, 250
                .Fields.Append "Cantidad", adInteger
                .Fields.Append "ResultadoFinal", adVarChar, 3000
                .Fields.Append "Resultado", adVarChar, 2
                .Fields.Append "idProducto", adInteger
                .Fields.Append "idMovimiento", adInteger
                .Fields.Append "idPuntoCarga", adInteger
                .LockType = adLockOptimistic
                .Open
             End With
             Set oRsTmp1 = mo_ReglasComunes.ConsumoCptEnImagenesPorIdAntencionParaParticulares(oRsCitasAnteriores.Fields!idAtencion)
             If oRsTmp1.RecordCount = 0 Then
                oRsTmp1.Close
                Set oRsTmp1 = mo_ReglasComunes.ConsumoCptEnImagenesPorIdAntencionConSeguro(oRsCitasAnteriores.Fields!idAtencion)
             End If
             If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                   lcResultado = IIf(mo_reglasImagen.TieneResultadoImagenesPorCpt(oRsTmp1!IdMovimiento, oRsTmp1!idProductoCpt) = True, "SI", "NO")
                   oRsServiciosIntermedios.AddNew
                   oRsServiciosIntermedios.Fields!IdOrden = oRsTmp1!IdOrden
                   oRsServiciosIntermedios.Fields!Codigo = oRsTmp1.Fields!Codigo
                   oRsServiciosIntermedios.Fields!nombre = oRsTmp1.Fields!nombre
                   oRsServiciosIntermedios.Fields!Cantidad = oRsTmp1.Fields!Cantidad
                   oRsServiciosIntermedios.Fields!idProducto = oRsTmp1!idProductoCpt
                   oRsServiciosIntermedios.Fields!resultadoFinal = IIf(IsNull(oRsTmp1!resultadoFinal), "", oRsTmp1!resultadoFinal)
                   oRsServiciosIntermedios!IdMovimiento = oRsTmp1!IdMovimiento
                   oRsServiciosIntermedios!idPuntoCarga = oRsTmp1!idPuntoCarga
                   oRsServiciosIntermedios!resultado = lcResultado
                   oRsServiciosIntermedios.Update
                   oRsTmp1.MoveNext
                Loop
             End If
             Set grdImagenes.DataSource = oRsServiciosIntermedios
        Case 2   'farmacia
             If oRsServiciosIntermedios.State = 1 Then
                oRsServiciosIntermedios.Close
             End If
             Set oRsServiciosIntermedios = mo_ReglasComunes.ConsumoMedicamentosPorIdCuentaConSeguroYparticular(ml_idCuentaAtencion)
             Set grdFarmacia.DataSource = oRsServiciosIntermedios
        Case 3   'Laboratorio
             If oRsServiciosIntermedios.State = adStateOpen Then
                 Set oRsServiciosIntermedios = Nothing
             End If
             With oRsServiciosIntermedios
                .Fields.Append "idOrden", adDouble
                .Fields.Append "Codigo", adVarChar, 20
                .Fields.Append "Nombre", adVarChar, 250
                .Fields.Append "Cantidad", adInteger
                .Fields.Append "Resultado", adVarChar, 2
                .Fields.Append "idProducto", adInteger
                .LockType = adLockOptimistic
                .Open
             End With
             Set oRsTmp1 = mo_ReglasComunes.ConsumoCptEnLaboratorioPorIdAntencionConSeguro(oRsCitasAnteriores.Fields!idAtencion)
             If oRsTmp1.RecordCount = 0 Then
                oRsTmp1.Close
                Set oRsTmp1 = mo_ReglasComunes.ConsumoCptEnLaboratorioPorIdAntencionParaParticulares(oRsCitasAnteriores.Fields!idAtencion)
             End If
             If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                   mo_ReglasLaboratorio.ResultadosAutomaticosActualizaHaciaGalenhos oRsTmp1!IdOrden
                   
                   oRsServiciosIntermedios.AddNew
                   oRsServiciosIntermedios.Fields!IdOrden = oRsTmp1.Fields!IdOrden
                   oRsServiciosIntermedios.Fields!Codigo = oRsTmp1.Fields!Codigo
                   oRsServiciosIntermedios.Fields!nombre = oRsTmp1.Fields!nombre
                   oRsServiciosIntermedios.Fields!Cantidad = oRsTmp1.Fields!Cantidad
                   oRsServiciosIntermedios.Fields!idProducto = oRsTmp1.Fields!idProductoCpt
                   If mo_ReglasLaboratorio.PruebaTieneResultado(oRsTmp1!Codigo, oRsTmp1!IdOrden) = True Then
                      oRsServiciosIntermedios.Fields!resultado = "SI"
                   Else
                      oRsServiciosIntermedios.Fields!resultado = "NO"
                   End If
                   oRsServiciosIntermedios.Update
                   oRsTmp1.MoveNext
                Loop
             End If
             Set grdLaboratorio.DataSource = oRsServiciosIntermedios
        Case 4  'Otros CPT
             Set oRsServiciosIntermedios = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(ml_idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
             Set grdOtrosCpt.DataSource = oRsServiciosIntermedios
        End Select
        Set oRsTmp1 = Nothing
        Set mo_reglasImagen = Nothing
   End If
errTab:
End Sub






