VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form AtencionesCenso 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Censo de Estancia Hospitalaria"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12735
   Icon            =   "AtencionesCenso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabCenso 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "HOSPITALIZADOS"
      TabPicture(0)   =   "AtencionesCenso.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatosHistoria"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "CONFIGURACIÓN"
      TabPicture(1)   =   "AtencionesCenso.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDesde(1)"
      Tab(1).Control(1)=   "Frame(6)"
      Tab(1).Control(2)=   "Frame(4)"
      Tab(1).Control(3)=   "Frame(2)"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(5)=   "Label1(0)"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtDesde 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   -69360
         MaxLength       =   4
         TabIndex        =   55
         Top             =   480
         Width           =   1395
      End
      Begin VB.Frame Frame 
         Caption         =   "3er Rango (En Porcentaje)"
         Height          =   4215
         Index           =   6
         Left            =   -66480
         TabIndex        =   39
         Top             =   960
         Width           =   4095
         Begin VB.TextBox txtDesde 
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
            Index           =   3
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   50
            Top             =   480
            Width           =   1365
         End
         Begin VB.Frame Frame 
            Caption         =   "Composición de color RGB"
            Height          =   3015
            Index           =   7
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   3855
            Begin VB.Frame frameColor 
               Height          =   1095
               Index           =   3
               Left            =   120
               TabIndex        =   54
               Top             =   1800
               Width           =   3615
            End
            Begin VB.HScrollBar scrLimRojo 
               Height          =   375
               Index           =   3
               Left            =   720
               TabIndex        =   46
               Top             =   360
               Width           =   2415
            End
            Begin VB.HScrollBar scrLimVerde 
               Height          =   375
               Index           =   3
               Left            =   720
               TabIndex        =   45
               Top             =   840
               Width           =   2415
            End
            Begin VB.HScrollBar scrLimAzul 
               Height          =   375
               Index           =   3
               Left            =   720
               TabIndex        =   44
               Top             =   1320
               Width           =   2415
            End
            Begin VB.TextBox txtRojo 
               Height          =   375
               Index           =   3
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   43
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtVerde 
               Height          =   375
               Index           =   3
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   42
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txtAzul 
               Height          =   375
               Index           =   3
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   41
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Rojo"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   49
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Verde"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   48
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Azul"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   47
               Top             =   1440
               Width           =   855
            End
         End
         Begin VB.Label lblDesde 
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
            Index           =   3
            Left            =   2880
            TabIndex        =   60
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label 
            Caption         =   "Mayor a"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "2do Rango (En Porcentaje)"
         Height          =   4215
         Index           =   4
         Left            =   -70680
         TabIndex        =   24
         Top             =   960
         Width           =   4095
         Begin VB.Frame Frame 
            Caption         =   "Composición de color RGB"
            Height          =   3015
            Index           =   5
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   3855
            Begin VB.Frame frameColor 
               Height          =   1095
               Index           =   2
               Left            =   120
               TabIndex        =   53
               Top             =   1800
               Width           =   3615
            End
            Begin VB.TextBox txtAzul 
               Height          =   375
               Index           =   2
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   33
               Top             =   1320
               Width           =   495
            End
            Begin VB.TextBox txtVerde 
               Height          =   375
               Index           =   2
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   32
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txtRojo 
               Height          =   375
               Index           =   2
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   31
               Top             =   360
               Width           =   495
            End
            Begin VB.HScrollBar scrLimAzul 
               Height          =   375
               Index           =   2
               Left            =   720
               TabIndex        =   30
               Top             =   1320
               Width           =   2415
            End
            Begin VB.HScrollBar scrLimVerde 
               Height          =   375
               Index           =   2
               Left            =   720
               TabIndex        =   29
               Top             =   840
               Width           =   2415
            End
            Begin VB.HScrollBar scrLimRojo 
               Height          =   375
               Index           =   2
               Left            =   720
               TabIndex        =   28
               Top             =   360
               Width           =   2415
            End
            Begin VB.Label Label4 
               Caption         =   "Azul"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   36
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Verde"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   35
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Rojo"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   34
               Top             =   480
               Width           =   855
            End
         End
         Begin VB.TextBox txtHasta 
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
            Index           =   2
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   26
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txtDesde 
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
            Index           =   2
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   25
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label lblDesde 
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
            Index           =   2
            Left            =   2880
            TabIndex        =   59
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblHasta 
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
            Index           =   2
            Left            =   2880
            TabIndex        =   58
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Menor o Igual a"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label 
            Caption         =   "Mayor a"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "1er Rango (En Porcentaje)"
         Height          =   4215
         Index           =   2
         Left            =   -74880
         TabIndex        =   11
         Top             =   960
         Width           =   4095
         Begin VB.TextBox txtHasta 
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
            Index           =   1
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   22
            Top             =   480
            Width           =   1365
         End
         Begin VB.Frame AtencionesCenso 
            Caption         =   "Composición de color RGB"
            Height          =   3015
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   3855
            Begin VB.Frame frameColor 
               Height          =   1095
               Index           =   1
               Left            =   120
               TabIndex        =   52
               Top             =   1800
               Width           =   3615
            End
            Begin VB.HScrollBar scrLimRojo 
               Height          =   375
               Index           =   1
               Left            =   720
               TabIndex        =   18
               Top             =   360
               Width           =   2415
            End
            Begin VB.HScrollBar scrLimVerde 
               Height          =   375
               Index           =   1
               Left            =   720
               TabIndex        =   17
               Top             =   840
               Width           =   2415
            End
            Begin VB.HScrollBar scrLimAzul 
               Height          =   375
               Index           =   1
               Left            =   720
               TabIndex        =   16
               Top             =   1320
               Width           =   2415
            End
            Begin VB.TextBox txtRojo 
               Height          =   375
               Index           =   1
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   15
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtVerde 
               Height          =   375
               Index           =   1
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   14
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txtAzul 
               Height          =   375
               Index           =   1
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   13
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Rojo"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   21
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Verde"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Azul"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   1440
               Width           =   855
            End
         End
         Begin VB.Label lblHasta 
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
            Index           =   1
            Left            =   2880
            TabIndex        =   57
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Menor o Igual a"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   8
         Top             =   5160
         Width           =   12555
         Begin VB.CommandButton btnSalir 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "AtencionesCenso.frx":0D02
            DownPicture     =   "AtencionesCenso.frx":11C6
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
            Left            =   6420
            Picture         =   "AtencionesCenso.frx":16B2
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton btnAceptarConfig 
            Caption         =   "Aceptar (F2)"
            DisabledPicture =   "AtencionesCenso.frx":1B9E
            DownPicture     =   "AtencionesCenso.frx":1FFE
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
            Left            =   4890
            Picture         =   "AtencionesCenso.frx":2473
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   5160
         Width           =   12555
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Exportar (F2)"
            DisabledPicture =   "AtencionesCenso.frx":28E8
            DownPicture     =   "AtencionesCenso.frx":2D48
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
            Left            =   4890
            Picture         =   "AtencionesCenso.frx":31BD
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton btnCancelar 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "AtencionesCenso.frx":3632
            DownPicture     =   "AtencionesCenso.frx":3AF6
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
            Left            =   6420
            Picture         =   "AtencionesCenso.frx":3FE2
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   225
            Width           =   1365
         End
      End
      Begin VB.Frame fraDatosHistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   12555
         Begin VB.TextBox txtHospitalizados 
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
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "HOSPITALIZADOS"
            Top             =   240
            Width           =   3105
         End
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   3240
            Picture         =   "AtencionesCenso.frx":44CE
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1305
         End
         Begin UltraGrid.SSUltraGrid grdCenso 
            Height          =   4095
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   7223
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
            Caption         =   "Censo diario"
         End
      End
      Begin VB.Label Label1 
         Caption         =   "UIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -69840
         TabIndex        =   56
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "AtencionesCenso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de atenciones Censo de hospitalizados
'        Programado por: Cachay F
'        Fecha: Febrero 2015
'
'------------------------------------------------------------------------------------
Option Explicit

Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim mo_Formulario As New sighentidades.Formulario
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsHospitalizados As New Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Const ml_Rango1 = 1, ml_Rango2 = 2, ml_Rango3 = 3
Dim ml_Rojo1, ml_Verde1, ml_Azul1 As Integer
Dim ml_Rojo2, ml_Verde2, ml_Azul2 As Integer
Dim ml_Rojo3, ml_Verde3, ml_Azul3 As Integer
Dim oDoAtencionHospCenso1 As New DoAtencionHospCenso
Dim oDoAtencionHospCenso2 As New DoAtencionHospCenso
Dim oDoAtencionHospCenso3 As New DoAtencionHospCenso
Dim ms_MensajeError As String

Private Sub btnAceptar_Click()
    Dim oRsTemp As New Recordset
    Set oRsTemp = grdCenso.DataSource
    If oRsTemp.RecordCount > 0 Then
        Dim oclAtencionesCenso As New clAtencionesCenso
        Me.MousePointer = 11
        oclAtencionesCenso.CrearReporte oRsTemp, Me.hwnd
        Me.MousePointer = 1
    Else
        MsgBox "No existe información para exportar", vbInformation, Me.Caption
        Exit Sub
    End If
End Sub

Function ValidarDatosObligatorios() As Boolean
    Dim sMensaje As String
    sMensaje = ""
    ValidarDatosObligatorios = False
    
    If Me.txtHasta(ml_Rango1).Text = "" Then
        sMensaje = sMensaje + "Ingrese el valor 'Menor o Igual a' en el 1er rango" + Chr(13)
    End If
    
    If Me.txtDesde(ml_Rango1).Text = "" Then
        sMensaje = sMensaje + "Ingrese el valor de la UIT" + Chr(13)
    End If
    
    If Me.txtDesde(ml_Rango2).Text = "" Then
        sMensaje = sMensaje + "Ingrese el valor 'Mayor a' en el 2do rango" + Chr(13)
    End If
    
    If Me.txtHasta(ml_Rango2).Text = "" Then
        sMensaje = sMensaje + "Ingrese el valor 'Menor o Igual a' en el 2do rango" + Chr(13)
    End If
    
    If Me.txtDesde(ml_Rango3).Text = "" Then
        sMensaje = sMensaje + "Ingrese el valor 'Mayor a' en el 3er rango" + Chr(13)
    End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function

Function ValidarReglas() As Boolean
    ValidarReglas = False
    
    If Val(Me.txtHasta(2).Text) < Val(Me.txtDesde(2).Text) Then
        MsgBox "El 2do rango es incorrecto, el valor superior no puede ser menor que el valor inferior", vbExclamation, Me.Caption
        Exit Function
    End If
    
    ValidarReglas = True
End Function

Private Sub btnAceptarConfig_Click()
    If ValidarDatosObligatorios() Then
        If ValidarReglas() Then
            If ModificarDatos() Then
                MsgBox "Los datos se guardarón correctamente", vbInformation, Me.Caption
'                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo modificar los datos de la configuración" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
    End If
End Sub

Function ModificarDatos() As Boolean
    ModificarDatos = False
    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_reglasComunes.AtenHospCensoSeleccionarModificar(oDoAtencionHospCenso1, _
                                                                        oDoAtencionHospCenso2, _
                                                                        oDoAtencionHospCenso3)
    ms_MensajeError = mo_reglasComunes.MensajeError
End Function

Sub CargaDatosAlObjetosDeDatos()
    With oDoAtencionHospCenso1
        .IdRangoCensoHosp = ml_Rango1
        .IdUsuarioAuditoria = 1
        .RangoFinal = Me.txtHasta(ml_Rango1).Text
        .RangoInicial = Me.txtDesde(ml_Rango1).Text
        .RGBAZUL = Me.txtAzul(ml_Rango1).Text
        .RGBROJO = Me.txtRojo(ml_Rango1).Text
        .RGBVERDE = Me.txtVerde(ml_Rango1).Text
    End With
    
    With oDoAtencionHospCenso2
        .IdRangoCensoHosp = ml_Rango2
        .IdUsuarioAuditoria = 1
        .RangoFinal = Me.txtHasta(ml_Rango2).Text
        .RangoInicial = Me.txtDesde(ml_Rango2).Text
        .RGBAZUL = Me.txtAzul(ml_Rango2).Text
        .RGBROJO = Me.txtRojo(ml_Rango2).Text
        .RGBVERDE = Me.txtVerde(ml_Rango2).Text
    End With
    
    With oDoAtencionHospCenso3
        .IdRangoCensoHosp = ml_Rango3
        .IdUsuarioAuditoria = 1
        .RangoInicial = Me.txtDesde(ml_Rango3).Text
        .RGBAZUL = Me.txtAzul(ml_Rango3).Text
        .RGBROJO = Me.txtRojo(ml_Rango3).Text
        .RGBVERDE = Me.txtVerde(ml_Rango3).Text
    End With
End Sub

Sub LimpiarHospitalizados()
    If oRsHospitalizados.RecordCount > 0 Then
        oRsHospitalizados.MoveFirst
        Do While Not oRsHospitalizados.EOF
            oRsHospitalizados.Delete
            oRsHospitalizados.Update
            oRsHospitalizados.MoveNext
        Loop
    End If
End Sub

Private Sub btnBuscar_Click()
    Dim oRpt As New Recordset
    Dim oclAtencionesCenso As New clAtencionesCenso
    Dim oDoAtencionHospCenso1 As New DoAtencionHospCenso
    Dim oDoAtencionHospCenso2 As New DoAtencionHospCenso
    Dim oDoAtencionHospCenso3 As New DoAtencionHospCenso
    Set oRpt = oclAtencionesCenso.AtencionesCensoEstanciaHospitalariaPacientes
    
    LimpiarHospitalizados
    
    Set oDoAtencionHospCenso1 = mo_reglasComunes.AtenHospCensoSeleccionarPorId(ml_Rango1)
    Set oDoAtencionHospCenso2 = mo_reglasComunes.AtenHospCensoSeleccionarPorId(ml_Rango2)
    Set oDoAtencionHospCenso3 = mo_reglasComunes.AtenHospCensoSeleccionarPorId(ml_Rango3)
                
    If oRpt.RecordCount > 0 Then
        oRpt.MoveFirst
        Do While Not oRpt.EOF
        oRsHospitalizados.AddNew
        oRsHospitalizados.Fields!NroDocumento = oRpt.Fields!NroDocumento
        oRsHospitalizados.Fields!Apellido_Paterno = oRpt.Fields!Apellido_Paterno
        oRsHospitalizados.Fields!Apellido_Materno = oRpt.Fields!Apellido_Materno
        oRsHospitalizados.Fields!Nombres = oRpt.Fields!Nombres
        oRsHospitalizados.Fields!Fec_Nac = oRpt.Fields!Fec_Nac
        oRsHospitalizados.Fields!NroCama = oRpt.Fields!NroCama
        oRsHospitalizados.Fields!Servicio_Ingreso_Origen = oRpt.Fields!Servicio_Ingreso_Origen
        oRsHospitalizados.Fields!FechaHora_IngresoEstabl = oRpt.Fields!FechaHora_IngresoEstabl
        oRsHospitalizados.Fields!Servicio_Actual_Atencion = oRpt.Fields!Servicio_Actual_Atencion
        oRsHospitalizados.Fields!IdCuenta_Atencion = oRpt.Fields!IdCuenta_Atencion
        oRsHospitalizados.Fields!TotalPorPagar = IIf(IsNull(oRpt.Fields!TotalPorPagar), 0, oRpt.Fields!TotalPorPagar)
        
        If Val(oRsHospitalizados.Fields!TotalPorPagar) <= 2 * oDoAtencionHospCenso1.RangoFinal * (oDoAtencionHospCenso1.RangoInicial / 100) Then
            oRsHospitalizados.Fields!RGBROJO = oDoAtencionHospCenso1.RGBROJO
            oRsHospitalizados.Fields!RGBVERDE = oDoAtencionHospCenso1.RGBVERDE
            oRsHospitalizados.Fields!RGBAZUL = oDoAtencionHospCenso1.RGBAZUL
        Else
            If Val(oRsHospitalizados.Fields!TotalPorPagar) > 2 * oDoAtencionHospCenso2.RangoInicial * (oDoAtencionHospCenso1.RangoInicial / 100) And _
                    Val(oRsHospitalizados.Fields!TotalPorPagar) <= 2 * oDoAtencionHospCenso2.RangoFinal * (oDoAtencionHospCenso1.RangoInicial / 100) Then
                oRsHospitalizados.Fields!RGBROJO = oDoAtencionHospCenso2.RGBROJO
                oRsHospitalizados.Fields!RGBVERDE = oDoAtencionHospCenso2.RGBVERDE
                oRsHospitalizados.Fields!RGBAZUL = oDoAtencionHospCenso2.RGBAZUL
            Else
                If Val(oRsHospitalizados.Fields!TotalPorPagar) > 2 * oDoAtencionHospCenso3.RangoInicial * (oDoAtencionHospCenso1.RangoInicial / 100) Then
                    oRsHospitalizados.Fields!RGBROJO = oDoAtencionHospCenso3.RGBROJO
                    oRsHospitalizados.Fields!RGBVERDE = oDoAtencionHospCenso3.RGBVERDE
                    oRsHospitalizados.Fields!RGBAZUL = oDoAtencionHospCenso3.RGBAZUL
                End If
            End If
        End If
        oRpt.MoveNext
        Loop
    End If
    Set grdCenso.DataSource = oRsHospitalizados
    mo_Apariencia.ConfigurarFilasBiColores grdCenso, sighentidades.GrillaConFilasBicolor
    
'    grdCenso.Bands(0).Col .Override.RowAppearance.BackColor = &HFDF0E6
    
'    grdEmpleados.Bands(1).Override.RowAlternateAppearance.BackColor = &HFFFFFF
'    grdEmpleados.Bands(1).Override.RowAppearance.BackColor = &HDAFDFE
End Sub

Sub CreaTemporaloHospitalizados()
    If oRsHospitalizados.State = 1 Then
       Set oRsHospitalizados = Nothing
    End If
    With oRsHospitalizados
        .Fields.Append "NroDocumento", adVarChar, 255, adFldIsNullable
        .Fields.Append "Apellido_Paterno", adVarChar, 255, adFldIsNullable
        .Fields.Append "Apellido_Materno", adVarChar, 255, adFldIsNullable
        .Fields.Append "Nombres", adVarChar, 255, adFldIsNullable
        .Fields.Append "Fec_Nac", adVarChar, 255, adFldIsNullable
        .Fields.Append "NroCama", adVarChar, 255, adFldIsNullable
        .Fields.Append "Servicio_Ingreso_Origen", adVarChar, 255, adFldIsNullable
        .Fields.Append "FechaHora_IngresoEstabl", adVarChar, 255, adFldIsNullable
        .Fields.Append "Servicio_Actual_Atencion", adVarChar, 255, adFldIsNullable
        .Fields.Append "IdCuenta_Atencion", adInteger
        .Fields.Append "TotalPorPagar", adInteger, adFldIsNullable
        .Fields.Append "RGBRojo", adInteger
        .Fields.Append "RGBVerde", adInteger
        .Fields.Append "RGBAzul", adInteger
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub btnSalir_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Private Sub Form_Load()
    InicializarColores
    CreaTemporaloHospitalizados
    CargarDatosAlosControles
End Sub

Private Sub grdCenso_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCenso.Bands(0).Columns("NroDocumento").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("Apellido_Paterno").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("Apellido_Materno").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("Nombres").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("Fec_Nac").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("NroCama").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("Servicio_Ingreso_Origen").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("FechaHora_IngresoEstabl").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("Servicio_Actual_Atencion").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("IdCuenta_Atencion").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("TotalPorPagar").Activation = ssActivationActivateNoEdit
    grdCenso.Bands(0).Columns("RGBRojo").Hidden = True
    grdCenso.Bands(0).Columns("RGBVerde").Hidden = True
    grdCenso.Bands(0).Columns("RGBAzul").Hidden = True
End Sub

Private Sub grdCenso_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
'        If IsNull(Row.Cells("TotalPorPagar").GetText()) Then
'            Row.Appearance.BackColor = frameColor(ml_Rango1).BackColor
'        Else
'            If Row.Cells("TotalPorPagar").GetText() = "" Then
'                Row.Appearance.BackColor = frameColor(ml_Rango1).BackColor
'            Else
'                If Val(Row.Cells("TotalPorPagar").GetText()) <= Val(Me.lblHasta(ml_Rango1).Caption) Then
'                    Row.Appearance.BackColor = frameColor(ml_Rango1).BackColor
'                Else
'                    If Val(Row.Cells("TotalPorPagar").GetText()) > Val(Me.lblDesde(ml_Rango2).Caption) And _
'                       Val(Row.Cells("TotalPorPagar").GetText()) <= Val(Me.lblHasta(ml_Rango2).Caption) Then
'                       Row.Appearance.BackColor = frameColor(ml_Rango2).BackColor
'                    Else
'                        If Val(Row.Cells("TotalPorPagar").GetText()) > Val(Me.lblDesde(ml_Rango3).Caption) Then
'                            Row.Appearance.BackColor = frameColor(ml_Rango3).BackColor
'                        End If
'                    End If
'                End If
'            End If
'        End If
    Row.Appearance.BackColor = RGB(Val(Row.Cells("RGBRojo").GetText()), Val(Row.Cells("RGBVerde").GetText()), Val(Row.Cells("RGBAzul").GetText()))
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
End Sub

Sub CambiarColorPorScroll(ByVal Index As Integer)
    Me.txtRojo(Index).Text = scrLimRojo(Index).Value
    Me.txtVerde(Index).Text = scrLimVerde(Index).Value
    Me.txtAzul(Index).Text = scrLimAzul(Index).Value
    frameColor(Index).BackColor = RGB(Val(Me.txtRojo(Index).Text), Val(Me.txtVerde(Index).Text), Val(Me.txtAzul(Index).Text))
End Sub

Private Sub scrLimAzul_Change(Index As Integer)
    CambiarColorPorScroll (Index)
End Sub

Private Sub scrLimAzul_Scroll(Index As Integer)
    CambiarColorPorScroll (Index)
End Sub

Private Sub scrLimRojo_Change(Index As Integer)
    CambiarColorPorScroll (Index)
End Sub

Private Sub scrLimRojo_Scroll(Index As Integer)
    CambiarColorPorScroll (Index)
End Sub

Private Sub scrLimVerde_Change(Index As Integer)
    CambiarColorPorScroll (Index)
End Sub

Private Sub scrLimVerde_Scroll(Index As Integer)
    CambiarColorPorScroll (Index)
End Sub

Private Sub txtDesde_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    Else
        If Not (mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtDesde_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     If Index <> ml_Rango1 Then
        If Val(Me.txtDesde(Index).Text) > 100 Then Me.txtDesde(Index).Text = 100
    End If
    If Index = ml_Rango2 Then
        Me.txtHasta(ml_Rango1).Text = Me.txtDesde(ml_Rango2).Text
        Me.lblHasta(ml_Rango1).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtHasta(ml_Rango1).Text) / 100)
        Me.lblDesde(ml_Rango2).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtDesde(ml_Rango2).Text) / 100)
    End If
    If Index = 3 Then
        Me.txtHasta(ml_Rango2).Text = Me.txtDesde(ml_Rango3).Text
        Me.lblHasta(ml_Rango2).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtHasta(ml_Rango2).Text) / 100)
        Me.lblDesde(ml_Rango3).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtDesde(ml_Rango3).Text) / 100)
    End If
End Sub

Private Sub txtHasta_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtHasta_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Val(Me.txtHasta(Index).Text) > 100 Then Me.txtHasta(Index).Text = 100
    If Index = 1 Then
        Me.txtDesde(ml_Rango2).Text = Me.txtHasta(ml_Rango1).Text
        Me.lblDesde(ml_Rango2).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtDesde(ml_Rango2).Text) / 100)
        Me.lblHasta(ml_Rango1).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtHasta(ml_Rango1).Text) / 100)
    End If
    If Index = 2 Then
        Me.txtDesde(ml_Rango3).Text = Me.txtHasta(ml_Rango2).Text
        Me.lblDesde(ml_Rango3).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtDesde(ml_Rango3).Text) / 100)
        Me.lblHasta(ml_Rango2).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtHasta(ml_Rango2).Text) / 100)
    End If
End Sub

Private Sub txtHospitalizados_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    End If
End Sub

Sub CargarDatosAlosControles()
    Dim oDoAtencionHospCenso As New DoAtencionHospCenso
    
    'Carga el primer rango
    Set oDoAtencionHospCenso = mo_reglasComunes.AtenHospCensoSeleccionarPorId(ml_Rango1)
    Me.txtHasta(ml_Rango1).Text = oDoAtencionHospCenso.RangoFinal
    Me.txtRojo(ml_Rango1).Text = oDoAtencionHospCenso.RGBROJO: scrLimRojo(ml_Rango1).Value = oDoAtencionHospCenso.RGBROJO
    Me.txtVerde(ml_Rango1).Text = oDoAtencionHospCenso.RGBVERDE: scrLimVerde(ml_Rango1).Value = oDoAtencionHospCenso.RGBVERDE
    Me.txtAzul(ml_Rango1).Text = oDoAtencionHospCenso.RGBAZUL: scrLimAzul(ml_Rango1).Value = oDoAtencionHospCenso.RGBAZUL
    
    Me.txtDesde(ml_Rango1).Text = CInt(oDoAtencionHospCenso.RangoInicial)
    
    'Carga el segundo rango
    Set oDoAtencionHospCenso = mo_reglasComunes.AtenHospCensoSeleccionarPorId(ml_Rango2)
    Me.txtDesde(ml_Rango2).Text = oDoAtencionHospCenso.RangoInicial
    Me.txtHasta(ml_Rango2).Text = oDoAtencionHospCenso.RangoFinal
    Me.txtRojo(ml_Rango2).Text = oDoAtencionHospCenso.RGBROJO: scrLimRojo(ml_Rango2).Value = oDoAtencionHospCenso.RGBROJO
    Me.txtVerde(ml_Rango2).Text = oDoAtencionHospCenso.RGBVERDE: scrLimVerde(ml_Rango2).Value = oDoAtencionHospCenso.RGBVERDE
    Me.txtAzul(ml_Rango2).Text = oDoAtencionHospCenso.RGBAZUL: scrLimAzul(ml_Rango2).Value = oDoAtencionHospCenso.RGBAZUL
    
    'Carga el 3er rango
    Set oDoAtencionHospCenso = mo_reglasComunes.AtenHospCensoSeleccionarPorId(ml_Rango3)
    Me.txtDesde(ml_Rango3).Text = oDoAtencionHospCenso.RangoInicial
    Me.txtRojo(ml_Rango3).Text = oDoAtencionHospCenso.RGBROJO: scrLimRojo(ml_Rango3).Value = oDoAtencionHospCenso.RGBROJO
    Me.txtVerde(ml_Rango3).Text = oDoAtencionHospCenso.RGBVERDE: scrLimVerde(ml_Rango3).Value = oDoAtencionHospCenso.RGBVERDE
    Me.txtAzul(ml_Rango3).Text = oDoAtencionHospCenso.RGBAZUL: scrLimAzul(ml_Rango3).Value = oDoAtencionHospCenso.RGBAZUL
    
    frameColor(ml_Rango1).BackColor = RGB(Val(Me.txtRojo(ml_Rango1).Text), Val(Me.txtVerde(ml_Rango1).Text), Val(Me.txtAzul(ml_Rango1).Text))
    frameColor(ml_Rango2).BackColor = RGB(Val(Me.txtRojo(ml_Rango2).Text), Val(Me.txtVerde(ml_Rango2).Text), Val(Me.txtAzul(ml_Rango2).Text))
    frameColor(ml_Rango3).BackColor = RGB(Val(Me.txtRojo(ml_Rango3).Text), Val(Me.txtVerde(ml_Rango3).Text), Val(Me.txtAzul(ml_Rango3).Text))
    
    Me.lblHasta(ml_Rango1).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtHasta(ml_Rango1).Text) / 100)
    Me.lblDesde(ml_Rango2).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtDesde(ml_Rango2).Text) / 100)
    Me.lblHasta(ml_Rango2).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtHasta(ml_Rango2).Text) / 100)
    Me.lblDesde(ml_Rango3).Caption = 2 * Val(Me.txtDesde(ml_Rango1).Text) * (Val(Me.txtDesde(ml_Rango3).Text) / 100)

    Set oDoAtencionHospCenso = Nothing
End Sub

Sub InicializarColores()
    ml_Rojo1 = 0: ml_Verde1 = 0: ml_Azul1 = 0
    ml_Rojo2 = 0: ml_Verde2 = 0: ml_Azul2 = 0
    ml_Rojo3 = 0: ml_Verde3 = 0: ml_Azul3 = 0
    
    frameColor(1).BackColor = RGB(ml_Rojo1, ml_Verde1, ml_Azul1)
    frameColor(2).BackColor = RGB(ml_Rojo2, ml_Verde2, ml_Azul2)
    frameColor(3).BackColor = RGB(ml_Rojo3, ml_Verde3, ml_Azul3)
    
    scrLimRojo(ml_Rango1).Min = 0: scrLimRojo(ml_Rango1).Max = 255: scrLimRojo(ml_Rango1).SmallChange = 1
    scrLimVerde(ml_Rango1).Min = 0: scrLimVerde(ml_Rango1).Max = 255: scrLimVerde(ml_Rango1).SmallChange = 1
    scrLimAzul(ml_Rango1).Min = 0: scrLimAzul(ml_Rango1).Max = 255: scrLimAzul(ml_Rango1).SmallChange = 1
    
    scrLimRojo(ml_Rango2).Min = 0: scrLimRojo(ml_Rango2).Max = 255: scrLimRojo(ml_Rango2).SmallChange = 1
    scrLimVerde(ml_Rango2).Min = 0: scrLimVerde(ml_Rango2).Max = 255: scrLimVerde(ml_Rango2).SmallChange = 1
    scrLimAzul(ml_Rango2).Min = 0: scrLimAzul(ml_Rango2).Max = 255: scrLimAzul(ml_Rango2).SmallChange = 1
    
    scrLimRojo(ml_Rango3).Min = 0: scrLimRojo(ml_Rango3).Max = 255: scrLimRojo(ml_Rango3).SmallChange = 1
    scrLimVerde(ml_Rango3).Min = 0: scrLimVerde(ml_Rango3).Max = 255: scrLimVerde(ml_Rango3).SmallChange = 1
    scrLimAzul(ml_Rango3).Min = 0: scrLimAzul(ml_Rango3).Max = 255: scrLimAzul(ml_Rango3).SmallChange = 1
    
End Sub

Private Sub txtRojo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtVerde_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAzul_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRojo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    CambiarColorTexto (Index)
End Sub

Private Sub txtVerde_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    CambiarColorTexto (Index)
End Sub

Private Sub txtAzul_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    CambiarColorTexto (Index)
End Sub

Sub CambiarColorTexto(Index)
    If Me.txtRojo(Index).Text = "" Then Me.txtRojo(Index).Text = 0
    If Me.txtVerde(Index).Text = "" Then Me.txtVerde(Index).Text = 0
    If Me.txtAzul(Index).Text = "" Then Me.txtAzul(Index).Text = 0

    If Val(Me.txtRojo(Index).Text) > 255 Then Me.txtRojo(Index).Text = 255
    If Val(Me.txtVerde(Index).Text) > 255 Then Me.txtVerde(Index).Text = 255
    If Val(Me.txtAzul(Index).Text) > 255 Then Me.txtAzul(Index).Text = 255
    scrLimRojo(Index).Value = Me.txtRojo(Index).Text
    scrLimVerde(Index).Value = Me.txtVerde(Index).Text
    scrLimAzul(Index).Value = Me.txtAzul(Index).Text
    frameColor(Index).BackColor = RGB(Val(Me.txtRojo(Index).Text), Val(Me.txtVerde(Index).Text), Val(Me.txtAzul(Index).Text))
End Sub
