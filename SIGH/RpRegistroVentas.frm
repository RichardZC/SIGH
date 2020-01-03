VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form RpRegistroVentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Ventas"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "RpRegistroVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   7
      Top             =   6360
      Width           =   5565
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "RpRegistroVentas.frx":0CCA
         DownPicture     =   "RpRegistroVentas.frx":112A
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
         Left            =   1350
         Picture         =   "RpRegistroVentas.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "RpRegistroVentas.frx":1A14
         DownPicture     =   "RpRegistroVentas.frx":1ED8
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
         Left            =   2940
         Picture         =   "RpRegistroVentas.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6345
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5565
      Begin SISGalenPlus.UcTipoImpresion UcTipoImpresion1 
         Height          =   555
         Left            =   4110
         TabIndex        =   39
         Top             =   5745
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   979
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
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
         Left            =   90
         TabIndex        =   38
         Top             =   5520
         Width           =   1110
      End
      Begin VB.CheckBox chkSoloCredito 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo CREDITOS"
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
         Left            =   1800
         TabIndex        =   37
         Top             =   5475
         Width           =   1575
      End
      Begin SISGalenPlus.XP_ProgressBar progressRpt 
         Height          =   225
         Left            =   105
         TabIndex        =   36
         Top             =   5940
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   397
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
         Color           =   12937777
      End
      Begin VB.CheckBox chkSoloAnuladas 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo anuladas"
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
         Left            =   3870
         Picture         =   "RpRegistroVentas.frx":28B0
         TabIndex        =   35
         Top             =   5475
         Width           =   1605
      End
      Begin VB.Frame FraFiltros 
         Height          =   1935
         Left            =   90
         TabIndex        =   26
         Top             =   945
         Width           =   5400
         Begin VB.ComboBox cmbIdCaja 
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
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   210
            Width           =   3705
         End
         Begin VB.ComboBox cmbIdTurno 
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
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   630
            Width           =   3705
         End
         Begin VB.ComboBox cmbIdResponsable 
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
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1050
            Width           =   3705
         End
         Begin VB.ComboBox cmbIdTipoComprobante 
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
            Left            =   1575
            TabIndex        =   27
            Top             =   1455
            Width           =   3720
         End
         Begin VB.Label Label1 
            Caption         =   "Turno"
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
            Left            =   90
            TabIndex        =   34
            Top             =   675
            Width           =   1365
         End
         Begin VB.Label Label4 
            Caption         =   "Caja"
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
            Left            =   90
            TabIndex        =   33
            Top             =   255
            Width           =   1365
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cajero"
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
            TabIndex        =   32
            Top             =   1095
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante"
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
            TabIndex        =   31
            Top             =   1515
            Width           =   1110
         End
      End
      Begin VB.Frame FraXitem 
         Height          =   510
         Left            =   5130
         TabIndex        =   18
         Top             =   5715
         Visible         =   0   'False
         Width           =   315
         Begin VB.CommandButton btnProdConsulta 
            Caption         =   "..."
            Height          =   315
            Left            =   1545
            TabIndex        =   22
            ToolTipText     =   "Busca CPT"
            Top             =   225
            Width           =   315
         End
         Begin VB.TextBox txtProductoConsulta 
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
            MaxLength       =   50
            TabIndex        =   20
            Top             =   225
            Width           =   3315
         End
         Begin VB.TextBox txtCodProductoConsulta 
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
            Left            =   480
            TabIndex        =   19
            Top             =   225
            Width           =   1000
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "CPT"
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
            TabIndex        =   21
            Top             =   285
            Width           =   330
         End
      End
      Begin VB.Frame Frame0 
         Height          =   645
         Left            =   90
         TabIndex        =   17
         Top             =   2970
         Width           =   5415
         Begin VB.OptionButton optReporteX 
            Caption         =   "Por CPT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   2
            Left            =   4230
            TabIndex        =   25
            Top             =   150
            Width           =   975
         End
         Begin VB.OptionButton optReporteX 
            Caption         =   "Detall x Items"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   150
            TabIndex        =   24
            Top             =   150
            Width           =   1440
         End
         Begin VB.OptionButton optReporteX 
            Caption         =   "Detall x Documento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   1
            Left            =   1897
            TabIndex        =   23
            Top             =   150
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.Frame fraFarmacia 
         Height          =   1215
         Left            =   90
         TabIndex        =   12
         Top             =   4245
         Visible         =   0   'False
         Width           =   5415
         Begin VB.ComboBox cmbAlmOrigen 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            TabIndex        =   15
            Top             =   630
            Width           =   3720
         End
         Begin VB.ComboBox cmdVendorFarmacia 
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
            Left            =   1560
            TabIndex        =   13
            Text            =   "cmdVendorFarmacia"
            Top             =   240
            Width           =   3705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Farmacia"
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
            TabIndex        =   16
            Top             =   690
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor Farm"
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
            TabIndex        =   14
            Top             =   300
            Width           =   1260
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo"
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
         Left            =   90
         TabIndex        =   8
         Top             =   3615
         Width           =   5415
         Begin Threed.SSOption optFarmacia 
            Height          =   255
            Left            =   150
            TabIndex        =   9
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
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
            Caption         =   "Farmacia"
         End
         Begin Threed.SSOption optServicio 
            Height          =   255
            Left            =   1897
            TabIndex        =   10
            Top             =   240
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "Servicios"
         End
         Begin Threed.SSOption optFarmaciaServicios 
            Height          =   255
            Left            =   4230
            TabIndex        =   11
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
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
            Caption         =   "Ambos"
            Value           =   -1
         End
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1605
         TabIndex        =   0
         Top             =   165
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   1605
         TabIndex        =   1
         Top             =   567
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Emisión Final"
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
         Left            =   120
         TabIndex        =   6
         Top             =   615
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Emisión Inicial"
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
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   1245
      End
   End
End
Attribute VB_Name = "RpRegistroVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Registro de Ventas
'        Programado por: Barrantes D
'        Fecha: Enero 2011
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_cmbIdCaja As New ListaDespleglable
Dim mo_cmbIdTurno As New ListaDespleglable
Dim mo_cmbIdResponsable As New sighEntidades.ListaDespleglable
Dim mo_cmdVendorFarmacia As New sighEntidades.ListaDespleglable
Dim mo_cmbAlmacenOrigen As New sighEntidades.ListaDespleglable
Dim ml_IdTipoReporte As Long, ml_idUsuario As Long
Dim ml_IdCaja As Long, ml_IdTurno As Long, ml_idTComprobante As Long
Dim ms_IdCaja As String, ms_IdTurno As String, ms_idTComprobante As String, ms_IdTipoReporte As String
Dim rsRsTmp As New ADODB.Recordset
Dim oRsCajeros As New Recordset
Dim mo_cmbIdTipoComprobante As New ListaDespleglable


Property Let idUsuario(lIdValue As Long)
  ml_idUsuario = lIdValue
End Property

Property Let IdTipoReporte(lIdValue As Long)
  ml_IdTipoReporte = lIdValue
End Property

Private Sub btnAceptar_Click()
  If Trim(txtFechaInicio.Text) = "" Or Not IsDate(txtFechaInicio.Text) Then
    MsgBox "Por favor ingrese la fecha inicial", vbInformation, Me.Caption
    Exit Sub
  End If
  If Trim(txtFechaFin.Text) = "" Or Not IsDate(txtFechaFin.Text) Then
    MsgBox "Por favor ingrese la fecha final", vbInformation, Me.Caption
    Exit Sub
  End If
  If cmbIdCaja.Text <> "" Then
    If cmbIdTurno.Text = "" Then
      MsgBox "Por favor elija el Turno", vbInformation, Me.Caption
      Exit Sub
    End If
  End If
  If cmbIdTipoComprobante.Text = "" Then
     cmbIdTipoComprobante_Click
  End If
  '+++++++++++++++
  Me.MousePointer = 11
  'CrearReporteDeRegistroDeVentas txtFechaInicio.Text, txtFechaFin.Text, Val(mo_cmbIdCaja.BoundText), Val(mo_cmbIdTurno.BoundText), ml_idTComprobante, ml_IdTipoReporte
'  CrearReporteDeRegistroDeVentasDebb txtFechaInicio.Text, txtFechaFin.Text, Val(mo_cmbIdCaja.BoundText), _
'                                     Val(mo_cmbIdTurno.BoundText), ml_idTComprobante, ml_IdTipoReporte, _
'                                     cmbIdResponsable.Text, Val(mo_cmdVendorFarmacia.BoundText), _
'                                     IIf(Me.chkExcel.Value = 1, True, False), Val(mo_cmbAlmacenOrigen.BoundText), _
'                                     IIf(Val(mo_cmbAlmacenOrigen.BoundText) = 0, "", "  (Farm:" & Trim(cmbAlmOrigen.Text) & ")")
  
  
   'modificado por Samuel
  If optReporteX(1).Value = True Then
      CrearReporteDeRegistroDeVentasDebb txtFechaInicio.Text, txtFechaFin.Text, Val(mo_cmbIdCaja.BoundText), _
                                     Val(mo_cmbIdTurno.BoundText), ml_idTComprobante, ml_IdTipoReporte, _
                                     cmbIdResponsable.Text, Val(mo_cmdVendorFarmacia.BoundText), _
                                     IIf(Me.chkExcel.Value = 1, True, False), Val(mo_cmbAlmacenOrigen.BoundText), _
                                     IIf(Val(mo_cmbAlmacenOrigen.BoundText) = 0, "", "  (Farm:" & Trim(cmbAlmOrigen.Text) & ")"), _
                                     IIf(Me.chkSoloAnuladas.Value = 1, True, False)
  ElseIf optReporteX(0).Value = True Then
        CrearReporteDeCierreCaja txtFechaInicio.Text, txtFechaFin.Text, Val(mo_cmbIdCaja.BoundText), _
                                Val(mo_cmbIdTurno.BoundText), ml_IdTipoReporte, Val(mo_cmbIdResponsable.BoundText), _
                                IIf(UcTipoImpresion1.OpcionImpresionElejida = sghTIexcel, True, False), _
                                Val(mo_cmbAlmacenOrigen.BoundText), Val(mo_cmdVendorFarmacia.BoundText), ml_idTComprobante
  Else
      If txtProductoConsulta.Text = "" Then
            MsgBox "Por favor elija PROCEDIMIENTO CPT", vbInformation, Me.Caption
            Exit Sub
      End If
      CrearReporteXitem txtFechaInicio.Text, txtFechaFin.Text, _
                        IIf(Me.chkExcel.Value = 1, True, False), _
                        txtCodProductoConsulta.Text, _
                        Me.txtProductoConsulta.Text, Val(Me.txtCodProductoConsulta.Tag)
  End If
  Me.MousePointer = 1
  'Me.progressRpt.Value = 0
  
  
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub



Private Sub cmbIdCaja_Click()
  ml_IdCaja = mo_cmbIdCaja.BoundText
  ms_IdCaja = " Caja: " & cmbIdCaja.Text & ". "
End Sub



Private Sub cmbIdTipoComprobante_Click()
     If Val(mo_cmbIdTipoComprobante.BoundText) > 0 Then
        ml_idTComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
        ms_idTComprobante = " Comprobante de Pago: " & cmbIdTipoComprobante.Text
     Else
        ml_idTComprobante = 0
        ms_idTComprobante = " Comprobante de Pago: <<TODOS>> "
     End If
End Sub

Private Sub cmbIdTipoComprobante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmbIdTipoComprobante_Click
    End If
End Sub

Private Sub cmbIdTurno_Click()
  ml_IdTurno = mo_cmbIdTurno.BoundText
  ms_IdTurno = " Turno: " & cmbIdTurno.Text & ". "
End Sub





Private Sub Form_Initialize()
  Set mo_cmbIdCaja.MiComboBox = cmbIdCaja
  Set mo_cmbIdTurno.MiComboBox = cmbIdTurno
  Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
  Set mo_cmdVendorFarmacia.MiComboBox = cmdVendorFarmacia
  Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
  Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
End Sub

Private Sub Form_Load()
  Dim oRsPermisos As New Recordset
  mo_Formulario.HabilitarDeshabilitar txtCodProductoConsulta, False
  mo_Formulario.HabilitarDeshabilitar txtProductoConsulta, False
  '
  ml_IdCaja = 0: ml_IdTurno = 0: ml_idTComprobante = 0: ml_IdTipoReporte = 0
  ms_IdCaja = "": ms_IdTurno = "": ms_idTComprobante = " Boletas de Venta y facturas": ms_IdTipoReporte = " Farmacia y Servicios "
  Me.txtFechaInicio.Text = Format(Date, "dd/mm/yyyy") & " 00:00:00"
  Me.txtFechaFin.Text = Format(Date, "dd/mm/yyyy") & " 23:59:59"
  Me.Caption = "Registro de Ventas"
  
  mo_cmbIdTurno.BoundColumn = "IdTurno"
  mo_cmbIdTurno.ListField = "Descripcion"
  Set mo_cmbIdTurno.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
        
  mo_cmbIdCaja.BoundColumn = "IdCaja"
  mo_cmbIdCaja.ListField = "Descripcion"
  Set mo_cmbIdCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()
  
  Set oRsCajeros = mo_AdminCaja.CajerosSeleccionarTodos()
  mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
  mo_cmbIdResponsable.ListField = "DCajero"
  Set mo_cmbIdResponsable.RowSource = oRsCajeros
  If oRsCajeros.RecordCount > 0 Then
    Set oRsPermisos = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(Val(sighEntidades.Usuario))
    oRsPermisos.Filter = "idPermiso=1000"
    If oRsPermisos.RecordCount > 0 Then
       mo_cmbIdResponsable.BoundText = sighEntidades.Usuario
       mo_Formulario.HabilitarDeshabilitar cmbIdResponsable, False
    End If
    oRsPermisos.Close
  End If
  
  mo_cmdVendorFarmacia.BoundColumn = "IdEmpleado"
  mo_cmdVendorFarmacia.ListField = "DEmpleado"
  Set mo_cmdVendorFarmacia.RowSource = mo_ReglasComunes.EmpleadosSeleccionarTodos
  
  mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
  mo_cmbAlmacenOrigen.ListField = "Descripcion"
  Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1")
  
  mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
  mo_cmbIdTipoComprobante.ListField = "Descripcion"
  Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()
  cmbIdTipoComprobante_Click

  
  ml_idTComprobante = 0
  ml_IdTipoReporte = 0
  
  Set oRsPermisos = Nothing
End Sub







Private Sub optFarmacia_Click(Value As Integer)
    If optFarmacia.Value = True Then
        ml_IdTipoReporte = 1
        ms_IdTipoReporte = " Tipo: Farmacia. "
        fraFarmacia.Visible = True
    End If
End Sub

Private Sub optFarmaciaServicios_Click(Value As Integer)
    If optFarmaciaServicios.Value = True Then
        ml_IdTipoReporte = 0
        ms_IdTipoReporte = " Tipo: Farmacia y Servicios. "
        fraFarmacia.Visible = False
        mo_cmdVendorFarmacia.BoundText = ""
    End If
End Sub



Private Sub optServicio_Click(Value As Integer)
    If optServicio.Value = True Then
        ml_IdTipoReporte = 2
        ms_IdTipoReporte = " Tipo: Servicios. "
        fraFarmacia.Visible = False
        mo_cmdVendorFarmacia.BoundText = ""
    End If
End Sub







Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyEscape
      btnCancelar_Click
    Case vbKeyF2
      btnAceptar_Click
  End Select
End Sub

Private Sub txtFechaFin_GotFocus()
  SeleccionaMask txtFechaFin
End Sub

Private Sub txtFechaFin_LostFocus()
If Not IsDate(txtFechaFin.Text) Then

        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaFin.Text = sighEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
    
  
End Sub

Private Sub txtFechaInicio_GotFocus()
  SeleccionaMask txtFechaInicio
End Sub

Sub CrearReporteDeRegistroDeVentas(FI As String, FF As String, IdCaja As Long, IdTurno As Long, idTComprobante As Long, idTReporte As Long)
  Dim oExcel As Excel.Application
  Dim oWorkBookPlantilla As Workbook
  Dim oWorkBook As Workbook
  Dim oWorkSheet As Worksheet
  Dim rsreporte As New Recordset
  Dim lcLlave As String
  Dim lnImpSubTot As Double, lntImpSubTot As Double
  Dim lnImpAnul As Double, lntImpAnul As Double
  Dim lnImpExo As Double, lntImpExo As Double
  Dim lnImpDevol As Double, lntImpDevol As Double
  Dim lnImpPagCta As Double, lntImpPagCta As Double
  Dim lnImpTot As Double, lntImpTot As Double: Dim lnDctos As Double
  Dim lnSubTotal As Double, lnIGV As Double, lnTSubTotal As Double, lntIgv As Double
  Dim iFila As Long, lnImpRedondeo As Double, lntImpRedondeo As Double
  Dim lRecordCount As Long, lcCadenaConexion As String
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  Dim lnIdPagosACuenta As Long
  'On Error GoTo ManejadorError
  Dim rsRsTmp As New Recordset 'Dim rsRsTmp As New ADODB.Recordset
  Dim mo_ReporteUtil As New ReporteUtil
  Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
  Dim oConexion As New ADODB.Connection
    
  Const COL_FECHA = 2
  Const COL_USUARIO = 3
  Const COL_BOLETA = 4
  Const COL_NRO_HISTORIA = 5
  Const COL_RAZON_SOCIAL = 6
  Const COL_SUBTOTAL = 7
  Const COL_REDONDEO = 8
  Const COL_ANULADO = 9
  Const COL_EXONERADO = 10
  Const COL_DEVOLUCION = 11
  Const COL_PAGOCTA = 12
  
  Const COL_TOTAL_BRUTO = 13
  Const COL_IGV = 14
  Const COL_TOTAL_NETO = 15
    
  lnIdPagosACuenta = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
  
  oConexion.Open sighEntidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
    
  With rsRsTmp
    .Fields.Append "FechaCobranza", adDouble
    .Fields.Append "NroSerie", adVarChar, 5, adFldIsNullable
    .Fields.Append "NroDocumento", adVarChar, 20, adFldIsNullable
    .Fields.Append "NroHistoriaClinica", adVarChar, 50, adFldIsNullable
    .Fields.Append "RazonSocial", adVarChar, 100, adFldIsNullable
    .Fields.Append "adelantos", adDouble
    .Fields.Append "ImporteEXO", adDouble
    .Fields.Append "idEstadoComprobante", adUnsignedBigInt
    .Fields.Append "totalPorPagar", adDouble
    .Fields.Append "IdProducto", adUnsignedBigInt
    .Fields.Append "subtotal", adDouble
    .Fields.Append "igv", adDouble
    .Fields.Append "totalPagado", adDouble
    .Fields.Append "redondeo", adDouble
    .Fields.Append "cajero", adVarChar, 20, adFldIsNullable
    .Fields.Append "Boleta", adVarChar, 30, adFldIsNullable
    .Fields.Append "anulado", adDouble
    .Fields.Append "Devolucion", adDouble
    .LockType = adLockOptimistic
    .Open
  End With
  
  'Crea nueva hoja
  Set oExcel = GalenhosExcelApplication()  'New Excel.Application
  Set oWorkBook = oExcel.Workbooks.Add
  'Abre, copia y cierra la plantilla
  Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECajaRegistroDeVentas.xls")
  oWorkBookPlantilla.Worksheets("RegistroDeVentas").Copy Before:=oWorkBook.Sheets(1)
  oWorkBookPlantilla.Close
  'Activa la primera hoja
  Set oWorkSheet = oWorkBook.Sheets(1)
  'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
  oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
  oWorkSheet.Cells(2, 3).Value = "REGISTRO DE VENTAS"
  oWorkSheet.Cells(4, 3).Value = "Del: " & FI & " al " & FF & ". " & ms_IdCaja & ms_IdTurno & ms_idTComprobante & ms_IdTipoReporte
  
  'SERVICIOS emitidos en CAJA SERVICIO
  If idTReporte = 0 Or idTReporte = 2 Then
    Set rsreporte = mo_ReglasFacturacion.SeleccionaServicios(FI, FF, IdCaja, IdTurno, idTComprobante, oConexion)
    If rsreporte.RecordCount > 0 Then
      With rsRsTmp
        rsreporte.MoveFirst
        Do While Not rsreporte.EOF
          If Val(rsreporte.Fields!nrodocumento) = 472294 Then lRecordCount = 0
          .AddNew
          .Fields!nroSerie = rsreporte.Fields!nroSerie
          .Fields!nrodocumento = rsreporte.Fields!nrodocumento
          .Fields!boleta = Trim(rsreporte.Fields!nroSerie) & " - " & Trim(rsreporte.Fields!nrodocumento)
          .Fields!NroHistoriaClinica = mo_ReporteUtil.NullToVacio(rsreporte.Fields!NroHistoriaClinica)
          .Fields!razonSocial = mo_ReporteUtil.NullToVacio(rsreporte.Fields!razonSocial)
          If rsreporte.Fields!Adelantos > 0 Then
            .Fields!Adelantos = rsreporte.Fields!Adelantos
          ElseIf rsreporte!idProducto = lnIdPagosACuenta Then
            .Fields!Adelantos = rsreporte.Fields!Subtotal
          Else
            .Fields!Adelantos = 0
          End If
          .Fields!importeEXO = IIf(IsNull(rsreporte.Fields!exoneraciones), 0, rsreporte.Fields!exoneraciones)
          .Fields!idEstadoComprobante = rsreporte.Fields!idEstadoComprobante
          .Fields!TotalPorPagar = IIf(IsNull(rsreporte.Fields!Subtotal), 0, rsreporte.Fields!Subtotal) ' IIf(IsNull(rsReporte.Fields!precio), 0, rsReporte.Fields!precio)
          .Fields!idProducto = rsreporte.Fields!idProducto
          .Fields!Subtotal = IIf(rsreporte.Fields!idEstadoComprobante = 9, 0, rsreporte.Fields!TotalPagado)
          .Fields!IGV = 0
          If rsreporte!idEstadoComprobante = 6 Then
             .Fields!TotalPagado = -rsreporte.Fields!TotalPagado
             .Fields!Devolucion = rsreporte.Fields!TotalPagado
          Else
             .Fields!TotalPagado = rsreporte.Fields!TotalPagado
             .Fields!Devolucion = 0
          End If
          .Fields!cajero = mo_ReglasFacturacion.SeleccionaDatosCajeroRpt(rsreporte.Fields!IdCajero, sghIniciales, oConexion)
          .Fields!FechaCobranza = rsreporte.Fields!FechaCobranza
          .Fields!anulado = IIf(rsreporte!idEstadoComprobante = 9, rsreporte!TotalPagado, 0)
          .Update
          rsreporte.MoveNext
        Loop
      End With
    End If
    rsreporte.Close
  End If
  
  'MEDICAMENTOS emitidos en CAJA SERVICIO
  If idTReporte = 0 Or idTReporte = 1 Then
    Set rsreporte = mo_ReglasFacturacion.SeleccionaFarmacia(FI, FF, IdCaja, IdTurno, idTComprobante, oConexion)
    If rsreporte.RecordCount > 0 Then
      With rsRsTmp
        rsreporte.MoveFirst
        Do While Not rsreporte.EOF
          .AddNew
          .Fields!nroSerie = rsreporte.Fields!nroSerie
          .Fields!nrodocumento = rsreporte.Fields!nrodocumento
          .Fields!boleta = Trim(rsreporte.Fields!nroSerie) & " - " & Trim(rsreporte.Fields!nrodocumento)
          .Fields!NroHistoriaClinica = mo_ReporteUtil.NullToVacio(rsreporte.Fields!NroHistoriaClinica)
          .Fields!razonSocial = mo_ReporteUtil.NullToVacio(rsreporte.Fields!razonSocial)
          .Fields!Adelantos = IIf(IsNull(rsreporte.Fields!Adelantos), 0, rsreporte.Fields!Adelantos)
          .Fields!importeEXO = IIf(IsNull(rsreporte.Fields!exoneraciones), 0, rsreporte.Fields!exoneraciones)
          .Fields!idEstadoComprobante = rsreporte.Fields!idEstadoComprobante
          .Fields!TotalPorPagar = IIf(IsNull(rsreporte.Fields!Subtotal), 0, rsreporte.Fields!Subtotal)
          .Fields!idProducto = rsreporte.Fields!idProducto
          .Fields!TotalPagado = rsreporte.Fields!TotalPagado
          .Fields!redondeo = mo_ReglasFacturacion.DevuelveRedondeoEnCadaBoletaFarmacia(rsreporte!TotalPagado, rsreporte!exoneraciones + rsreporte!Dctos, rsreporte!IdComprobantePago, rsreporte!idEstadoComprobante, oConexion.ConnectionString)
          .Fields!cajero = mo_ReglasFacturacion.SeleccionaDatosCajeroRpt(rsreporte.Fields!IdCajero, sghIniciales, oConexion)
          .Fields!FechaCobranza = rsreporte.Fields!FechaCobranza
          .Fields!Subtotal = IIf(rsreporte.Fields!idEstadoComprobante = 9, 0, rsreporte.Fields!TotalPagado / 1.19)
          .Fields!IGV = IIf(rsreporte.Fields!idEstadoComprobante = 9, 0, rsreporte.Fields!TotalPagado * 0.19 / 1.19)
          .Fields!anulado = IIf(rsreporte!idEstadoComprobante = 9, rsreporte!TotalPagado, 0)
          .Update
          rsreporte.MoveNext
        Loop
      End With
    End If
    rsreporte.Close
  End If

  'I = rsRsTmp.RecordCount
  Set rsreporte = rsRsTmp.Clone
  If rsreporte.RecordCount = 0 Then
    MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
  Else
    rsreporte.Sort = "FechaCobranza, nroSerie, nroDocumento"
    '
    iFila = 7
    lRecordCount = 0: lntImpSubTot = 0: lntImpAnul = 0: lntImpExo = 0: lntImpDevol = 0
    lntImpPagCta = 0: lntImpTot = 0: lntImpRedondeo = 0
    rsreporte.MoveFirst
    Do While Not rsreporte.EOF
      oWorkSheet.Cells(iFila, COL_FECHA).Value = "'" & Format(rsreporte.Fields!FechaCobranza, "dd/MM/yyyy")
      oWorkSheet.Cells(iFila, COL_USUARIO).Value = rsreporte!cajero
      oWorkSheet.Cells(iFila, COL_BOLETA).Value = rsreporte!nroSerie + " - " + rsreporte!nrodocumento
      oWorkSheet.Cells(iFila, COL_NRO_HISTORIA).Value = mo_ReporteUtil.NullToVacio(rsreporte!NroHistoriaClinica)
      oWorkSheet.Cells(iFila, COL_RAZON_SOCIAL).Value = mo_ReporteUtil.NullToVacio(rsreporte!razonSocial)
      lnImpSubTot = rsreporte!TotalPorPagar
      lnImpExo = rsreporte!importeEXO
      lnImpAnul = IIf(rsreporte!idEstadoComprobante = 9, rsreporte!TotalPagado, 0)
      lnImpPagCta = rsreporte!Adelantos
      lnSubTotal = rsreporte!Subtotal
      lnIGV = rsreporte!IGV
      lnImpTot = rsreporte!TotalPagado
      lnImpRedondeo = rsreporte!redondeo

      lRecordCount = lRecordCount + 1
      lcLlave = (rsreporte!nroSerie + rsreporte!nrodocumento)
      'lnImpAnul = 0: lnImpDevol = 0
      'lnImpExo = 0:  lnImpPagCta = 0: lnImpTot = 0
      'lnDctos = IIf(IsNull(rsReporte!adelantos), 0, rsReporte!adelantos) 'Descuentos en el COMPROBANTE PAGO
      Do While Not rsreporte.EOF And lcLlave = (rsreporte!nroSerie + rsreporte!nrodocumento)
        'lnImpExo = lnImpExo + IIf(IsNull(rsReporte!importeEXO), 0, rsReporte!importeEXO)             'Exonerado
        Select Case rsreporte!idEstadoComprobante
          Case 6 'Devolucion
            'lnImpDevol = lnImpDevol + rsReporte!totalPorPagar
          Case 9 'Anulado
            'lnImpAnul = lnImpAnul + rsReporte!totalPorPagar
          Case 4
            If rsreporte!idProducto = lnIdPagosACuenta Then 'Pago a cuenta
              'lnImpPagCta = lnImpPagCta + rsReporte!totalPorPagar
              'lnImpTot = lnImpTot + rsReporte!totalPorPagar
            Else
              'lnImpTot = lnImpTot + rsReporte!totalPorPagar     'solo pago
            End If
        End Select
        rsreporte.MoveNext
        If rsreporte.EOF Then Exit Do
      Loop
      'lnImpTot = lnImpTot - lnDctos
      ' lnImpSubTot = lnImpTot + lnImpAnul + lnImpExo + lnImpDevol
      If lnImpAnul > 0 Then
        lnImpTot = 0
      ElseIf lnImpDevol > 0 Then
        lnImpTot = 0
      End If
      oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lnImpSubTot
      oWorkSheet.Cells(iFila, COL_REDONDEO).Value = lnImpRedondeo
      oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lnImpExo
      oWorkSheet.Cells(iFila, COL_ANULADO).Value = lnImpAnul
      oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lnImpDevol
      oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lnImpPagCta
      oWorkSheet.Cells(iFila, COL_TOTAL_BRUTO).Value = lnSubTotal
      oWorkSheet.Cells(iFila, COL_IGV).Value = lnIGV
      oWorkSheet.Cells(iFila, COL_TOTAL_NETO).Value = lnImpTot

      lntImpSubTot = lntImpSubTot + lnImpSubTot
      If lnImpAnul = 0 Then lntImpExo = lntImpExo + lnImpExo
      lntImpAnul = lntImpAnul + lnImpAnul
      If lnImpAnul = 0 Then lntImpDevol = lntImpDevol + lnImpDevol
      lntImpTot = lntImpTot + lnImpTot
      If lnImpAnul = 0 Then lntImpPagCta = lntImpPagCta + lnImpPagCta
      iFila = iFila + 1
      lntImpRedondeo = lntImpRedondeo + lnImpRedondeo
      lnTSubTotal = lnTSubTotal + lnSubTotal
      lntIgv = lntIgv + lnIGV
    Loop
    iFila = iFila + 1
    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, COL_TOTAL_NETO
    oWorkSheet.Cells(iFila, 2).Value = "Cantidad de Documentos: " + Trim(Str(lRecordCount))
    oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lntImpSubTot
    oWorkSheet.Cells(iFila, COL_REDONDEO).Value = lntImpRedondeo
    oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lntImpExo
    oWorkSheet.Cells(iFila, COL_ANULADO).Value = lntImpAnul
    oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lntImpDevol
    oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lntImpPagCta
    oWorkSheet.Cells(iFila, COL_TOTAL_BRUTO).Value = lnTSubTotal
    oWorkSheet.Cells(iFila, COL_IGV).Value = lntIgv
    oWorkSheet.Cells(iFila, COL_TOTAL_NETO).Value = lntImpTot
    '
    oWorkSheet.PageSetup.PrintTitleRows = "$2:$6"
    If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
    oExcel.Visible = True
    oWorkSheet.PrintPreview
    'oWorkSheet.PrintOut
    'oWorkBook.Close SaveChanges:=False
    MsgBox "Reporte creado correctamente.", vbInformation, "Registro de ventas."
  End If
  oConexion.Close
  
    Set oWorkSheet = Nothing
    Set oExcel = Nothing
  Exit Sub

ManejadorError:
  Select Case Err.Number
    Case 1004
      MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
      MsgBox Err.Description
  End Select
  'Resume
  Exit Sub
End Sub

'debb-16/05/2016
Sub CrearReporteDeRegistroDeVentasDebb(FI As String, FF As String, IdCaja As Long, IdTurno As Long, _
                                       idTComprobante As Long, idTReporte As Long, lcNCajero As String, _
                                       lnIdVendedorFarmacia As Long, lbEnExcel As Boolean, lnIdFarmacia As Long, _
                                       dFarmacia As String, lbSoloAnuladas As Boolean)


    
  Dim rsreporte As New Recordset
  Dim oRsTmp1 As New Recordset
  Dim lcLlave As String
  Dim lnImpSubTot As Double, lntImpSubTot As Double
  Dim lnImpAnul As Double, lntImpAnul As Double
  Dim lnImpExo As Double, lntImpExo As Double
  Dim lnImpDevol As Double, lntImpDevol As Double
  Dim lnImpPagCta As Double, lntImpPagCta As Double
  Dim lnImpTot As Double, lntImpTot As Double: Dim lnDctos As Double
  Dim lnSubTotal As Double, lnIGV As Double, lnTSubTotal As Double, lntIgv As Double
  Dim iFila As Long, lnImpRedondeo As Double, lntImpRedondeo As Double
  Dim lRecordCount As Long, lcCadenaConexion As String
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  Dim lnIdPagosACuenta As Long
  'On Error GoTo ManejadorError
  Dim rsRsTmp As New Recordset 'Dim rsRsTmp As New ADODB.Recordset
  Dim mo_ReporteUtil As New ReporteUtil
  Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
  Dim oConexion As New ADODB.Connection
  Dim lbTienePagoAcuenta As Boolean, lnImporteTotalItems As Double
  Dim lnRecordCount As Long, f As Long, lcFiltrar As String
  Dim lcBoletaSerie As String, lcBoletaDocumento As String, lnIdComprobantePago As Long, lcSubTitulo As String
    
  Const COL_FECHA = 2
  Const COL_USUARIO = 3
  Const COL_BOLETA = 4
  Const COL_NRO_HISTORIA = 5
  Const COL_RAZON_SOCIAL = 6
  Const COL_SUBTOTAL = 7
  Const COL_REDONDEO = 8
  Const COL_ANULADO = 9
  Const COL_EXONERADO = 10
  Const COL_DEVOLUCION = 11
  Const COL_PAGOCTA = 12
  
  Const COL_TOTAL_BRUTO = 13
  Const COL_IGV = 14
  Const COL_TOTAL_NETO = 15
  
  lcBoletaSerie = "": lcBoletaDocumento = "": lnIdComprobantePago = 0
    
  lnIdPagosACuenta = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighEntidades.CadenaConexion
  
    
  With rsRsTmp
    .Fields.Append "FechaCobranza", adDate
    .Fields.Append "NroSerie", adVarChar, 5, adFldIsNullable
    .Fields.Append "NroDocumento", adVarChar, 20, adFldIsNullable
    .Fields.Append "NroHistoriaClinica", adVarChar, 50, adFldIsNullable
    .Fields.Append "RazonSocial", adVarChar, 100, adFldIsNullable
    .Fields.Append "adelantos", adDouble
    .Fields.Append "ImporteEXO", adDouble
    .Fields.Append "idEstadoComprobante", adUnsignedBigInt
    .Fields.Append "totalPorPagar", adDouble
    .Fields.Append "IdProducto", adUnsignedBigInt
    .Fields.Append "subtotal", adDouble
    .Fields.Append "igv", adDouble
    .Fields.Append "totalPagado", adDouble
    .Fields.Append "redondeo", adDouble
    .Fields.Append "cajero", adVarChar, 20, adFldIsNullable
    .Fields.Append "Boleta", adVarChar, 30, adFldIsNullable
    .Fields.Append "anulado", adDouble
    .Fields.Append "Devolucion", adDouble
    .LockType = adLockOptimistic
    .Open
  End With
  
  'SERVICIOS emitidos en CAJA SERVICIO
'  If idTReporte = 0 Or idTReporte = 2 Then
    Set rsreporte = mo_ReglasFacturacion.SeleccionaServiciosDebb(FI, FF, IdCaja, IdTurno, idTComprobante, _
                                                                 oConexion, Val(mo_cmbIdResponsable.BoundText), _
                                                                 ml_IdTipoReporte, IIf(Me.chkSoloCredito.Value = 1, True, False))
    If lbSoloAnuladas = True Then
       rsreporte.Filter = "IdEstadoComprobante = 9"
    End If
    lnRecordCount = rsreporte.RecordCount
    If lnRecordCount > 0 Then
      f = 0
      Me.progressRpt.Min = 0: Me.progressRpt.Max = lnRecordCount
      With rsRsTmp
        rsreporte.MoveFirst
        Do While Not rsreporte.EOF
                f = f + 1
                Me.progressRpt.Value = f
                DoEvents
                Me.Refresh
                '
                
                lRecordCount = 0
                If lnIdVendedorFarmacia > 0 And rsreporte.Fields!IdTipoOrden <> 1 Then
                    Set oRsTmp1 = mo_ReglasFacturacion.FarmaciaVendedorChequeaSiEsSuBoleta(rsreporte.Fields!IdComprobantePago, lnIdVendedorFarmacia)
                    If oRsTmp1.RecordCount = 0 Then
                       lRecordCount = 1
                    End If
                    oRsTmp1.Close
                End If
                If lRecordCount = 0 And lnIdFarmacia > 0 And rsreporte.Fields!IdTipoOrden <> 1 Then
                   If rsreporte.Fields!idFarmacia > 0 Then
                      If lnIdFarmacia <> rsreporte.Fields!idFarmacia Then
                         lRecordCount = 1
                      End If
                   Else
                      If mo_ReglasFacturacion.FarmaciaChequeaSiBoletaEsDeLaFarmacia(rsreporte.Fields!IdComprobantePago, lnIdFarmacia) = False Then
                         lRecordCount = 1
                      End If
                   End If
                End If
                If lRecordCount = 0 Then
If Val(rsreporte.Fields!nrodocumento) = 1372884 Then
lcBoletaSerie = ""
End If
                    lcBoletaSerie = rsreporte.Fields!nroSerie
                    lcBoletaDocumento = rsreporte.Fields!nrodocumento
                    lnIdComprobantePago = rsreporte.Fields!IdComprobantePago
                    '***debbsetiembre2014(inicio)
                    lbTienePagoAcuenta = mo_ReglasFacturacion.ChequeaSiEsPagosAcuenta(rsreporte!IdComprobantePago, _
                                                                oConexion, lnIdPagosACuenta, lnImporteTotalItems, _
                                                                rsreporte!IdTipoOrden, rsreporte!exoneraciones, _
                                                                rsreporte!Adelantos, rsreporte!idEstadoComprobante, _
                                                                rsreporte!TotalPagado)
                                                                
                    
                    If lnImporteTotalItems <> 9999 Then
                        .AddNew
                        .Fields!nroSerie = rsreporte.Fields!nroSerie
                        .Fields!nrodocumento = rsreporte.Fields!nrodocumento
                        .Fields!boleta = Trim(rsreporte.Fields!nroSerie) & " - " & Trim(rsreporte.Fields!nrodocumento)
                        If IsNull(rsreporte!NroHistoriaClinica) Then
                           .Fields!NroHistoriaClinica = ""
                        Else
                           .Fields!NroHistoriaClinica = HCigualDNI_DevuelveHistoriaConCerosIzquierda(rsreporte!NroHistoriaClinica, False)
                        End If
                        .Fields!razonSocial = mo_ReporteUtil.NullToVacio(rsreporte.Fields!razonSocial)
                        If rsreporte.Fields!Adelantos > 0 Then
                          .Fields!Adelantos = rsreporte.Fields!Adelantos
                        ElseIf lbTienePagoAcuenta And rsreporte.Fields!IdTipoOrden = 1 Then
                          .Fields!Adelantos = rsreporte.Fields!Subtotal
                        Else
                          .Fields!Adelantos = 0
                        End If
                        If rsreporte!idEstadoComprobante = 9 Then
                           .Fields!importeEXO = 0
                        Else
                           .Fields!importeEXO = IIf(IsNull(rsreporte.Fields!exoneraciones), 0, rsreporte.Fields!exoneraciones)
                        End If
                        .Fields!idEstadoComprobante = rsreporte.Fields!idEstadoComprobante
                        .Fields!TotalPorPagar = IIf(IsNull(rsreporte.Fields!Subtotal), 0, rsreporte.Fields!Subtotal) ' IIf(IsNull(rsReporte.Fields!precio), 0, rsReporte.Fields!precio)
                        '.Fields!idProducto = rsReporte.Fields!idProducto
                        If rsreporte.Fields!IdTipoOrden = 1 Then
                          .Fields!Subtotal = IIf(rsreporte.Fields!idEstadoComprobante = 9, 0, rsreporte.Fields!Subtotal)
                          .Fields!IGV = IIf(rsreporte.Fields!idEstadoComprobante = 9, 0, IIf(IsNull(rsreporte.Fields!IGV), 0, rsreporte.Fields!IGV))
                        Else
                          .Fields!Subtotal = IIf(rsreporte.Fields!idEstadoComprobante = 9, 0, rsreporte.Fields!Subtotal)
                          .Fields!IGV = IIf(rsreporte.Fields!idEstadoComprobante = 9, 0, IIf(IsNull(rsreporte.Fields!IGV), 0, rsreporte.Fields!IGV))
                        End If
                        If rsreporte!idEstadoComprobante = 6 Then
                           .Fields!TotalPagado = -rsreporte.Fields!TotalPagado
                           .Fields!Devolucion = rsreporte.Fields!TotalPagado
                        ElseIf rsreporte!idEstadoComprobante = 4 Then
                           .Fields!TotalPagado = rsreporte.Fields!TotalPagado
                           .Fields!Devolucion = 0
                        End If
                        .Fields!cajero = mo_ReglasFacturacion.SeleccionaDatosCajeroRpt(rsreporte.Fields!IdCajero, sghIniciales, oConexion)
                        .Fields!FechaCobranza = IIf(IsNull(rsreporte!FechaCobranza), rsreporte!fechaEmision, rsreporte!FechaCobranza)
                        .Fields!anulado = IIf(rsreporte!idEstadoComprobante = 9, rsreporte!TotalPagado, 0)
                        '***debbsetiembre2014(inicio)
                        '.Fields!redondeo = IIf(rsReporte!IdEstadoComprobante = 4 And rsReporte.Fields!IdTipoOrden <> 1 And (rsReporte.Fields!TotalPagado - lnImporteTotalItems) > 0, rsReporte.Fields!TotalPagado - lnImporteTotalItems, 0)
                        .Fields!redondeo = lnImporteTotalItems
                        '***debbsetiembre2014(final)
                        .Update
                    End If
                End If
                rsreporte.MoveNext
        Loop
      End With
    End If
    rsreporte.Close
 ' End If
 
  '******************Notas de crédito*****************    kike 2017
  Set rsreporte = mo_AdminCaja.NotaCreditoDevueltosPorNumYFecha("", "", CDate(FI), CDate(FF))
  lcFiltrar = ""
  If IdCaja > 0 Then
     lcFiltrar = lcFiltrar & "idCaja=" & IdCaja
  End If
  If Val(mo_cmbIdResponsable.BoundText) > 0 Then
  
     lcFiltrar = lcFiltrar & IIf(lcFiltrar = "", "", " and ") & "idCajero=" & Val(mo_cmbIdResponsable.BoundText)
  End If
  If optFarmacia.Value = True Then
     If lnIdFarmacia > 0 Then
        lcFiltrar = lcFiltrar & IIf(lcFiltrar = "", "", " and ") & "idFarmacia=" & lnIdFarmacia
     Else
        lcFiltrar = lcFiltrar & IIf(lcFiltrar = "", "", " and ") & "idFarmacia>0"
     End If
  End If
  If optServicio.Value = True Then
     lcFiltrar = lcFiltrar & IIf(lcFiltrar = "", "", " and ") & "idFarmacia=null"
  End If
  If lcFiltrar <> "" Then
     rsreporte.Filter = lcFiltrar
  End If
  If rsreporte.RecordCount > 0 Then
    rsreporte.MoveFirst
    Do While Not rsreporte.EOF
        With rsRsTmp
        .AddNew
        .Fields!FechaCobranza = rsreporte!fecha
        .Fields!nroSerie = rsreporte.Fields!nroSerie
        .Fields!nrodocumento = rsreporte.Fields!nrodocumento
        .Fields!boleta = Trim(rsreporte.Fields!nroSerie) & " - " & Trim(rsreporte.Fields!nrodocumento)
        .Fields!NroHistoriaClinica = "Not Cred"
        .Fields!razonSocial = Left(mo_ReporteUtil.NullToVacio(rsreporte.Fields!razonSocial), 100)
        .Fields!TotalPagado = -rsreporte.Fields!Total
        .Fields!cajero = mo_ReglasFacturacion.SeleccionaDatosCajeroRpt(rsreporte.Fields!IdCajero, sghIniciales, oConexion)
        .Update
        End With
         rsreporte.MoveNext
    Loop
  End If
  rsreporte.Close

  Set rsreporte = rsRsTmp.Clone
  If rsreporte.RecordCount = 0 Then
    MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
  Else
    rsreporte.Sort = "FechaCobranza, nroSerie, nroDocumento"
    '
    If lbEnExcel = False Then
'rsRsTmp.Save "c:\prueba.xlm", adPersistXML
'Recordset_a_Xml rsRsTmp, "c:\prueba.xml"
    
        Set EVentas.DataSource = rsreporte
        EVentas.RightMargin = 10
        EVentas.TopMargin = 10
        EVentas.LeftMargin = 10
        EVentas.BottomMargin = 10
        EVentas.ReportWidth = 9945
        
        EVentas.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
        EVentas.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
        EVentas.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
        EVentas.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
        EVentas.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
        EVentas.Sections("cabecera").Controls("lblPc").Caption = "PC: " & sighEntidades.RetornaNombrePC
        EVentas.Sections("cabecera").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighEntidades.Usuario)
        EVentas.Sections("cabecera").Controls("lblSubTitulo").Caption = "Del: " & FI & " al " & FF & ". " & ms_IdCaja & ms_IdTurno & _
                                                              ms_idTComprobante & ms_IdTipoReporte & " " & lcNCajero & _
                    IIf(Val(mo_cmdVendorFarmacia.BoundText) > 0, " (Vendedor Farmacia: " & Trim(cmdVendorFarmacia.Text) & ")", "") & _
                                                              dFarmacia & IIf(lbSoloAnuladas = True, " (solo ANULADAS)", "") & IIf(Me.chkSoloCredito.Value = 1, "  (solo con CREDITOS)", "  (sin considerar CREDITOS)")
        Set EVentas.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
        EVentas.Sections("pie").Controls("lblPie").Caption = "Cantidad de documentos: " & Trim(Str(rsreporte.RecordCount))
        EVentas.Orientation = rptOrientPortrait
        If lcBuscaParametro.SeleccionaFilaParametro(534) = "S" Then
           EVentas.Show 1
        Else
           EVentas.PrintReport True    ' EVentas.Show 1
        End If

        'debb-27/05/2015
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        mo_ReglasComunes.grabaTablaAuditoria ("EVentas: " & _
                                       FI & " " & FF)
        Set mo_ReglasComunes = Nothing
        '
        '
        MsgBox "Reporte creado correctamente.", vbInformation, "Registro de ventas."
    Else
        
       lcSubTitulo = "Del: " & FI & " al " & FF & ". " & ms_IdCaja & ms_IdTurno & ms_idTComprobante & ms_IdTipoReporte & " " & lcNCajero & IIf(Val(mo_cmdVendorFarmacia.BoundText) > 0, " (Vendedor Farmacia: " & Trim(cmdVendorFarmacia.Text) & ")", "") & dFarmacia
        
        Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
        mo_ReglasReportes.ExportarRecordSetAexcelFast rsreporte, "REPORTE DE VENTAS", lcSubTitulo, "total", Me.hwnd
        'mo_ReglasReportes.ExportarRecordSetAexcel rsreporte, "Reporte de Ventas", lcSubTitulo, "", Me.hwnd
        Set mo_ReglasReportes = Nothing
        
    
    End If
  End If
  oConexion.Close
  
  Exit Sub

ManejadorError:
  Select Case Err.Number
    Case 1004
      MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
      MsgBox Err.Description & Chr(13) & "idComprobantePago: " & lnIdComprobantePago & Chr(13) & _
             "Serie: " & lcBoletaSerie & Chr(13) & "B.Numero: " & lcBoletaDocumento

  End Select
  'Resume
  Exit Sub


End Sub


Private Sub txtFechaInicio_LostFocus()
If Not IsDate(txtFechaInicio.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaInicio.Text = sighEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub





'modificado por Samuel 07/08
Sub CrearReporteDeCierreCaja(FI As String, FF As String, IdCaja As Long, IdTurno As Long, idTReporte As Long, lcNCajero As Long, lbEnExcel As Boolean, idFarmacia As Long, idVendedor As Long, idTComprobante As Long)
    Dim rsreporte As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp123 As New Recordset
    Dim oCantidad As Long
    Dim oTotal As Double
'    Dim oRedondeoA As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    
On Error GoTo ManejadorError
    Dim rsRsTmp As New Recordset
    Dim rsRsTmpRed As New ADODB.Recordset
    Dim mo_ReporteUtil As New ReporteUtil
    Dim oConexion As New Connection
    Dim lnRecordCount As Long, f As Long
    
    Dim dl_SubTotal As Double
    Dim dl_IGV As Double
    Dim dl_Total As Double
    Dim dl_Exoneraciones As Double
    Dim dl_Adelantos As Double
    Dim dl_Anulaciones As Double
    Dim dl_Devoluciones As Double
    Dim dl_Pagado As Double
    Dim dl_Redondeo As Double
    
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient

    With rsRsTmp
      .Fields.Append "Nombre", adVarChar, 300
      .Fields.Append "Cantidad", adInteger
      .Fields.Append "Total", adDouble
      .Fields.Append "codigo", adVarChar, 20
      .Fields.Append "Tipo", adVarChar, 20
      .LockType = adLockOptimistic
      .Open
    End With
    dl_SubTotal = 0
    dl_IGV = 0
    dl_Total = 0
    dl_Exoneraciones = 0
    dl_Adelantos = 0
    dl_Anulaciones = 0
    dl_Devoluciones = 0
    dl_Redondeo = 0
    dl_Pagado = 0
    oCantidad = 0

    If ml_IdTipoReporte = 2 Or ml_IdTipoReporte = 0 Then
        Set rsreporte = mo_ReglasFacturacion.SeleccionaCierreCaja(FI, FF, IdCaja, IdTurno, 1, lcNCajero, idTComprobante, _
                                                             idFarmacia, idVendedor, IIf(chkSoloCredito.Value = 1, True, False))
        Set oRsTmp = mo_ReglasFacturacion.SeleccionaCierreCajaTotales(FI, FF, IdCaja, IdTurno, 1, lcNCajero, idTComprobante, _
                                                            idFarmacia, idVendedor, IIf(chkSoloCredito.Value = 1, True, False))
        Set rsRsTmpRed = mo_ReglasFacturacion.SeleccionaCierreCajaServiciosServiciosHospitalarios(FI, FF, IdCaja, IdTurno, lcNCajero, idTComprobante, lcBuscaParametro.SeleccionaFilaParametro(245))
        If rsRsTmpRed.RecordCount > 0 Then dl_Adelantos = IIf(IsNull(rsRsTmpRed.Fields!ServiciosHospitalarios), 0, rsRsTmpRed.Fields!ServiciosHospitalarios)
        rsRsTmpRed.Close
        If oRsTmp.RecordCount > 0 Then
            dl_SubTotal = IIf(IsNull(oRsTmp.Fields!Subtotal), 0, oRsTmp.Fields!Subtotal)
            dl_IGV = IIf(IsNull(oRsTmp.Fields!IGV), 0, oRsTmp.Fields!IGV)
            dl_Total = IIf(IsNull(oRsTmp.Fields!Total), 0, oRsTmp.Fields!Total)
            dl_Exoneraciones = IIf(IsNull(oRsTmp.Fields!exoneraciones), 0, oRsTmp.Fields!exoneraciones)
            dl_Adelantos = IIf(IsNull(oRsTmp.Fields!Adelantos), dl_Adelantos, dl_Adelantos + oRsTmp.Fields!Adelantos)
            dl_Anulaciones = IIf(IsNull(oRsTmp.Fields!anulado), 0, oRsTmp.Fields!anulado)
            dl_Devoluciones = IIf(IsNull(oRsTmp.Fields!Devolucion), 0, oRsTmp.Fields!Devolucion)
            dl_Pagado = IIf(IsNull(oRsTmp.Fields!pagado), 0, oRsTmp.Fields!pagado)
         End If
        Set rsRsTmpRed = mo_ReglasFacturacion.SeleccionaCierreCajaServiciosRedondeo(FI, FF, IdCaja, IdTurno, lcNCajero, idTComprobante)
        If rsRsTmpRed.RecordCount > 0 Then
            If Not IsNull(rsRsTmpRed.Fields!redondeo) Then
                dl_Redondeo = rsRsTmpRed.Fields!redondeo - (dl_Exoneraciones)
            End If
        End If
        rsRsTmpRed.Close
        lnRecordCount = rsreporte.RecordCount
        If lnRecordCount > 0 Then
            f = 0
            Me.progressRpt.Min = 0: Me.progressRpt.Max = lnRecordCount
            rsreporte.MoveFirst
            Do While Not rsreporte.EOF
                    f = f + 1
                    Me.progressRpt.Value = f
                    DoEvents
                    Me.Refresh
                    rsRsTmp.AddNew
                    rsRsTmp.Fields!Codigo = rsreporte.Fields!Codigo
                    rsRsTmp.Fields!nombre = rsreporte.Fields!nombre
                    rsRsTmp.Fields!Cantidad = rsreporte.Fields!Cantidad
                    rsRsTmp.Fields!Total = rsreporte.Fields!Total
                    rsRsTmp.Fields!tipo = " "
                    oCantidad = oCantidad + rsreporte.Fields!Cantidad
                    oTotal = oTotal + rsreporte.Fields!Total
                    rsreporte.MoveNext
            Loop
'            oTotal = oTotal - dl_IGV
        End If
    End If
    If ml_IdTipoReporte = 1 Or ml_IdTipoReporte = 0 Then
        Set rsreporte = mo_ReglasFacturacion.SeleccionaCierreCaja(FI, FF, IdCaja, IdTurno, 2, lcNCajero, idTComprobante, idFarmacia, idVendedor, IIf(chkSoloCredito.Value = 1, True, False))
        Set oRsTmp = mo_ReglasFacturacion.SeleccionaCierreCajaTotales(FI, FF, IdCaja, IdTurno, 2, lcNCajero, idTComprobante, idFarmacia, idVendedor, IIf(chkSoloCredito.Value = 1, True, False))
        'Set rsRsTmpRed = mo_ReglasFacturacion.SeleccionaCierreCajaRedondeo(FI, FF, IdCaja, IdTurno, 2, lcNCajero, idTComprobante, idFarmacia, idVendedor)
        
        'oRedondeoA = IIf(IsNull(rsRsTmpRed.Fields!Total), 0, rsRsTmpRed.Fields!Total) - IIf(IsNull(rsRsTmpRed.Fields!totalpagar), 0, rsRsTmpRed.Fields!totalpagar)
        If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            dl_SubTotal = dl_SubTotal + IIf(IsNull(oRsTmp.Fields!Subtotal), 0, oRsTmp.Fields!Subtotal)
            dl_IGV = dl_IGV + IIf(IsNull(oRsTmp.Fields!IGV), 0, oRsTmp.Fields!IGV)
            dl_Total = dl_Total + IIf(IsNull(oRsTmp.Fields!Total), 0, oRsTmp.Fields!Total)
            dl_Exoneraciones = dl_Exoneraciones + IIf(IsNull(oRsTmp.Fields!exoneraciones), 0, oRsTmp.Fields!exoneraciones)
            dl_Adelantos = dl_Adelantos + IIf(IsNull(oRsTmp.Fields!Adelantos), 0, oRsTmp.Fields!Adelantos)
            dl_Anulaciones = dl_Anulaciones + IIf(IsNull(oRsTmp.Fields!anulado), 0, oRsTmp.Fields!anulado)
            dl_Devoluciones = dl_Devoluciones + IIf(IsNull(oRsTmp.Fields!Devolucion), 0, oRsTmp.Fields!Devolucion)
            dl_Pagado = dl_Pagado + IIf(IsNull(oRsTmp.Fields!pagado), 0, oRsTmp.Fields!pagado)
         End If
        Set rsRsTmpRed = mo_ReglasFacturacion.SeleccionaCierreCajaInsumosRedondeo(FI, FF, IdCaja, IdTurno, lcNCajero, idTComprobante, idFarmacia, idVendedor)
        If rsRsTmpRed.RecordCount > 0 Then
            If Not IsNull(oRsTmp.Fields!exoneraciones) And Not IsNull(rsRsTmpRed.Fields!redondeo) And Not IsNull(oRsTmp.Fields!Adelantos) Then
               dl_Redondeo = Abs(dl_Redondeo + Abs(rsRsTmpRed.Fields!redondeo) - IIf(IsNull(oRsTmp.Fields!exoneraciones), 0, oRsTmp.Fields!exoneraciones) - IIf(IsNull(oRsTmp.Fields!Adelantos), 0, oRsTmp.Fields!Adelantos))
            End If
        End If
        lnRecordCount = rsreporte.RecordCount
        Dim lnTotalInsumos As Double, lnTotalMedicinas As Double
        lnTotalInsumos = 0
        lnTotalMedicinas = 0
        If lnRecordCount > 0 Then
            f = 0
            Me.progressRpt.Min = 0: Me.progressRpt.Max = lnRecordCount
            rsreporte.MoveFirst
            Do While Not rsreporte.EOF
                    If oRsTmp123.State = 1 Then oRsTmp123.Close
                    Set oRsTmp123 = mo_ReglasFacturacion.FactCatalogoBienesInsumosSeleccionarXcodigo(rsreporte.Fields!Codigo, oConexion)
                    f = f + 1
                    Me.progressRpt.Value = f
                    DoEvents
                    Me.Refresh
                    rsRsTmp.AddNew
                    rsRsTmp.Fields!Codigo = rsreporte.Fields!Codigo
                    rsRsTmp.Fields!nombre = rsreporte.Fields!nombre
                    rsRsTmp.Fields!Cantidad = rsreporte.Fields!Cantidad
                    rsRsTmp.Fields!Total = rsreporte.Fields!Total
                    rsRsTmp.Fields!tipo = IIf(IsNull(oRsTmp123!TipoProducto), " ", _
                                          IIf(oRsTmp123!TipoProducto = "1", "INSUMO", "MEDICAMENTO"))
                    oCantidad = oCantidad + rsreporte.Fields!Cantidad
                    oTotal = oTotal + rsreporte.Fields!Total
                    rsreporte.MoveNext
            Loop
            'oTotal = oTotal + oRedondeoA
        End If
    End If
    lnRecordCount = rsRsTmp.RecordCount
    If lnRecordCount > 0 Then
        f = 0
        Me.progressRpt.Min = 0: Me.progressRpt.Max = lnRecordCount
        rsRsTmp.MoveFirst
        Do While Not rsRsTmp.EOF
                Select Case rsRsTmp!tipo
                Case "INSUMO"
                    lnTotalInsumos = lnTotalInsumos + rsRsTmp!Total
                Case "MEDICAMENTO"
                    lnTotalMedicinas = lnTotalMedicinas + rsRsTmp!Total
                End Select
                
                f = f + 1
                Me.progressRpt.Value = f
                DoEvents
                Me.Refresh
                rsRsTmp.MoveNext
        Loop
    End If

  If rsRsTmp.RecordCount = 0 Then
    MsgBox "No existe información con esos Datos", vbOKOnly + vbInformation, Me.Caption
  Else
    
    If lbEnExcel = False Then
        
    
'        Dim lcRutaExportar As String
'
'        lcRutaExportar = Trim(lcBuscaParametro.SeleccionaFilaParametro(313))
'        If Right(lcRutaExportar, 1) <> "\" Then
'           lcRutaExportar = lcRutaExportar & "\"
'        End If
        
        Dim lbSePuedeImprimirPDF As Boolean, lcArchivoPDF As String
        If UcTipoImpresion1.OpcionImpresionElejida = sghTIpdf Then
           lcArchivoPDF = sighEntidades.DevuelveRutaConSlashInvertida(lcBuscaParametro.SeleccionaFilaParametro(313)) & "CajaXitems.pdf"
           If SePuedeImprimirPDF(lcArchivoPDF, True) = True Then
              lbSePuedeImprimirPDF = True
           End If
        End If
        
        Set ECaja.DataSource = rsRsTmp
        ECaja.RightMargin = 10
        ECaja.TopMargin = 50
        ECaja.LeftMargin = 10
        ECaja.BottomMargin = 50
        ECaja.Sections("Sección2").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
        ECaja.Sections("Sección2").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
        ECaja.Sections("Sección2").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
        ECaja.Sections("Sección2").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
        ECaja.Sections("Sección2").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
        ECaja.Sections("Sección2").Controls("lblPc").Caption = "PC: " & sighEntidades.RetornaNombrePC
        ECaja.Sections("Sección2").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighEntidades.Usuario)
        ECaja.Sections("Sección5").Controls("lblSubTotal").Caption = "Total Bruto : " & Format(dl_SubTotal, "0.00")
        ECaja.Sections("Sección5").Controls("lblTotal").Caption = "Total    : " & Format((dl_Total - dl_Anulaciones), "0.00")
        ECaja.Sections("Sección5").Controls("lblIGV").Caption = "IGV      : " & Format(dl_IGV, "0.00")
        ECaja.Sections("Sección5").Controls("lblExoneraciones").Caption = "Exoneraciones : " & Format(dl_Exoneraciones, "0.00")
        ECaja.Sections("Sección5").Controls("lblAdelantos").Caption = "Adelantos     : " & Format(dl_Adelantos, "0.00")
        ECaja.Sections("Sección5").Controls("lblRedondeo").Caption = "Redondeo      : " & Format(dl_Redondeo, "0.00")
        ECaja.Sections("Sección5").Controls("lblAnulaciones").Caption = "Anulaciones   : " & Format(dl_Anulaciones, "0.00")
        ECaja.Sections("Sección5").Controls("lblDevoluciones").Caption = "Devoluciones  : " & Format(dl_Devoluciones, "0.00")
        'ECaja.Sections("Sección5").Controls("lblPagado").Caption = "Pagado        : " & Format(dl_Pagado, "0.00")
        ECaja.Sections("Sección2").Controls("lblSubTitulo").Caption = "Del: " & FI & " al " & FF & ". " & ms_IdCaja & ms_IdTurno & ms_idTComprobante & ms_IdTipoReporte & " " & lcNCajero & IIf(Val(mo_cmdVendorFarmacia.BoundText) > 0, " (Vendedor Farmacia: " & Trim(cmdVendorFarmacia.Text) & ")", "") & IIf(Me.chkSoloCredito.Value = 1, " (solo CREDITOS)", " (sin considerar CREDITOS)")
        ECaja.Sections("Sección5").Controls("EtMedicamentosInsumos").Caption = "Medicamentos: " & Format(lnTotalMedicinas, "0.00") & _
                                                                               " <> Insumos: " & Format(lnTotalInsumos, "0.00")
        ECaja.Orientation = rptOrientPortrait
        If lbSePuedeImprimirPDF = True Then
           ECaja.PrintReport False
        Else
           ECaja.Show 1
        End If
        'debb-27/05/2015
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        mo_ReglasComunes.grabaTablaAuditoria ("ECaja: " & _
                                       FI & " " & FF)
        Set mo_ReglasComunes = Nothing
        '
        'MsgBox "Reporte creado correctamente.", vbInformation, "Registro de ventas."
        
        If lbSePuedeImprimirPDF = True Then
           
           MsgBox "Se creó archivo : " & lcArchivoPDF
        Else
           MsgBox "Reporte creado correctamente.", vbInformation, "Registro de ventas."
        End If
        SeteaOtraImpresoraDefault sighEntidades.ImpresoraDefaultDeEstaPC
    Else
        
        Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes, lcSubTitulo As String
        mo_ReglasReportes.ExportarRecordSetAexcel rsRsTmp, "Reporte de Ventas", lcSubTitulo, "", Me.hwnd
    
    
    
    End If
  End If
  oConexion.Close
  
  Exit Sub

ManejadorError:
  Select Case Err.Number
    Case 1004
      MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
      MsgBox Err.Description
  End Select
  Exit Sub
  Resume
End Sub


'debb-13/07/2015
Private Sub btnProdConsulta_Click()
        Dim oFrm As New SIGHNegocios.BuscaServicio
        oFrm.MostrarFormulario
        If oFrm.idRegistroSeleccionado <> 0 Then
            Me.txtCodProductoConsulta.Tag = CStr(oFrm.idRegistroSeleccionado)
            Call ObtenerNombreServicio(oFrm.idRegistroSeleccionado, Me.txtCodProductoConsulta, Me.txtProductoConsulta)
        End If
        Set oFrm = Nothing
End Sub
Sub ObtenerNombreServicio(IdServicio As Long, txtCode As TextBox, txtName As TextBox)
    Dim dOServ As New DOCatalogoServicio
    Set dOServ = mo_ReglasFacturacion.CatalogoServiciosSeleccionarPorId(IdServicio)
    If Not dOServ Is Nothing Then
        txtCode.Text = dOServ.Codigo
        txtName.Text = dOServ.nombre
    End If
    Set dOServ = Nothing
End Sub

Private Sub optReporteX_Click(Index As Integer)
    chkSoloAnuladas.Visible = False
    chkSoloAnuladas.Value = 0
    chkExcel.Visible = True
    UcTipoImpresion1.Visible = False
    chkSoloCredito.Visible = False
    chkSoloCredito.Value = 0
    Select Case Index
    Case 0, 1
         FraXitem.Visible = False
         Frame4.Visible = True
         FraFiltros.Visible = True
         chkExcel.Value = 0
         If Index = 1 Then
            chkSoloAnuladas.Visible = True
            chkSoloCredito.Visible = True
         Else
            chkExcel.Visible = False
            UcTipoImpresion1.Visible = True
         End If
    Case 2
         optFarmaciaServicios.Value = True
         optFarmaciaServicios_Click 1
         FraXitem.Visible = True
         Frame4.Visible = False
         FraFiltros.Visible = False
         chkExcel.Value = 0
         FraXitem.Top = Frame4.Top
         FraXitem.Left = Frame4.Left
         FraXitem.Width = Frame4.Width
    End Select
End Sub


'debb-17/03/2016
Sub CrearReporteXitem(FI As String, FF As String, lbEnExcel As Boolean, lcCpt As String, _
                      lcCptDescripcion As String, lnIdCpt As Long)
  Dim rsreporte As New Recordset
  Dim oRsTmp1 As New Recordset
  Dim lcLlave As String
  Dim lnImpSubTot As Double, lntImpSubTot As Double
  Dim lnImpAnul As Double, lntImpAnul As Double
  Dim lnImpExo As Double, lntImpExo As Double
  Dim lnImpDevol As Double, lntImpDevol As Double
  Dim lnImpPagCta As Double, lntImpPagCta As Double
  Dim lnImpTot As Double, lntImpTot As Double: Dim lnDctos As Double
  Dim lnSubTotal As Double, lnIGV As Double, lnTSubTotal As Double, lntIgv As Double
  Dim iFila As Long, lnImpRedondeo As Double, lntImpRedondeo As Double
  Dim lRecordCount As Long, lcCadenaConexion As String
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  Dim lnIdPagosACuenta As Long
  'On Error GoTo ManejadorError
  Dim rsRsTmp As New Recordset 'Dim rsRsTmp As New ADODB.Recordset
  Dim mo_ReporteUtil As New ReporteUtil
  Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
  Dim oConexion As New ADODB.Connection
  Dim lbTienePagoAcuenta As Boolean, lnImporteTotalItems As Double
  Dim lnRecordCount As Long, f As Long
  Dim lcBoletaSerie As String, lcBoletaDocumento As String, lnIdComprobantePago As Long
  Dim lnExoneracion As Double, lcSubTitulo As String
    
 
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighEntidades.CadenaConexion
  
    
  With rsRsTmp
    .Fields.Append "Nro", adInteger
    .Fields.Append "Boleta", adVarChar, 15, adFldIsNullable
    .Fields.Append "Fecha", adVarChar, 16, adFldIsNullable
    .Fields.Append "NroHistoria", adVarChar, 10, adFldIsNullable
    .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
    .Fields.Append "cantidad", adInteger
    .Fields.Append "precio", adDouble
    .Fields.Append "subTotal", adDouble
    .Fields.Append "exoneracion", adDouble
    .Fields.Append "total", adDouble
    .LockType = adLockOptimistic
    .Open
  End With
  Set rsreporte = mo_AdminCaja.CajaComprobantesPagoSoloServicioXitem(Val(txtCodProductoConsulta.Tag), CDate(FI), CDate(FF))
  rsreporte.Filter = IIf(Me.chkSoloCredito.Value <> 1, "tieneCredito=null", "tieneCredito<>null")
  lnRecordCount = rsreporte.RecordCount
  If lnRecordCount > 0 Then
      f = 0
      Me.progressRpt.Min = 0: Me.progressRpt.Max = lnRecordCount
      With rsRsTmp
        rsreporte.MoveFirst
        Do While Not rsreporte.EOF
                f = f + 1
                Me.progressRpt.Value = f
                DoEvents
                Me.Refresh
                '
                lnExoneracion = 0
                If rsreporte!exoneraciones > 0 Then
                   If IsNull(rsreporte!idCuentaAtencion) Then
                      lnExoneracion = mo_AdminCaja.BoletaCajaProrrateaExoneracionXitem(rsreporte!TotalBoleta, _
                                                                                       rsreporte!exoneraciones, rsreporte!Total)
                   Else
                      lnExoneracion = mo_AdminCaja.BoletaCajaCalculaExoneracionXitem(rsreporte!TotalBoleta, _
                                                                                   rsreporte!exoneraciones, _
                                                                                   rsreporte!idProducto, rsreporte!IdOrden, _
                                                                                   rsreporte!IdComprobantePago, _
                                                                                   rsreporte!idCuentaAtencion, oConexion, _
                                                                                   rsreporte!Total)
                   End If
                End If
                .AddNew
                .Fields!nro = f
                .Fields!boleta = Trim(rsreporte!nroSerie) + "-" + rsreporte!nrodocumento
                .Fields!fecha = Format(rsreporte!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
                If Not IsNull(rsreporte!NroHistoriaClinica) Then
                   .Fields!NroHistoria = Trim(Str(rsreporte!NroHistoriaClinica))
                Else
                   .Fields!NroHistoria = " "
                End If
                .Fields!Paciente = Left(rsreporte!razonSocial, 40)
                .Fields!Cantidad = rsreporte!Cantidad
                .Fields!precio = rsreporte!precio
                .Fields!Subtotal = rsreporte!Total
                .Fields!exoneracion = lnExoneracion
                .Fields!Total = rsreporte!Total - lnExoneracion
                .Update
                rsreporte.MoveNext
        Loop
      End With
  End If
  If rsRsTmp.RecordCount = 0 Then
    MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
  Else
    lcSubTitulo = "Cpt: " & Trim(txtCodProductoConsulta.Text) & " - " & Trim(txtProductoConsulta.Text) & _
                  "   (Del: " & FI & " al " & FF & ") " & IIf(Me.chkSoloCredito.Value = 1, " (solo CREDITOS)", " (sin considerar CREDITOS)")
    If lbEnExcel Then
        Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
        mo_ReglasReportes.ExportarRecordSetAexcel rsRsTmp, "REPORTE POR PROCEDIMIENTO (CPT) PAGADOS EN CAJA", lcSubTitulo, "", Me.hwnd, True, True
        Set mo_ReglasReportes = Nothing
    Else
        Set EVentasCPT.DataSource = rsRsTmp
        EVentasCPT.RightMargin = 10
        EVentasCPT.TopMargin = 10
        EVentasCPT.LeftMargin = 10
        EVentasCPT.BottomMargin = 10
        EVentasCPT.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
        EVentasCPT.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
        EVentasCPT.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
        EVentasCPT.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
        EVentasCPT.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
        EVentasCPT.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
        Set EVentasCPT.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
        EVentasCPT.Sections("pie").Controls("lblPie").Caption = "Cantidad: " & Trim(Str(rsreporte.RecordCount))
        EVentasCPT.Orientation = rptOrientPortrait
        EVentasCPT.Show 1
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        mo_ReglasComunes.grabaTablaAuditoria ("EVentas x cpt: " & _
                                       FI & " " & FF)
        Set mo_ReglasComunes = Nothing
    
    End If
  End If
  oConexion.Close
  
  Exit Sub

ManejadorError:
  Select Case Err.Number
    Case 1004
      MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
      MsgBox Err.Description & Chr(13) & "idComprobantePago: " & lnIdComprobantePago & Chr(13) & _
             "Serie: " & lcBoletaSerie & Chr(13) & "B.Numero: " & lcBoletaDocumento

  End Select
  'Resume
  Exit Sub
End Sub


Sub Recordset_a_Xml(oRecordset As Recordset, Path_XML As String)
On Error GoTo errSub
'Variables para la conexión ado, el recordset _
 y el objeto para generar el xml
Dim obj_DOMDocument As DOMDocument
    ' Graba el contenido del Recordset en el Obj DOMDocument.
    Set obj_DOMDocument = New DOMDocument
    oRecordset.Save obj_DOMDocument, adPersistXML
    ' Genera el archivo xml
    obj_DOMDocument.Save Path_XML
    Exit Sub
errSub:
      MsgBox "Error:" & Err.Number & vbNewLine & _
       "Descripción:" & Err.Description, vbCritical
End Sub


