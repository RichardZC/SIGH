VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form rProducPagoDeuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Producción, Pagos y Deudas  por Fechas"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "rProducPagoDeuda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   9195
      Begin VB.ComboBox cmbIdPuntoDeCarga 
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
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   630
         Width           =   2235
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
         Height          =   315
         Left            =   7920
         Picture         =   "rProducPagoDeuda.frx":0CCA
         TabIndex        =   10
         Top             =   570
         Width           =   1125
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
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
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   6900
         TabIndex        =   2
         Top             =   210
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
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   8280
         TabIndex        =   3
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pto. Carga"
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
         TabIndex        =   12
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   6390
         TabIndex        =   9
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
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
         TabIndex        =   8
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   5
      Top             =   2520
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rProducPagoDeuda.frx":0FDC
         DownPicture     =   "rProducPagoDeuda.frx":143C
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
         Left            =   3210
         Picture         =   "rProducPagoDeuda.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rProducPagoDeuda.frx":1D26
         DownPicture     =   "rProducPagoDeuda.frx":21EA
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
         Left            =   4740
         Picture         =   "rProducPagoDeuda.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "rProducPagoDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Producción por Pago Deuda
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
'***************daniel barrantes**************
'***************Registro de datos de filtro para el Reporte
'***************Historias Clinicas que pasadas 24 horas no regresan al ARCHIVO
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_cmbIdPuntoCarga As New SIGHEntidades.ListaDespleglable
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_idUsuario As Long

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptClaseCry As New rCrystal
        oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
        oRptClaseCry.FechaInicio = Format(txtFdesde.Text & " " & txtHrInicio.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
        oRptClaseCry.FechaFin = Format(txtFhasta.Text & " " & txtHrFin.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
        oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
        oRptClaseCry.IdPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
        oRptClaseCry.TipoReporte = Me.Name
        oRptClaseCry.Show vbModal
        Set oRptClaseCry = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:   F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ")     Punto Carga: " & cmbIdPuntoDeCarga.Text
    If cmbIdPuntoDeCarga.Text = "" Then
        sMensaje = sMensaje + "Por favor elija el Punto de Carga" + Chr(13)
        cmbIdPuntoDeCarga.SetFocus
    End If
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub











Private Sub Form_Initialize()
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga

End Sub

Sub InicializaFechaHora()
    txtFdesde.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    txtFhasta.Text = Date
    txtHrInicio.Text = "00:00"
    txtHrFin.Text = "23:59"

End Sub

Private Sub Form_Load()
    InicializaFechaHora
    '
    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCargaSegunFiltro("idUPS=1")
    Dim rsIdAlmacen As Recordset
    Set rsIdAlmacen = mo_reglasComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghImageneología, ml_idUsuario)
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbIdPuntoCarga.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
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






Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_reglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_LostFocus()
If Not SIGHEntidades.ValidaHora(txtHrFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtHrInicio_LostFocus()
If Not SIGHEntidades.ValidaHora(txtHrInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub
