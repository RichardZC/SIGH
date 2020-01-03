VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ImagIngresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13755
   Icon            =   "ImagIngresos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   13
      Top             =   7590
      Width           =   13710
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ImagIngresos.frx":0CCA
         DownPicture     =   "ImagIngresos.frx":112A
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
         Picture         =   "ImagIngresos.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ImagIngresos.frx":1A14
         DownPicture     =   "ImagIngresos.frx":1ED8
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
         Picture         =   "ImagIngresos.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   1365
      End
   End
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
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   13755
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
         Left            =   5700
         TabIndex        =   2
         Top             =   630
         Width           =   3240
      End
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
         Left            =   10800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   630
         Width           =   2805
      End
      Begin VB.TextBox txtEstado 
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
         Left            =   11940
         MaxLength       =   30
         TabIndex        =   6
         Top             =   240
         Width           =   1665
      End
      Begin VB.TextBox txtNmovimiento 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   5
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox txtNdocumento 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   1
         Top             =   690
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   5700
         TabIndex        =   7
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
         Left            =   9855
         TabIndex        =   16
         Top             =   690
         Width           =   855
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
         Height          =   240
         Left            =   11310
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   285
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
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
         Left            =   4800
         TabIndex        =   10
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° Documento"
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
         TabIndex        =   9
         Top             =   735
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Recepciona"
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
         Left            =   4680
         TabIndex        =   8
         Top             =   690
         Width           =   930
      End
   End
   Begin SIGHImagen.ucServicios ucProductos 
      Height          =   5715
      Left            =   30
      TabIndex        =   3
      Top             =   1740
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   5477
   End
End
Attribute VB_Name = "ImagIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento Ingresos de Placas
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdMovimiento As Long
Dim mi_Opcion As sghOpciones
Dim ms_MensajeError As String
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_cmbIdEstado As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdPuntoCarga As New SIGHEntidades.ListaDespleglable
Dim mo_cmbResponsable As New SIGHEntidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim lbPrimeraVez As Boolean
Dim ml_IdPaciente As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdFuenteFinanciamiento  As Long
Dim ml_IdServicioPaciente As Long
Dim oDOPaciente As New doPaciente
Dim oDoImagMovimiento As New DoImagMovimiento
Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
Dim rsProductos As Recordset
Dim ml_PuntoCarga As Long
Const ml_IdTipoFinanciamiento As Long = 1
Const lcConstanteMovimientoEntrada As String = "E"
Dim ml_IdTipoVentaSeleccionada As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
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


Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    Me.txtNmovimiento = oDoImagMovimiento.IdMovimiento
                    MsgBox "Se gregó correctamente el Movimiento N° " & oDoImagMovimiento.IdMovimiento, vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo agregar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Se Modificó correctamente el Movimiento N° " & oDoImagMovimiento.IdMovimiento, vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
            If MsgBox("¿Realmente desea Anular?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                 Exit Sub
            End If
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                    MsgBox "Los datos se Anularon correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo anular los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
        
End Sub


Function ValidarDatosObligatorios() As Boolean
    
    ValidarDatosObligatorios = False
    ms_MensajeError = ""
    If txtNdocumento.Text = "" Then
        ms_MensajeError = ms_MensajeError & "Tiene que registrar el Documento" & Chr(13)
    End If
    If cmbResponsable.Text = "" Then
        ms_MensajeError = ms_MensajeError & "Tiene que elegir el Responsable que Recepciona" & Chr(13)
    End If
    If cmbIdPuntoDeCarga.Text = "" Then
        ms_MensajeError = ms_MensajeError & "Tiene que elegir el Punto de Carga" & Chr(13)
    End If
    Select Case mi_Opcion
    Case sghAgregar, sghModificar
        Set rsProductos = Me.ucProductos.FacturacionProductos
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                If rsProductos!idProducto = 0 Then
                   rsProductos.Delete
                   rsProductos.Update
                Else
                   If rsProductos!Cantidad <= 0 Then
                      ms_MensajeError = ms_MensajeError & "El producto: " & rsProductos!codigo & " " & Trim(rsProductos!nombreProducto) & "   Tiene problemas con la Cantidad" & Chr(13)
                   End If
                   If rsProductos!PrecioUnitario <= 0 Then
                      ms_MensajeError = ms_MensajeError & "El producto: " & rsProductos!codigo & " " & Trim(rsProductos!nombreProducto) & "   Tiene problemas con el Precio" & Chr(13)
                   End If
                End If
                rsProductos.MoveNext
            Loop
        End If
        If Me.ucProductos.DevuelveTotalPagar <= 0 Then
           ms_MensajeError = ms_MensajeError & "El Importe Total es 0.....verifique" & Chr(13)
        End If
    End Select
    If ms_MensajeError = "" Then
       ValidarDatosObligatorios = True
    Else
       MsgBox ms_MensajeError, vbInformation, Me.Caption
    End If
End Function

Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With oDoImagMovimiento
            .fecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .IdImagEstado = sghEstadoTabla.sghRegistrado    'Registrado
            .IdPuntoCarga = ml_PuntoCarga
            .IdTipoConcepto = sghTipoConceptoImagen.sghImgTCingreso  'Ingresos
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .MovTipo = lcConstanteMovimientoEntrada
        End With
        With oDoImagMovimientoIngresos
            .IdPersonaRecepciona = Val(mo_cmbResponsable.BoundText)
            .NroDocumento = txtNdocumento.Text
            .IdUsuarioAuditoria = ml_idUsuario
        End With
    Case sghModificar
        With oDoImagMovimiento
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoImagMovimientoIngresos
            .IdPersonaRecepciona = Val(mo_cmbResponsable.BoundText)
            .NroDocumento = txtNdocumento.Text
            .IdUsuarioAuditoria = ml_idUsuario
        End With
    Case sghEliminar
        With oDoImagMovimiento
            .IdUsuarioAuditoria = ml_idUsuario
        End With
    End Select
End Sub

Function ValidarReglas() As Boolean
   ValidarReglas = False
    

    
   ValidarReglas = True
End Function
Function AgregarDatos() As Boolean
    AgregarDatos = mo_ReglasImagenes.ImagMovimientoIngresosAgregar(oDoImagMovimiento, oDoImagMovimientoIngresos, rsProductos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_ReglasImagenes.MensajeError
    ml_IdMovimiento = oDoImagMovimiento.IdMovimiento
End Function

Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasImagenes.ImagMovimientoIngresosModificar(oDoImagMovimiento, oDoImagMovimientoIngresos, rsProductos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_ReglasImagenes.MensajeError
End Function

Function EliminarDatos() As Boolean
    EliminarDatos = mo_ReglasImagenes.ImagMovimientoIngresosAnular(oDoImagMovimiento, oDoImagMovimientoIngresos, rsProductos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_ReglasImagenes.MensajeError
End Function





Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub



Private Sub cmbIdPuntoDeCarga_Click()
    ml_PuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
    ucProductos.IdPuntoCarga = ml_PuntoCarga
    '
    mo_cmbResponsable.BoundColumn = "idEmpleado"
    mo_cmbResponsable.ListField = "ApNom"
    Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =" & mo_ReglasFarmacia.EmpleadosDevuelveIdCargoSegunPuntoCarga(ml_PuntoCarga))
End Sub

Private Sub Form_Initialize()
    Set mo_cmbResponsable.MiComboBox = cmbResponsable
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
    
End Sub

Private Sub Form_Load()
    txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL
    txtEstado.Text = "Registrado"
    
    CargaDataCombos
    
    Me.ucProductos.HabilitaIngresoDePrecio = False
    Me.ucProductos.PermiteVerColumnaCantidadFallada = False
    Me.ucProductos.idUsuario = ml_idUsuario
    Me.ucProductos.Inicializar
    Me.ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    Me.ucProductos.TipoProducto = sghServicio
    Me.ucProductos.IdPuntoCarga = ml_PuntoCarga

    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Ingreso"
    Case sghModificar
        Me.Caption = "Modificar Ingreso"
    Case sghConsultar
        Me.Caption = "Consultar Ingreso"
    Case sghEliminar
        Me.Caption = "Eliminar Ingreso"
    End Select
    
    CargarDatosAlFormulario
End Sub

Sub CargarDatosAlFormulario()
 mo_Formulario.HabilitarDeshabilitar Me.txtNmovimiento, False
 mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
 mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False

 Select Case mi_Opcion
     Case sghAgregar
        Me.ucProductos.IdOrden = -999
        Me.ucProductos.CargaProductosPorIdOrden
        Me.ucProductos.AgregaProducto
     Case sghModificar
        CargarDatosALosControles
     Case sghConsultar
        CargarDatosALosControles
     Case sghEliminar
        CargarDatosALosControles
 End Select
End Sub

Sub CargarDatosALosControles()
        mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
        
        Set oDoImagMovimiento = mo_ReglasImagenes.ImagMovimientoSeleccionarPorId(ml_IdMovimiento)
        txtFregistro.Text = Format(oDoImagMovimiento.fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
        txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDeImagen("idImagEstado=" & oDoImagMovimiento.IdImagEstado)
        txtNmovimiento.Text = ml_IdMovimiento
        mo_cmbIdPuntoCarga.BoundText = oDoImagMovimiento.IdPuntoCarga
        '
        Set oDoImagMovimientoIngresos = mo_ReglasImagenes.ImagMovimientoIngresosSeleccionarPorId(ml_IdMovimiento)
        txtNdocumento.Text = oDoImagMovimientoIngresos.NroDocumento
        mo_cmbResponsable.BoundText = oDoImagMovimientoIngresos.IdPersonaRecepciona
        mb_ExistenDatos = True
         
        'Cargar datos de los servicios
        Me.ucProductos.LimpiarGrilla
        Me.ucProductos.IdMovimiento = ml_IdMovimiento
        Me.ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
        Me.ucProductos.CargaProductosPorIdMovimiento
        
        If oDoImagMovimiento.IdImagEstado = 0 Or mi_Opcion = sghConsultar Then
           btnAceptar.Enabled = False
        End If
        
        Select Case mi_Opcion
        Case sghModificar
        Case sghEliminar
        Case sghConsultar
        End Select
   
   
End Sub




Sub CargaDataCombos()

    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCargaSegunFiltro("idUPS=1")
    Dim rsPuntoCargaDondeLabora As Recordset
    Set rsPuntoCargaDondeLabora = mo_reglasComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghImageneología, ml_idUsuario)
    If rsPuntoCargaDondeLabora.RecordCount > 0 Then
       mo_cmbIdPuntoCarga.BoundText = rsPuntoCargaDondeLabora.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
       cmbIdPuntoDeCarga_Click
    End If
    '


End Sub


Private Sub txtNdocumento_KeyDown(KeyCode As Integer, Shift As Integer)
        mo_Teclado.RealizarNavegacion KeyCode, txtNdocumento
End Sub
