VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FarmNotaSalida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "farmNotaSalida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CargaInventarioExcel 
      Enabled         =   0   'False
      Height          =   315
      Left            =   14685
      Picture         =   "farmNotaSalida.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8940
      Width           =   435
   End
   Begin VB.Frame FrmExcel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Medicamentos e Insumos cargados del EXCEL"
      ForeColor       =   &H00000000&
      Height          =   1560
      Left            =   0
      TabIndex        =   25
      Top             =   2640
      Visible         =   0   'False
      Width           =   15030
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         DisabledPicture =   "farmNotaSalida.frx":110C
         DownPicture     =   "farmNotaSalida.frx":15D0
         Height          =   700
         Left            =   13635
         Picture         =   "farmNotaSalida.frx":1ABC
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   195
         Width           =   1365
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         DisabledPicture =   "farmNotaSalida.frx":1FA8
         DownPicture     =   "farmNotaSalida.frx":2408
         Height          =   700
         Left            =   12105
         Picture         =   "farmNotaSalida.frx":287D
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   195
         Width           =   1365
      End
      Begin VB.CheckBox chkTodos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Todos/Ninguno"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   150
         TabIndex        =   26
         Top             =   345
         Width           =   1785
      End
      Begin UltraGrid.SSUltraGrid grdExcel 
         Height          =   3240
         Left            =   75
         TabIndex        =   29
         Top             =   1110
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   5715
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "farmNotaSalida.frx":2CBF
         Caption         =   ".."
      End
      Begin VB.Label lblConsideraciones 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Cada Producto debe tener PRECIOS y TIPO DE SALIDA (Fact-config->cat.BienesInsumos->particular->doble clic"""
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2460
         TabIndex        =   31
         Top             =   405
         Width           =   9420
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   60
      TabIndex        =   19
      Top             =   7770
      Width           =   15120
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "farmNotaSalida.frx":2CFB
         DownPicture     =   "farmNotaSalida.frx":31BF
         Height          =   700
         Left            =   7703
         Picture         =   "farmNotaSalida.frx":36AB
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         Height          =   700
         Left            =   120
         Picture         =   "farmNotaSalida.frx":3B97
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "farmNotaSalida.frx":4070
         DownPicture     =   "farmNotaSalida.frx":44D0
         Height          =   700
         Left            =   6173
         Picture         =   "farmNotaSalida.frx":4945
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraCabecera 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   60
      TabIndex        =   5
      Top             =   30
      Width           =   15105
      Begin VB.CommandButton cmdStockMinimo 
         Height          =   330
         Left            =   6855
         Picture         =   "farmNotaSalida.frx":4DBA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Lista de ITEMS que están debajo de su STOCK MINIMO"
         Top             =   675
         Width           =   300
      End
      Begin VB.CommandButton CargaPedidosExcel 
         Caption         =   "Requerimiento"
         Height          =   700
         Left            =   13515
         Picture         =   "farmNotaSalida.frx":5344
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Carga REQUERIMIENTOS desde  C:\pedidos.XLS (a=código, b=producto, c=cantidad pedida)   <<empieza en FILA=2>> <<Hoja1>>"
         Top             =   1695
         Width           =   1365
      End
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
         Left            =   1500
         TabIndex        =   23
         Top             =   690
         Width           =   5340
      End
      Begin VB.TextBox txtNotaSalida 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   9
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   13470
         MaxLength       =   30
         TabIndex        =   8
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox txtNdocum 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8910
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1560
         Width           =   1635
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   315
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1980
         Width           =   5325
      End
      Begin VB.ComboBox cmbConcepto 
         Height          =   330
         Left            =   1500
         TabIndex        =   0
         Top             =   1140
         Width           =   5340
      End
      Begin VB.ComboBox cmbTipoDocum 
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
         Left            =   1500
         TabIndex        =   7
         Top             =   1560
         Width           =   5340
      End
      Begin VB.ComboBox cmbAlmDestino 
         Height          =   330
         Left            =   8910
         TabIndex        =   1
         Top             =   1140
         Width           =   5970
      End
      Begin VB.TextBox txtHoraRegistro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10380
         MaxLength       =   30
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   8910
         TabIndex        =   10
         Top             =   300
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén origen"
         Height          =   210
         Left            =   150
         TabIndex        =   24
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   210
         Left            =   150
         TabIndex        =   18
         Top             =   1170
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
         Height          =   210
         Left            =   8040
         TabIndex        =   17
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Nota Salida"
         Height          =   210
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   210
         Left            =   12840
         TabIndex        =   15
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   210
         Left            =   8235
         TabIndex        =   14
         Top             =   1185
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Docum"
         Height          =   210
         Left            =   150
         TabIndex        =   13
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "N° Docum"
         Height          =   210
         Left            =   8010
         TabIndex        =   12
         Top             =   1590
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   2010
         Width           =   1170
      End
   End
   Begin SighFarmacia.ucNotaSalida grdProductos 
      Height          =   5025
      Left            =   60
      TabIndex        =   4
      Top             =   2670
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   8864
   End
   Begin UltraGrid.SSUltraGrid grdHistorico 
      Height          =   1470
      Left            =   60
      TabIndex        =   33
      Top             =   8910
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   2593
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "farmNotaSalida.frx":5786
      Caption         =   "Consumo histórico del PACIENTE x CUENTA"
   End
End
Attribute VB_Name = "FarmNotaSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Notas de Salida
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_movNumero As String
Dim mo_cmbConceptos As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenDestino As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoDocum As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim oRsConceptos As New ADODB.Recordset
Dim oRsAlmacenOrigen As New ADODB.Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mRs_Productos As New ADODB.Recordset
Dim mo_farmMovimiento As New sighComun.DoFarmMovimiento
Dim mo_farmMovimiento1 As New DoFarmMovimiento
Const lcConstanteMovimientoSalida As String = "S"
Const lcConstanteMovimientoEntrada As String = "E"
Dim lnTotalDocumento As Double
Dim ms_MensajeError As String
Dim mo_farmMovimientoNotaIngreso As New sighComun.DOfarmMovimientoNotaIngreso
Dim oDoProveedores As New DoProveedores
Dim lcTipoLocalesAlmOrigen As String
Dim lbDocumentoEsAutomatico As Boolean
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcTipoLocalesAlmDestino As String
Dim mo_lbElEstablecimentoEsCS As Boolean
Dim ml_idUsuarioCreo As Long
Dim lcCodigoSismedFarmDestino As String
Dim lbLaFarmaciaOrigenEsUnidosis As Boolean
Dim lbLaFarmaciaDestinoEsUnidosis As Boolean
Dim oRsItemsUnidosis As New Recordset

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property







Private Sub ImprimeDocumento()
    Dim oRptClase As New rCrystal
    Dim oDOfarmAlmacen As New DoFarmAlmacen
    Set oDOfarmAlmacen = mo_ReglasFarmacia.FarmAlmacenSeleccionarPorId(Val(mo_cmbAlmacenOrigen.BoundText))
    oRptClase.MovTipo = "S"
    oRptClase.Documento = txtNotaSalida.Text
    oRptClase.TextoDelFiltro = "NOTA DE SALIDA"
    oRptClase.Almacen = cmbAlmDestino.Text
    oRptClase.AlmacenO = "(" & oDOfarmAlmacen.CodigoSismed & ")" & cmbAlmOrigen.Text
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = Trim(cmbTipoDocum.Text) & " - " & txtNdocum.Text
    oRptClase.Importe = lnTotalDocumento
    oRptClase.TipoReporte = "NiNs"
    oRptClase.Observaciones = Trim(Me.txtObservaciones.Text) & "  (" & Label1.Caption & ":  " & cmbConcepto.Text & ")"   'debb-07/10/2016
    oRptClase.EsUnaDonacion = IIf(mo_cmbConceptos.BoundText = "3", True, False)
    'If Trim(cmbTipoDocum.Text) <> "" Then
    '    oRptClase.Proveedor = Trim(cmbTipoDocum.Text) & "/" & Trim(txtNdocum.Text)
    'End If
    oRptClase.idUsuario = ml_idUsuarioCreo
    oRptClase.Show vbModal
    Set oRptClase = Nothing
    Set oDOfarmAlmacen = Nothing
End Sub

Private Sub btnImprimir_Click()
   ImprimeDocumento
End Sub



Private Sub CargaInventarioExcel_Click()
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    mo_ReglasReportes.ExportarRecordSetAexcelFast Me.grdHistorico.DataSource, Me.grdHistorico.Caption, "", "", Me.hwnd
    Set mo_ReglasReportes = Nothing
End Sub

Private Sub CargaPedidosExcel_Click()
    On Error GoTo ErrCargaExc
    FrmExcel.Visible = False
    If mi_Opcion = sghAgregar And mo_lnIdTablaLISTBARITEMS = 1305 Then
        If cmbAlmDestino.Text = "" Then
           MsgBox "Debe elegir FARMACIA DESTINO", vbInformation, ""
           Exit Sub
        End If
        Dim EXL As Excel.Application
        Set EXL = New Excel.Application
        Dim W As Excel.Workbook
        Set W = EXL.Workbooks.Open("c:\PEDIDOS.xls")
        Dim s As Excel.Worksheet
        Set s = W.Sheets("Hoja1")
        Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, lcMensaje As String
        Dim lcCodigo As String, lcFvencimiento As String, lnSaldo As Long, lcRegSanitario As String, lnPrecioUnitario As Double
        Dim oRsTmp As New Recordset, rs As New Recordset
        Dim oConexion As New Connection
        Dim lnIdProducto As Long, lcLote As String, lnIdTipoSalidaBienInsumo As Long, lcNombreProducto As String
        Dim lbEsNuevo As Boolean, lcDescripcion As String, lnCantidad As Long, lcFarmacia As String
        Dim lcError As String, lnNumero As Integer, lcProducto1 As String
        Me.MousePointer = 11
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        With rs
                  .Fields.Append "Nro", adInteger
                  .Fields.Append "idProducto", adInteger
                  .Fields.Append "Codigo", adVarChar, 20
                  .Fields.Append "Producto", adChar, 300
                  .Fields.Append "Cantidad", adInteger
                  .Fields.Append "Error", adVarChar, 150
                  .Fields.Append "Saldo", adInteger
                  .Fields.Append "seleccionar", adBoolean
                  .Fields.Append "Producto1", adChar, 300
                  .Fields.Append "idTipoSalidaBienInsumo", adInteger
                  .CursorType = adOpenKeyset
                  .LockType = adLockOptimistic
                  .Open
        End With
        Set grdExcel.DataSource = rs
        mo_Apariencia.ConfigurarFilasBiColores Me.grdExcel, SIGHEntidades.GrillaConFilasBicolor
        lnFila = 2
        lnFilaFinal = 10000
        lcMensaje = ""
        lnNumero = 1
        For lnFor = lnFila To lnFilaFinal
            lcRango = "A" + Trim(Str(lnFor))
            lcCodigo = Right("00000" & Trim(s.Range(lcRango).Value), 5)
            If Val(lcCodigo) = 0 Then
               Exit For
            End If
            lcRango = "B" + Trim(Str(lnFor))
            lcDescripcion = Trim(s.Range(lcRango).Value)
            lcRango = "C" + Trim(Str(lnFor))
            lnCantidad = Val(Trim(s.Range(lcRango).Value))
            lcRango = "D" + Trim(Str(lnFor))
            lcFarmacia = Trim(s.Range(lcRango).Value)
            lcError = "": lnIdProducto = 0: lnSaldo = 0: lcProducto1 = "": lnIdTipoSalidaBienInsumo = 0
           ' If lcCodigoSismedFarmDestino = lcFarmacia Then
                Set oRsTmp = mo_ReglasFacturacion.FactCatalogoBienesInsumosSeleccionarXcodigo(lcCodigo, oConexion)
                If oRsTmp.RecordCount > 0 Then
                    lcProducto1 = oRsTmp!Nombre
                    lnIdTipoSalidaBienInsumo = oRsTmp!idTipoSalidaBienInsumo
                    lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(oRsTmp!idProducto, sghPrecioCompra)
                    If lnPrecioUnitario = 0 Then
                       lcError = "NO tiene PRECIO DE DISTRIBUCION"
                    Else
                       lnIdProducto = oRsTmp!idProducto
                       oRsTmp.Close
                       Set oRsTmp = mo_ReglasFarmacia.farmSaldoSoloAlmacenSismed(lnIdProducto, oConexion)
                       oRsTmp.Filter = "idAlmacen=" & mo_cmbAlmacenOrigen.BoundText & " and idTipoSalidaBienInsumo=" & _
                                       Trim(Str(lnIdTipoSalidaBienInsumo))
                       If oRsTmp.RecordCount > 0 Then
                          lnSaldo = oRsTmp!Cantidad
                          If lnSaldo < lnCantidad Then
                             lcError = "NO hay SALDO SUFICIENTE"
                          End If
                       End If
                    End If
                Else
                    lcError = "No existe el CODIGO"
                End If
                oRsTmp.Close
                If lcError = "" Then
                   If lnIdTipoSalidaBienInsumo <> 1 And lnIdTipoSalidaBienInsumo <> 2 Then
                      'lcError = "TIPO SALIDA solo puede ser VENTA o INTERSAN"
                      If MsgBox("El tipo de salida actual es VENTA/INTERV.SANITARIAS, ¿despachará como VENTAS ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                         lnIdTipoSalidaBienInsumo = 2
                      Else
                         lnIdTipoSalidaBienInsumo = 1
                      End If
                   End If
                End If
                rs.AddNew
                rs!Nro = lnNumero
                rs!idProducto = lnIdProducto
                rs!codigo = lcCodigo
                rs!Producto = lcDescripcion
                rs!Cantidad = lnCantidad
                rs!Error = lcError
                rs!saldo = lnSaldo
                rs!seleccionar = IIf(lnSaldo >= lnCantidad, True, False)
                rs!producto1 = lcProducto1
                rs!idTipoSalidaBienInsumo = lnIdTipoSalidaBienInsumo
                rs.Update
                lnNumero = lnNumero + 1
            ' End If
        Next
        Set s = Nothing
'        W.Save
        W.Close
        Set W = Nothing
        Set EXL = Nothing
        If rs.RecordCount = 0 Then
           MsgBox "El EXCEL está correcto, pero no existen PRODUCTOS para la FARMACIA DESTINO", vbInformation, ""
        Else
           FrmExcel.Visible = True
           'FrmExcel.BackColor = vbBlue
           FrmExcel.Top = grdProductos.Top
           FrmExcel.Left = grdProductos.Left
           FrmExcel.Width = grdProductos.Width
           FrmExcel.Height = grdProductos.Height + 1000
           grdExcel.Left = FrmExcel.Left + 100
           grdExcel.Width = FrmExcel.Width - 300
           grdExcel.Height = FrmExcel.Height - 1300
        End If
        Me.MousePointer = 1
    End If
ErrCargaExc:
        If Err.Number > 0 Then
           MsgBox CargaPedidosExcel.ToolTipText & Chr(13) & Chr(13) & Err.Description, vbInformation, ""
        End If
        Set s = Nothing
        Set W = Nothing
        Set EXL = Nothing
        Set oRsTmp = Nothing
        Set rs = Nothing
        Set oConexion = Nothing
        Me.MousePointer = 1
End Sub

Private Sub chkTodos_Click()
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = grdExcel.DataSource
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
            If chkTodos.Value = 1 And oRsTmp1!saldo >= oRsTmp1!Cantidad Then
               oRsTmp1!seleccionar = True
            Else
               oRsTmp1!seleccionar = False
            End If
            oRsTmp1.MoveNext
       Loop
    End If
    Set grdExcel.DataSource = oRsTmp1
    Set oRsTmp1 = Nothing
End Sub

Private Sub cmbAlmDestino_Click()
    lbLaFarmaciaDestinoEsUnidosis = False
    lcCodigoSismedFarmDestino = ""
    Dim oRsTmp As New Recordset
    '** solo en caso de donaciones
    If mo_cmbConceptos.BoundText = "3" Then
        oRsConceptos.MoveFirst
        oRsConceptos.Find "idTipoConcepto=" & mo_cmbConceptos.BoundText
        mo_cmbTipoDocum.BoundText = oRsConceptos.Fields!DocumentoId
        lcTipoLocalesAlmDestino = ""
        Set oRsTmp = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idAlmacen=" & mo_cmbAlmacenDestino.BoundText)
        If oRsTmp.RecordCount > 0 Then
           lcTipoLocalesAlmDestino = oRsTmp.Fields!idTipoLocales
           If oRsTmp.Fields!idTipoLocales = "F" Then
              mo_cmbTipoDocum.BoundText = "15" 'ppa
              Me.txtNdocum.Text = ""
              lcCodigoSismedFarmDestino = oRsTmp!CodigoSismed
           End If
        End If
        oRsTmp.Close
    Else
        Set oRsTmp = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idAlmacen=" & mo_cmbAlmacenDestino.BoundText)
        If oRsTmp.RecordCount > 0 Then
           If oRsTmp.Fields!idTipoLocales = "F" And Not IsNull(oRsTmp!CodigoSismed) Then
              lcCodigoSismedFarmDestino = oRsTmp!CodigoSismed
           End If
        End If
        oRsTmp.Close
        lbLaFarmaciaDestinoEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(Val(mo_cmbAlmacenDestino.BoundText))
    End If
    Set oRsTmp = Nothing
End Sub

Private Sub cmbAlmDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmDestino

End Sub

Private Sub cmbAlmOrigen_Click()
    oRsAlmacenOrigen.MoveFirst
    oRsAlmacenOrigen.Find "idAlmacen=" & mo_cmbAlmacenOrigen.BoundText
    Set oRsConceptos = mo_ReglasFarmacia.FarmTipoConceptosDevuelveParaRegistroDeNiNs(oRsAlmacenOrigen.Fields!idTipoLocales, lcConstanteMovimientoSalida, oRsAlmacenOrigen.Fields!idTipoSuministro)
    mo_cmbConceptos.BoundColumn = "IdTipoConcepto"
    mo_cmbConceptos.ListField = "Concepto"
    Set mo_cmbConceptos.RowSource = mo_ReglasFarmacia.FarmTipoConceptosDevuelveParaRegistroDeNiNs(oRsAlmacenOrigen.Fields!idTipoLocales, lcConstanteMovimientoSalida, oRsAlmacenOrigen.Fields!idTipoSuministro)
    grdProductos.IdAlmacen = oRsAlmacenOrigen.Fields!IdAlmacen
    lcTipoLocalesAlmOrigen = oRsAlmacenOrigen.Fields!idTipoLocales
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    lbLaFarmaciaOrigenEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(Val(mo_cmbAlmacenOrigen.BoundText))
End Sub


Private Sub cmbAlmOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmOrigen

End Sub

Private Sub cmbConcepto_Click()
    If Val(mo_cmbConceptos.BoundText) = 0 Then
       Exit Sub
    End If
    oRsConceptos.MoveFirst
    oRsConceptos.Find "idTipoConcepto=" & mo_cmbConceptos.BoundText
    mo_cmbTipoDocum.BoundText = oRsConceptos.Fields!DocumentoId
    mo_cmbAlmacenDestino.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenDestino.ListField = "Descripcion"
    If mo_lbElEstablecimentoEsCS = True Then
       Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro(oRsConceptos.Fields!NsFiltroAlmacenDestinoCS & " and idEstado=1")
    Else
       Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro(oRsConceptos.Fields!NsFiltroAlmacenDestino & " and idEstado=1")
    End If
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If cmbAlmDestino.ListCount = 1 Then
       cmbAlmDestino.ListIndex = 0
    End If
    '
    grdProductos.MuestraLoteParaDespachoNS = IIf(oRsConceptos.Fields!MuestraLoteParaDespachoNS = "S", True, False)
    grdProductos.TipoPrecioParaNiNs = oRsConceptos.Fields!TipoPrecioParaNiNs
    '
    lbDocumentoEsAutomatico = IIf(oRsConceptos.Fields!DocumentoEsAutomatico = "S", True, False)
    If lbDocumentoEsAutomatico = True Then
       'SCCQ 09/10/2020 Cambio28 Inicio
       'txtNdocum.Text = Val(oRsConceptos.Fields!DocumentoUltimoNumero) + 1
       'SCCQ 09/10/2020 Cambio28 Fin
    Else
       txtNdocum.Text = ""
    End If
    
    '
    mo_Formulario.HabilitarDeshabilitar cmbAlmDestino, True
    If lbLaFarmaciaOrigenEsUnidosis = True And InStr("/4/5/6/7/", mo_cmbConceptos.BoundText) > 0 Then
       mo_Formulario.HabilitarDeshabilitar cmbAlmDestino, False
       cmbAlmDestino.Text = ""
    End If
End Sub

Private Sub cmbConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbConcepto

End Sub

Private Sub cmdAgregar_Click()
    
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = grdExcel.DataSource
    If oRsTmp1.RecordCount > 0 Then
       grdProductos.CargaProductosExcel oRsTmp1
    End If
    FrmExcel.Visible = False
    Set oRsTmp1 = Nothing
End Sub

Private Sub cmdSalir_Click()
    FrmExcel.Visible = False
End Sub

Private Sub cmdStockMinimo_Click()
    CargaItemsDebajoDeStockMinimo
End Sub

Private Sub Form_Activate()
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(mo_farmMovimiento.IdAlmacenOrigen) = True Then
        btnCancelar_Click
        Exit Sub
   End If

End Sub

Private Sub Form_Initialize()
    Set mo_cmbConceptos.MiComboBox = cmbConcepto
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
    Set mo_cmbAlmacenDestino.MiComboBox = cmbAlmDestino
    Set mo_cmbTipoDocum.MiComboBox = cmbTipoDocum

End Sub

Private Sub Form_Load()
    SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
    mo_lbElEstablecimentoEsCS = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
    ConfigurarGrdProductos
    CargarComboBoxes
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Nota Salida"
    Case sghModificar
        Me.Caption = "Modificar Nota Salida"
    Case sghConsultar
        Me.Caption = "Consultar Nota Salida"
        btnImprimir.Visible = True
    Case sghEliminar
        Me.Caption = "Anular Nota Salida"
    End Select
    CargarDatosAlFormulario
    CargaItemsDebajoDeStockMinimo
    If mi_Opcion = sghAgregar And mo_lnIdTablaLISTBARITEMS = 1305 Then   'ns del ALMACEN ESPECIALIZADO
       CargaPedidosExcel.Visible = True
    Else
       CargaPedidosExcel.Visible = False
    End If
End Sub
Sub ConfigurarGrdProductos()
    grdProductos.movNumero = ml_movNumero
    grdProductos.IdAlmacen = 0
    grdProductos.FechaMinimaDespacho = CDate(lcBuscaParametro.RetornaFechaServidorSQL) + Val(lcBuscaParametro.SeleccionaFilaParametro(220))
    grdProductos.inicializar
End Sub


Sub CargarComboBoxes()
    Set oRsItemsUnidosis = mo_ReglasFarmacia.farmUnidosisSeleccionarTodos
    
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    '
    'Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    If mo_lnIdTablaLISTBARITEMS <> 1305 Then
       Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1")
       Label4.Caption = "Farmacia origen"
    Else
       Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and idEstado=1")
       Label4.Caption = "Almacén origen"
    End If
    '
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    'Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    If mo_lnIdTablaLISTBARITEMS <> 1305 Then
       Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1")
    Else
       Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and idEstado=1")
    End If
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacenOrigen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
       oRsAlmacenOrigen.MoveFirst
       oRsAlmacenOrigen.Find "idAlmacen=" & mo_cmbAlmacenOrigen.BoundText
       lcTipoLocalesAlmOrigen = oRsAlmacenOrigen.Fields!idTipoLocales
       lbLaFarmaciaOrigenEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(Val(mo_cmbAlmacenOrigen.BoundText))
    End If
   '
    mo_cmbTipoDocum.BoundColumn = "idTipoDocumento"
    mo_cmbTipoDocum.ListField = "Nombre"
    Set mo_cmbTipoDocum.RowSource = mo_ReglasFarmacia.FarmTipoDocumentosDevuelveTodos
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError
    End If
End Sub
Sub CargarDatosAlFormulario()
    'SCCQ 14/10/2020 Cambio28 Inicio
    mo_Formulario.HabilitarDeshabilitar txtNdocum, False
    'SCCQ 14/10/2020 Cambio28 Fin
    mo_Formulario.HabilitarDeshabilitar Me.txtNotaSalida, False
    mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraRegistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbTipoDocum, False
  

     Select Case mi_Opcion
     Case sghAgregar
        txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL      'Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
        txtHoraRegistro.Text = lcBuscaParametro.RetornaHoraServidorSQL
        grdProductos.movNumero = ""
        grdProductos.LimpiarGrilla
        grdProductos.CargaProductosPorMovNumero
        grdProductos.AgregaRegistro
        
     Case sghModificar
        DeshabilitaCabecera
        CargarDatosALosControles
     Case sghConsultar
        DeshabilitaCabecera
        CargarDatosALosControles
        btnAceptar.Enabled = False
     Case sghEliminar
        DeshabilitaCabecera
        CargarDatosALosControles
 End Select
End Sub

Sub CargarDatosALosControles()
 '**************Datos de la tabla FarmMovimiento *****************
   mo_farmMovimiento.movNumero = ml_movNumero
   mo_farmMovimiento.MovTipo = lcConstanteMovimientoSalida
   If Not mo_ReglasFarmacia.FarmMovimientoSeleccionarPorId(mo_farmMovimiento) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
   End If
   txtNotaSalida.Text = ml_movNumero
   mo_cmbAlmacenOrigen.BoundText = mo_farmMovimiento.IdAlmacenOrigen
   mo_cmbConceptos.BoundText = mo_farmMovimiento.idTipoConcepto
   mo_cmbAlmacenDestino.BoundText = mo_farmMovimiento.IdAlmacenDestino
   mo_cmbTipoDocum.BoundText = mo_farmMovimiento.DocumentoIdtipo
   txtNdocum.Text = mo_farmMovimiento.DocumentoNumero
   txtObservaciones.Text = mo_farmMovimiento.Observaciones
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDelMovimiento("idEstadoMovimiento=" & mo_farmMovimiento.idEstadoMovimiento)
   txtFregistro.Text = Format(mo_farmMovimiento.fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   txtHoraRegistro.Text = Format(mo_farmMovimiento.fechaCreacion, SIGHEntidades.DevuelveHoraSoloFormato_HM)
   ml_idUsuarioCreo = mo_farmMovimiento.idUsuario
   lbLaFarmaciaOrigenEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(mo_cmbAlmacenOrigen.BoundText)
   '**************Datos de la tabla FarmMovimientoDetalle *****************
   grdProductos.movNumero = ml_movNumero
   grdProductos.CargaProductosPorMovNumero
   grdProductos.RefrescarDatos
   lnTotalDocumento = grdProductos.DevuelveTotal
   If mo_farmMovimiento.idEstadoMovimiento = 0 Then
      btnAceptar.Enabled = False
   End If
   'PAQUETES
   If mo_ReglasFarmacia.LaNIoNSesUnARMADO_PAQUETE(mo_farmMovimiento.IdAlmacenOrigen, mo_farmMovimiento.idTipoConcepto, _
                                                  mo_farmMovimiento.DocumentoNumero, False) = True Then
      MsgBox "No puede MODIFICAR/ELIMINAR la Nota de Ingreso, debe de usar la opción ARMADO DE PAQUETES", vbInformation, Me.Caption
      btnAceptar.Enabled = False
   End If
   '******permiso a Modificar documento con Fecha Anterior a la actual
   Dim mo_PermisosFacturacion As New PermisosFacturacion
   Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
   Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
   If mo_PermisosFacturacion.ActualizaFechaDocumentoES = False And mi_Opcion <> sghConsultar Then
      If CDate(lcBuscaParametro.RetornaFechaServidorSQL) <> CDate(txtFregistro.Text) Then
         MsgBox "No tiene ACCESO a Modificar/Anular una NS" & Chr(13) & " de una Fecha Registro diferente a la actual", vbExclamation, Me.Caption
         btnAceptar.Enabled = False
      End If
   End If
   Set mo_PermisosFacturacion = Nothing
   Set mo_ReglasSeguridad = Nothing
   'SCCQ 29/10/2020 Cambio32 Inicio
     '******Modificar documento con Fecha Anterior a la actual,
   '******siempre y cuando no hubieron SALIDAS
   Dim oRsTmp As New ADODB.Recordset
   Set mRs_Productos = grdProductos.DevuelveProductos
   If mRs_Productos.RecordCount > 0 Then
      mRs_Productos.MoveFirst
      Do While Not mRs_Productos.EOF
         Set oRsTmp = mo_ReglasFarmacia.farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote(mo_farmMovimiento.IdAlmacenDestino, mRs_Productos.Fields!idProducto, mRs_Productos.Fields!Lote, mRs_Productos.Fields!FechaVencimiento)
         If oRsTmp.RecordCount > 0 Then
            If oRsTmp.Fields!fechaCreacion >= CDate(txtFregistro.Text & " " & txtHoraRegistro.Text) Then
                MsgBox "No podrá Modificar/Anular la NS porque el destino ya despachó el producto: " & Chr(13) & Trim(mRs_Productos.Fields!codigo) & " - " & Trim(mRs_Productos.Fields!nombreProducto) & "   NS: " & oRsTmp.Fields!movNumero, vbExclamation, Me.Caption
                btnAceptar.Enabled = False
                Exit Do
            End If
         End If
         mRs_Productos.MoveNext
      Loop
   End If
   'SCCQ 29/10/2020 Cambio32 Fin
   'unidosis
   lbLaFarmaciaDestinoEsUnidosis = False
   If Trim(mo_farmMovimiento.Observaciones) <> "" Then
        mo_farmMovimiento1.movNumero = Trim(mo_farmMovimiento.Observaciones)
        mo_farmMovimiento1.MovTipo = lcConstanteMovimientoEntrada
        mo_farmMovimiento1.IdUsuarioAuditoria = mo_farmMovimiento.IdUsuarioAuditoria
        If mo_ReglasFarmacia.FarmMovimientoSeleccionarPorId(mo_farmMovimiento1) Then
            lbLaFarmaciaDestinoEsUnidosis = True
        End If
   End If
   '
   
End Sub

Sub DeshabilitaCabecera()
    mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbAlmDestino, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbTipoDocum, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbConcepto, False
  
End Sub
Private Sub btnCancelar_Click()
   If SIGHEntidades.ParaAuditoria = "" Then
      Me.Visible = False
      LimpiarVariablesDeMemoria
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      Me.Visible = False
      LimpiarVariablesDeMemoria
      SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
   End If
End Sub
Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(Val(mo_cmbAlmacenOrigen.BoundText)) = True Then
      btnCancelar_Click
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
        'SCCQ 09/10/2020 Cambio28 Inicio
        'Antes:  If ValidarDatosObligatorios() Then
        If ValidarDatosObligatorios("A") Then
        'SCCQ 09/10/2020 Cambio28 Fin
            CargaDatosAlObjetosDeDatos
            If AgregarDatos() Then
            'SCCQ 09/10/2020 Cambio28 Inicio
            'Antes: If MsgBox("Se agregó correctamente la Nota de Salida N° " + txtNotaSalida.Text + Chr(13) + Chr(13) + "Desea Imprimir el Documento ?", vbQuestion + vbYesNo, "") = vbYes Then
                If MsgBox("Se agregó correctamente la NOTA DE SALIDA N° " + txtNotaSalida.Text + Chr(13) + "Con " + Trim(cmbTipoDocum.Text) + " N° " + mo_farmMovimiento.DocumentoNumero + Chr(13) + Chr(13) + " Desea Imprimir el Documento ?", vbQuestion + vbYesNo, "") = vbYes Then
            'SCCQ 09/10/2020 Cambio28 Fin
                   ml_idUsuarioCreo = ml_idUsuario
                   ImprimeDocumento
                End If
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo agregar los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
   Case sghModificar
       'SCCQ 09/10/2020 Cambio28 Inicio
        'Antes:  If ValidarDatosObligatorios() Then
        If ValidarDatosObligatorios("M") Then
        'SCCQ 09/10/2020 Cambio28 Fin
            CargaDatosAlObjetosDeDatos
            If ModificarDatos() Then
                If MsgBox("Se Modificó correctamente la Nota de Salida N° " + txtNotaSalida.Text + Chr(13) + Chr(13) + "Desea Imprimir el Documento ?", vbQuestion + vbYesNo, "") = vbYes Then
                   ml_idUsuarioCreo = ml_idUsuario
                   ImprimeDocumento
                End If
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
   Case sghEliminar
        If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
            CargaDatosAlObjetosDeDatos
            If AnularNS() Then
                MsgBox " Se anuló la Nota de Salida N° " + txtNotaSalida.Text, vbInformation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
   End Select
End Sub
'SCCQ 09/10/2020 Cambio28 Inicio
'Antes:  Function ValidarDatosObligatorios() As Boolean
Function ValidarDatosObligatorios(modo As String) As Boolean
'SCCQ 09/10/2020 Cambio28 Fin
   ValidarDatosObligatorios = False
   ms_MensajeError = ""
   If cmbAlmOrigen.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
   ElseIf cmbConcepto.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Concepto" + Chr(13)
       cmbConcepto.SetFocus
   ElseIf Val(mo_cmbAlmacenDestino.BoundText) = 0 And Val(mo_cmbConceptos.BoundText) <> 20 Then   'ajuste inventario no necesita almacen destino
       ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Destino" + Chr(13)
       cmbAlmDestino.SetFocus
   ElseIf mo_cmbAlmacenOrigen.BoundText = mo_cmbAlmacenDestino.BoundText Then
       ms_MensajeError = ms_MensajeError + "El Almacén Origen y Destino deben ser DIFERENTES" + Chr(13)
   ElseIf cmbTipoDocum.Text <> "" Then
   'SCCQ 09/10/2020 Cambio28 Inicio
        If modo = "M" Then 'Modifica
   'SCCQ 09/10/2020 Cambio28 Fin
         If txtNdocum.Text = "" Then
          ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° Documento" + Chr(13)
          'txtNdocum.SetFocus
         End If
    'SCCQ 09/10/2020 Cambio28 Inicio
        End If
   'SCCQ 09/10/2020 Cambio28 Fin
      
   End If
'SCCQ 08/10/2020 Cambio28 Inicio
'   If mi_Opcion = sghAgregar And txtNdocum.Text <> "" Then
'      Dim oRsTmp As New ADODB.Recordset
'      Set oRsTmp = mo_ReglasFarmacia.farmMovimientoSeleccionarPorTipoYnumeroDocumento(txtNdocum.Text, Val(mo_cmbTipoDocum.BoundText))
'      oRsTmp.Filter = "idEstadoMovimiento=1"
'      If oRsTmp.RecordCount > 0 Then
'         ms_MensajeError = ms_MensajeError + "El Número de Documento: " & txtNdocum.Text & "   EXISTE en NS: " & Trim(oRsTmp.Fields!movNumero) & "     Fecha: " & oRsTmp.Fields!fechaCreacion & Chr(13)
'      End If
'      oRsTmp.Close
'      Set oRsTmp = Nothing
'   End If
'SCCQ 08/10/2020 Cambio28 Fin
   lnTotalDocumento = grdProductos.DevuelveTotal
   Set mRs_Productos = grdProductos.DevuelveProductos
   If mRs_Productos.RecordCount = 0 Then
       ms_MensajeError = ms_MensajeError + "Por favor Ingrese Productos" + Chr(13)
   Else
        Dim LdFechaMinimaDespacho As Date
        If mo_cmbConceptos.BoundText = "5" Then
           LdFechaMinimaDespacho = Date - 1000  'Devolucion por Vencimiento
        Else
           LdFechaMinimaDespacho = CDate(txtFregistro.Text) + Val(lcBuscaParametro.SeleccionaFilaParametro(220))
        End If
        mRs_Productos.MoveFirst
        Do While Not mRs_Productos.EOF
           If Trim(mRs_Productos.Fields!codigo) = "" Or Trim(mRs_Productos.Fields!nombreProducto) = "" Then
              mRs_Productos.Delete
              mRs_Productos.Update
           ElseIf mRs_Productos.Fields!Cantidad <= 0 Or mRs_Productos!Cantidad > mRs_Productos!saldo Then
              ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  Tiene problemas" + Chr(13)
           ElseIf mRs_Productos!Precio <= 0 Then
               ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  Tiene problemas con el Precio" + Chr(13)
           ElseIf mRs_Productos!FechaVencimiento < LdFechaMinimaDespacho Then
               ms_MensajeError = ms_MensajeError + "La F.Vencimiento mínima de despacho es: " & LdFechaMinimaDespacho & " para " & Trim(mRs_Productos.Fields!codigo) & " - " & Trim(mRs_Productos.Fields!nombreProducto) & Chr(13)
           End If
           mRs_Productos.MoveNext
        Loop
   End If
   'Es un despacho hacia la FARMACIA UNIDOSIS  debb-28/06/2019
   If lbLaFarmaciaDestinoEsUnidosis = True Then
        Dim lcCodigoConPunto As String
        Dim rs As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        If oRsItemsUnidosis.RecordCount = 0 Then
           ms_MensajeError = ms_MensajeError + "No hay ITEMS en la FARMACIA UNIDOSIS" + Chr(13)
        Else
           mRs_Productos.MoveFirst
           Do While Not mRs_Productos.EOF
              oRsItemsUnidosis.MoveFirst
              oRsItemsUnidosis.Find "codigo='" & mRs_Productos!codigo & "'"
              If oRsItemsUnidosis.EOF Then
                 ms_MensajeError = ms_MensajeError + "El ITEM " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  no pertenece a FARMACIA UNIDOSIS" + Chr(13)
              Else
                 lcCodigoConPunto = Trim(mRs_Productos!codigo) & SIGHEntidades.Pto
                 Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcCodigoConPunto, 1, oConexion)
                 If rs.RecordCount = 0 Then
                    ms_MensajeError = ms_MensajeError + "El ITEM " + lcCodigoConPunto + " - " + Trim(oRsItemsUnidosis!Descripcion) + "  no tiene PRECIO" + Chr(13)
                 End If
                 rs.Close
              End If
              mRs_Productos.MoveNext
           Loop
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set rs = Nothing
   End If
   '
   
   If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With mo_farmMovimiento
            .DocumentoIdtipo = Val(mo_cmbTipoDocum.BoundText)                   '10
            .DocumentoNumero = txtNdocum.Text
            .fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL       'igual
            .IdAlmacenDestino = Val(mo_cmbAlmacenDestino.BoundText)             '0
            .IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)               '8
            .idEstadoMovimiento = sghEstadoTabla.sghRegistrado                  'igual
            .idTipoConcepto = Val(mo_cmbConceptos.BoundText)                    '20
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .MovTipo = lcConstanteMovimientoSalida
            .Observaciones = txtObservaciones.Text
            .Total = lnTotalDocumento
            
        End With
   Case sghModificar
        With mo_farmMovimiento
            .DocumentoNumero = txtNdocum.Text
            .Observaciones = txtObservaciones.Text
            .IdUsuarioAuditoria = ml_idUsuario
            .Total = lnTotalDocumento
            '.FechaCreacion = txtFregistro.Text
        End With
   Case sghEliminar
        With mo_farmMovimiento
            .fechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .idEstadoMovimiento = sghEstadoTabla.sghAnulado    'Anulado
            .IdUsuarioAuditoria = ml_idUsuario
        End With
   End Select
End Sub
Function AgregarDatos() As Boolean
    Dim lbAgregarDatos As Boolean
    '*********  graba tabla RELMOD  ***************
    'SCCQ 09/10/2020 Cambio28 Inicio
'    If lbDocumentoEsAutomatico = True Then
'       Dim oRsTmp As New ADODB.Recordset
'       Dim lcFiltro As String
'       lcFiltro = "tipoAlmacen='" & oRsAlmacenOrigen.Fields!idTipoLocales & "' and tipoMov='S' and tipoSuministro='" & oRsAlmacenOrigen.Fields!idTipoSuministro & "' and DocumentoId=" & mo_cmbTipoDocum.BoundText
'       Set oRsTmp = mo_ReglasFarmacia.FarmRelModDevuelveSegunFiltro(lcFiltro)
'       If oRsTmp.RecordCount = 0 Then
'          AgregarDatos = False
'       Else
'          mo_ReglasFarmacia.FarmRelModActualizaSegunFiltro lcFiltro, txtNdocum.Text
'       End If
'       oRsTmp.Close
'       Set oRsTmp = Nothing
'    End If
     'SCCQ 09/10/2020 Cambio28 Fin
     'SCCQ 20/10/2020 Cambio28 Inicio
     If lbDocumentoEsAutomatico = True Then
       lbAgregarDatos = mo_ReglasFarmacia.AgregaDatosDeNotaSalida_NumDocAutomatico(oRsAlmacenOrigen.Fields!idTipoLocales, oRsAlmacenOrigen.Fields!idTipoSuministro, CLng(mo_cmbTipoDocum.BoundText), mo_farmMovimiento, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
       txtNdocum.Text = mo_farmMovimiento.DocumentoNumero
     Else
     'SCCQ 20/10/2020 Cambio28 Fin
      lbAgregarDatos = mo_ReglasFarmacia.AgregaDatosDeNotaSalida(mo_farmMovimiento, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
     'SCCQ 20/10/2020 Cambio28 Inicio
     End If
     'SCCQ 20/10/2020 Cambio28 Fin
    txtNotaSalida.Text = mo_farmMovimiento.movNumero
    If GeneraNIenFormaAutomatica(lbAgregarDatos) Then
        If lbLaFarmaciaDestinoEsUnidosis = True Then
            CreaNIaFarmaciaUNIDOSIS mRs_Productos
        Else
            With mo_farmMovimiento
                .MovTipo = lcConstanteMovimientoEntrada
            End With
            With mo_farmMovimientoNotaIngreso
                .MovTipo = lcConstanteMovimientoEntrada
                .DocumentoFechaRecepcion = mo_farmMovimiento.fechaCreacion
            End With
            With oDoProveedores
            End With
            AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, oDoProveedores, mRs_Productos, 0, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
            MsgBox "Se creó Nota de Ingreso en forma automática", vbInformation, Me.Caption
        End If
        
    End If
    
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
    AgregarDatos = lbAgregarDatos
End Function
Function ModificarDatos() As Boolean
    Dim lbModificarDatos As Boolean
    Dim lbModificarDatosNI As Boolean
    Dim lnTotal As Double
    lbModificarDatos = mo_ReglasFarmacia.ModificaDatosDeNotaSalida(mo_farmMovimiento, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    If GeneraNIenFormaAutomatica(lbModificarDatos) Then
        If lbLaFarmaciaDestinoEsUnidosis = True Then
            CreaNIaFarmaciaUNIDOSIS mRs_Productos
        Else
            Dim oRsTmp As New Recordset
            Dim oConexion As New ADODB.Connection
            Dim oMovimiento As New farmMovimiento
            Dim oMovimientoNotaIngreso As New FarmMovimientoNotaIngreso
            '
            oConexion.Open SIGHEntidades.CadenaConexion
            Set oMovimiento.Conexion = oConexion
            Set oMovimientoNotaIngreso.Conexion = oConexion
            '
            lnTotal = mo_farmMovimiento.Total
            Set oRsTmp = mo_ReglasFarmacia.farmMovimientoSeleccionarPorTipoYnumeroDocumento(mo_farmMovimiento.DocumentoNumero, mo_farmMovimiento.DocumentoIdtipo)
            oRsTmp.Filter = "movTipo='E' and idAlmacenDestino=" & mo_farmMovimiento.IdAlmacenDestino
            If oRsTmp.RecordCount > 0 Then
                oRsTmp.MoveFirst
                '
                With mo_farmMovimiento
                    .MovTipo = lcConstanteMovimientoEntrada
                    .movNumero = oRsTmp.Fields!movNumero
                End With
                If Not oMovimiento.SeleccionarPorId(mo_farmMovimiento) Then
                   MsgBox "Fallo en Nota de Ingreso automática" & Chr(13) & oMovimiento.MensajeError
                   Exit Function
                End If
                mo_farmMovimiento.Total = lnTotal
                '
                With mo_farmMovimientoNotaIngreso
                    .MovTipo = lcConstanteMovimientoEntrada
                    .movNumero = mo_farmMovimiento.movNumero
                End With
                If Not oMovimientoNotaIngreso.SeleccionarPorId(mo_farmMovimientoNotaIngreso) Then
                   MsgBox "Fallo en Nota de Ingreso automática" & Chr(13) & oMovimientoNotaIngreso.MensajeError
                   Exit Function
                End If
                With oDoProveedores
                End With
                lbModificarDatosNI = mo_ReglasFarmacia.ModificaDatosDeNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, oDoProveedores, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
                MsgBox "Se actualizó Nota de Ingreso en forma automática", vbInformation, Me.Caption
            End If
            Set oRsTmp = Nothing
            Set oConexion = Nothing
            Set oMovimiento = Nothing
            Set oMovimientoNotaIngreso = Nothing
        End If
    Else
        ms_MensajeError = mo_ReglasFarmacia.MensajeError
    End If
    ModificarDatos = lbModificarDatos
End Function
Function AnularNS() As Boolean
    Dim lbAnularNS As Boolean
    lbAnularNS = mo_ReglasFarmacia.AnulaNotaSalida(mo_farmMovimiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0, 0)
    AnularNS = lbAnularNS
    If GeneraNIenFormaAutomatica(lbAnularNS) Then
        If lbLaFarmaciaDestinoEsUnidosis = True Then
            CreaNIaFarmaciaUNIDOSIS mRs_Productos
        Else
            Dim oRsTmp As New Recordset
            Dim oConexion As New ADODB.Connection
            Dim oMovimiento As New farmMovimiento
            Dim oMovimientoNotaIngreso As New FarmMovimientoNotaIngreso
            '
            oConexion.Open SIGHEntidades.CadenaConexion
            Set oMovimiento.Conexion = oConexion
            Set oMovimientoNotaIngreso.Conexion = oConexion
            '
            Set oRsTmp = mo_ReglasFarmacia.farmMovimientoSeleccionarPorTipoYnumeroDocumento(mo_farmMovimiento.DocumentoNumero, mo_farmMovimiento.DocumentoIdtipo)
            oRsTmp.Filter = "movTipo='E' and idAlmacenDestino=" & mo_farmMovimiento.IdAlmacenDestino
            If oRsTmp.RecordCount > 0 Then
                oRsTmp.MoveFirst
                '
                With mo_farmMovimiento
                    .MovTipo = lcConstanteMovimientoEntrada
                    .movNumero = oRsTmp.Fields!movNumero
                End With
                If Not oMovimiento.SeleccionarPorId(mo_farmMovimiento) Then
                   MsgBox "Fallo en anulación de Nota de Ingreso automática" & Chr(13) & oMovimiento.MensajeError
                   Exit Function
                End If
                mo_farmMovimiento.idEstadoMovimiento = 0  'anulado
                mo_farmMovimiento.fechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                '
                With mo_farmMovimientoNotaIngreso
                    .MovTipo = lcConstanteMovimientoEntrada
                    .movNumero = mo_farmMovimiento.movNumero
                End With
                If Not oMovimientoNotaIngreso.SeleccionarPorId(mo_farmMovimientoNotaIngreso) Then
                   MsgBox "Fallo en anulación de Nota de Ingreso automática" & Chr(13) & oMovimientoNotaIngreso.MensajeError
                   Exit Function
                End If
                With oDoProveedores
                End With
                AnularNS = mo_ReglasFarmacia.AnulaNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, 0, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
                MsgBox "Se Anuló Nota de Ingreso en forma automática", vbInformation, Me.Caption
            End If
            Set oRsTmp = Nothing
            Set oConexion = Nothing
            Set oMovimiento = Nothing
            Set oMovimientoNotaIngreso = Nothing
        End If
    Else
        ms_MensajeError = mo_ReglasFarmacia.MensajeError
    End If
End Function






Private Sub Form_Unload(Cancel As Integer)
   If SIGHEntidades.ParaAuditoria = "" Then
      LimpiarVariablesDeMemoria
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      LimpiarVariablesDeMemoria
      SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
   End If
End Sub



Private Sub grdExcel_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdExcel.Bands(0).Columns("Nro").Width = 300
    grdExcel.Bands(0).Columns("Codigo").Width = 1000
    grdExcel.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    grdExcel.Bands(0).Columns("producto").Width = 6500
    grdExcel.Bands(0).Columns("producto").Activation = ssActivationActivateNoEdit
    grdExcel.Bands(0).Columns("Cantidad").Width = 800
    grdExcel.Bands(0).Columns("Saldo").Width = 800
    grdExcel.Bands(0).Columns("Saldo").Activation = ssActivationActivateNoEdit
    grdExcel.Bands(0).Columns("error").Width = 3800
    grdExcel.Bands(0).Columns("error").Activation = ssActivationActivateNoEdit
    grdExcel.Bands(0).Columns("Seleccionar").Width = 1000
    grdExcel.Bands(0).Columns("idProducto").Hidden = True
    grdExcel.Bands(0).Columns("Producto1").Hidden = True
    grdExcel.Bands(0).Columns("idTipoSalidaBienInsumo").Hidden = True

End Sub

Private Sub txtNdocum_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNdocum

End Sub

Private Sub txtNdocum_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
    End If

End Sub

Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtObservaciones

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

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbConceptos = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_cmbAlmacenDestino = Nothing
    Set mo_cmbTipoDocum = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oRsConceptos = Nothing
    Set oRsAlmacenOrigen = Nothing
    Set lcBuscaParametro = Nothing
    Set mRs_Productos = Nothing
    Set mo_farmMovimiento = Nothing
    Set mo_farmMovimientoNotaIngreso = Nothing
    Set oDoProveedores = Nothing
End Sub

'*****Genera NI en forma automática para:
'*****DISTRIBUCION del ALMACEN ESPECIALIZADO: crea automaticamente NI hacia alguna Farmacia
'*****DEVOLUCIONES de la FARMACIA: crea automaticamente NI hacia el ALMACEN ESPECIALIZADO
'*****DISTRIBUCION de la FARMACIA: crea automaticamente hacia alguna farmacia
'*****DONACIONES del ALMACEN ESPECIALIZADO: hacia una de las Farmacias
Function GeneraNIenFormaAutomatica(lbRealizoMantenimiento As Boolean) As Boolean
         GeneraNIenFormaAutomatica = False
         If lbRealizoMantenimiento = True And (Val(mo_cmbConceptos.BoundText) = 4 And lcTipoLocalesAlmOrigen = "A") Or (Val(mo_cmbConceptos.BoundText) >= 4 And Val(mo_cmbConceptos.BoundText) <= 7 And lcTipoLocalesAlmOrigen = "F") Or (mo_cmbConceptos.BoundText = 3 And lcTipoLocalesAlmDestino = "F") Then
            GeneraNIenFormaAutomatica = True
         End If
 End Function
 
 
  '*********Es un despacho hacia la FARMACIA UNIDOSIS*********
Sub CreaNIaFarmaciaUNIDOSIS(oRsProductosConLotes1 As Recordset)
        Dim mo_farmMovimientoNotaIngreso1 As New DOfarmMovimientoNotaIngreso
        Dim oDoProveedores1 As New DoProveedores
        Dim mo_farmMovimiento2 As New farmMovimiento
        Dim oConexion As New Connection
        Dim rs As New Recordset
        Dim mo_FarmMovimientoNotaIngreso2 As New FarmMovimientoNotaIngreso
        Dim lnTotalUnidosis As Double, lnImporte As Double, lcCodigoConPunto As String
        Dim lnConvertir As Long, ActualizarDatos1 As Boolean
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        lnTotalUnidosis = 0
        If mi_Opcion <> sghEliminar Then
            oRsProductosConLotes1.MoveFirst
            Do While Not oRsProductosConLotes1.EOF
                oRsItemsUnidosis.MoveFirst
                oRsItemsUnidosis.Find "codigo='" & Trim(oRsProductosConLotes1!codigo) & "'"
                If Not oRsItemsUnidosis.EOF Then
                    lcCodigoConPunto = Trim(oRsProductosConLotes1!codigo) & SIGHEntidades.Pto
                    Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcCodigoConPunto, 1, oConexion)
                    If rs.RecordCount > 0 Then
                        lnConvertir = Val(oRsItemsUnidosis!convertir)
                        lnImporte = Round(rs!PrecioUnitario * oRsProductosConLotes1!Cantidad * lnConvertir, 2)
                        oRsProductosConLotes1!idProducto = rs!idProducto
                        oRsProductosConLotes1!codigo = lcCodigoConPunto
                        oRsProductosConLotes1!Cantidad = oRsProductosConLotes1!Cantidad * lnConvertir
                        oRsProductosConLotes1!Precio = rs!PrecioUnitario
                        oRsProductosConLotes1!Total = lnImporte
                        oRsProductosConLotes1.Update
                        lnTotalUnidosis = lnTotalUnidosis + lnImporte
                    End If
                    rs.Close
                End If
                oRsProductosConLotes1.MoveNext
            Loop
        End If
        Select Case mi_Opcion
        Case sghEliminar, sghModificar
            Set mo_farmMovimiento2.Conexion = oConexion
            Set mo_FarmMovimientoNotaIngreso2.Conexion = oConexion
            mo_farmMovimiento1.IdUsuarioAuditoria = mo_farmMovimiento.IdUsuarioAuditoria
            If mo_farmMovimiento1.movNumero <> "" Then
               mo_farmMovimientoNotaIngreso1.MovTipo = lcConstanteMovimientoEntrada
               mo_farmMovimientoNotaIngreso1.movNumero = mo_farmMovimiento1.movNumero
               mo_farmMovimientoNotaIngreso1.IdUsuarioAuditoria = mo_farmMovimiento1.IdUsuarioAuditoria
               If mo_FarmMovimientoNotaIngreso2.SeleccionarPorId(mo_farmMovimientoNotaIngreso1) Then
                    If mi_Opcion = sghModificar Then
                        mo_farmMovimiento1.Total = lnTotalUnidosis
                        ActualizarDatos1 = mo_ReglasFarmacia.ModificaDatosDeNotaIngreso(mo_farmMovimiento1, _
                                      mo_farmMovimientoNotaIngreso1, oDoProveedores1, oRsProductosConLotes1, _
                                      mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
                        If ActualizarDatos1 = True Then
                           MsgBox "Se actualizó Nota de Ingreso en FARMACIA UNIDOSIS en forma automática", vbInformation, Me.Caption
                        End If
                    Else
                        mo_farmMovimiento1.idEstadoMovimiento = sghEstadoTabla.sghAnulado
                        mo_farmMovimiento1.idUsuarioAnulacion = SIGHEntidades.Usuario
                        mo_farmMovimiento1.fechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                        ActualizarDatos1 = mo_ReglasFarmacia.AnulaNotaIngreso(mo_farmMovimiento1, _
                                   mo_farmMovimientoNotaIngreso1, 0, oRsProductosConLotes1, mo_lnIdTablaLISTBARITEMS, _
                                   mo_lcNombrePc)
                        If ActualizarDatos1 = True Then
                           MsgBox "Se Anuló Nota de Ingreso en FARMACIA UNIDOSIS en forma automática", vbInformation, Me.Caption
                        End If
                    End If
               End If
            End If
        Case sghAgregar
            With mo_farmMovimiento1
                '.movNumero
                .MovTipo = lcConstanteMovimientoEntrada
                .IdAlmacenDestino = Val(mo_cmbAlmacenDestino.BoundText)
                .idTipoConcepto = 4   'distribucion
                .DocumentoIdtipo = 3  'guía remisión
                .DocumentoNumero = txtNdocum.Text 'format(Now, SIGHEntidades.DevuelveFechaSoloFormato_DMYHMS)
                .Total = lnTotalUnidosis
                .fechaCreacion = mo_farmMovimiento.fechaCreacion
                .IdUsuarioAuditoria = mo_farmMovimiento.IdUsuarioAuditoria
                .idUsuario = mo_farmMovimiento.IdUsuarioAuditoria
                .idEstadoMovimiento = sghEstadoTabla.sghRegistrado
                .IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)
            End With
            With mo_farmMovimientoNotaIngreso1
                '.movNumero
                .MovTipo = lcConstanteMovimientoEntrada
                .DocumentoFechaRecepcion = mo_farmMovimiento.fechaCreacion
                .OrigenIdTipo = 22
                .idTipoCompra = 1
                .idTipoProceso = 1
            End With
            With oDoProveedores1
            End With
            ActualizarDatos1 = mo_ReglasFarmacia.AgregaDatosDeNotaIngreso(mo_farmMovimiento1, _
                               mo_farmMovimientoNotaIngreso1, oDoProveedores1, oRsProductosConLotes1, 0, _
                               mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
            mo_farmMovimiento.Observaciones = mo_farmMovimiento1.movNumero
            Set mo_farmMovimiento2.Conexion = oConexion
            If mo_farmMovimiento2.Modificar(mo_farmMovimiento) Then
               MsgBox "Se creó Nota de Ingreso en forma automática en " & cmbAlmDestino.Text, vbInformation, Me.Caption
            End If
        End Select
        oConexion.Close
        Set mo_farmMovimientoNotaIngreso1 = Nothing
        Set oDoProveedores1 = Nothing
        Set mo_farmMovimiento2 = Nothing
        Set oConexion = Nothing
        Set rs = Nothing
End Sub



Sub CargaItemsDebajoDeStockMinimo()
 Dim oRsTmp As New Recordset
 CargaInventarioExcel.Enabled = True
 grdHistorico.Caption = "Lista de Medicamentos/Insumos por debajo de su STOCK MINIMO"
 Set oRsTmp = mo_ReglasFarmacia.FarmaciaItemsPorDebajoStockMinimo
 
 If cmbAlmOrigen.Text <> "" Then
    oRsTmp.Filter = "idAlmacen=" & mo_cmbAlmacenOrigen.BoundText
 End If
 Set grdHistorico.DataSource = oRsTmp
End Sub

