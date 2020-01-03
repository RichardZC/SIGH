VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucNotaIngresoICI 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13050
   ScaleHeight     =   5685
   ScaleWidth      =   13050
   Begin VB.TextBox txtBusca 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8190
      TabIndex        =   6
      Top             =   5370
      Width           =   1995
   End
   Begin VB.ComboBox cmbOrden 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "ucNotaIngresoICI.ctx":0000
      Left            =   6240
      List            =   "ucNotaIngresoICI.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5370
      Width           =   1905
   End
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   2475
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   4366
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
      Appearance      =   "ucNotaIngresoICI.ctx":0026
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   5325
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   9393
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   68157460
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Override        =   "ucNotaIngresoICI.ctx":0062
      Caption         =   "Productos ICI"
   End
   Begin VB.Label lblPrecios 
      AutoSize        =   -1  'True
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10320
      TabIndex        =   7
      Top             =   5430
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Orden:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   4
      Top             =   5430
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Teclas de ayuda: <F10> = Agregar         <Supr>  = Eliminar "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   5430
      Width           =   4995
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12240
      TabIndex        =   2
      Top             =   5430
      Width           =   555
   End
End
Attribute VB_Name = "ucNotaIngresoICI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Control para Items de la Nota de Ingreso
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim gridInfra As New GridInfragistic
Dim mb_CargandoProductos As Boolean
Dim mRs_Productos As New Recordset
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim lnMaximoNroItems As Long
Dim ml_movNumero As String
Dim lcSql As String
Dim ml_idProducto As Long
Dim ml_IdAlmacen As Long
Dim dTotalIngresado  As Double
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_idTipoPrecioParaNiNs As Long
Dim ml_idTipoConcepto As Long
Dim mb_EsUnaDonacionOestrategico As Long
Dim ml_DocumentoNumero As String
'kike 2017
Property Get DocumentoNumero() As String
    DocumentoNumero = ml_DocumentoNumero
End Property
Property Let EsUnaDonacionOestrategico(lValue As Long)
   mb_EsUnaDonacionOestrategico = lValue
End Property

Property Let TipoConcepto(lValue As Long)
   ml_idTipoConcepto = lValue
End Property

Property Let TipoPrecioParaNiNs(lValue As Long)
   ml_idTipoPrecioParaNiNs = lValue
End Property


Property Get idProducto() As Long
    idProducto = ml_idProducto
End Property

Property Let IdAlmacen(lValue As Long)
   ml_IdAlmacen = lValue
End Property
Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property


Sub inicializar()
    Set mRs_Productos = New Recordset
    GenerarRecordsetProductos
    lnMaximoNroItems = BuscarMaximoItemsEnParametros()

End Sub

Function BuscarMaximoItemsEnParametros() As Long
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
        Set lcBuscaParametro = Nothing
End Function

Sub AgregaProducto(lbPulsaF10 As Boolean)
On Error GoTo ErrAdd
    grdProductos.SetFocus
    If lbPulsaF10 Then
       SendKeys "{Tab}"
    End If
    mb_CargandoProductos = True
    AgregaRegistro
    mb_CargandoProductos = False
    Totalizar
    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
ErrAdd:
End Sub


Sub AgregaRegistro()
    On Error GoTo errAR
    With mRs_Productos
        .AddNew
        .Fields!idProducto = 0
        .Fields!codigo = ""
        .Fields!nombreProducto = ""
        .Fields!Lote = ""
        .Fields!FechaVencimiento = Date
        .Fields!Cantidad = 0
        .Fields!Precio = 0
        .Fields!total = 0
    End With
errAR:
End Sub



Sub CargaProductosPorMovNumero()
   Dim rs As Recordset
   Dim oConexion As New ADODB.Connection
   oConexion.Open SIGHEntidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   Set rs = mo_ReglasFarmacia.FarmMovimientosDetalleDevuelveTodosItems(oConexion, "E", ml_movNumero)
   'kike 2017
   If rs.RecordCount > 0 Then
        rs.MoveFirst
        If Not IsNull(rs!DocumentoNumero) Then
           ml_DocumentoNumero = rs!DocumentoNumero
        End If
   End If
   '
   CargarItemsALaGrilla rs
   oConexion.Close
   Set oConexion = Nothing
End Sub


Sub CargarItemsALaGrilla(rs As Recordset)
    mb_CargandoProductos = True
    
    Do While Not rs.EOF
        mRs_Productos.AddNew
        mRs_Productos!idProducto = rs!idProducto
        mRs_Productos!codigo = rs!codigo
        mRs_Productos!nombreProducto = rs!Nombre
        mRs_Productos!Lote = rs!Lote
        mRs_Productos!FechaVencimiento = rs!FechaVencimiento
        mRs_Productos!Cantidad = rs!Cantidad
        mRs_Productos!Precio = rs!Precio
        mRs_Productos!total = rs!total
        mRs_Productos!idTipoSalidaBienInsumo = rs!idTipoSalidaBienInsumo
        mRs_Productos!registroSanitario = rs!registroSanitario
        If Not IsNull(rs!esPaquete) Then
          mRs_Productos!esPaquete = rs!esPaquete
        End If
        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    Totalizar
    Set grdProductos.DataSource = mRs_Productos
End Sub

Sub CargaProductosPorTemporal(rs As Recordset)
    Dim oRs1 As New Recordset
    Dim lcCodigo As String, lcMensaje As String
    mb_CargandoProductos = True
    GenerarRecordsetProductos
    lcMensaje = ""
    Do While Not rs.EOF
        lcCodigo = Trim(rs!medcod)
        Set oRs1 = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb(lcCodigo, 1, 5)
        If oRs1.RecordCount > 0 Then
            mRs_Productos.AddNew
            mRs_Productos!idProducto = oRs1!idProducto
            mRs_Productos!codigo = lcCodigo
            mRs_Productos!nombreProducto = oRs1!nombreProducto
            mRs_Productos!Lote = Trim(rs!medLote)
            mRs_Productos!FechaVencimiento = rs!medFechVto
            mRs_Productos!Cantidad = rs!movCantid
            mRs_Productos!Precio = rs!movPrecio
            mRs_Productos!total = Round(rs!movCantid * rs!movPrecio, 2)
            If mb_EsUnaDonacionOestrategico > 0 Then
               mRs_Productos!idTipoSalidaBienInsumo = mb_EsUnaDonacionOestrategico
            Else
               mRs_Productos!idTipoSalidaBienInsumo = SIGHEntidades.ElijeSiEsEstrategicoDevuelveId(oRs1.Fields("idTipoSalidaBienInsumo").Value)
            End If
            mRs_Productos!registroSanitario = rs!medRegsan
            mRs_Productos.Update
        Else
            lcMensaje = lcMensaje & "No tiene el CODIGO SISMED: " & lcCodigo & Chr(13)
        End If
        oRs1.Close
        rs.MoveNext
    Loop
    Set oRs1 = Nothing
    If lcMensaje <> "" Then
       MsgBox "No se podrá IMPORTAR los siguientes códigos DIGEMID:" & Chr(13) & lcMensaje & Chr(13) & _
              "ANTES DE USAR ESTA OPCION DEBE IMPORTAR LOS MEDICAMENTOS/INSUMOS", vbInformation
    End If
    mb_CargandoProductos = False
    Totalizar
    Set grdProductos.DataSource = mRs_Productos
End Sub


Sub Totalizar()
    Dim rsProductos As New ADODB.Recordset
    Set rsProductos = mRs_Productos.Clone
    dTotalIngresado = 0
    If rsProductos.RecordCount > 0 Then
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                rsProductos.Fields!total = Round(rsProductos.Fields!Cantidad * rsProductos.Fields!Precio, 2)
                rsProductos.Update
                dTotalIngresado = dTotalIngresado + rsProductos!total
                'rsProductos.Update
                rsProductos.MoveNext
            Loop
        End If
    End If
    lblTotal.Caption = "Total:    " & Format(dTotalIngresado, "####,###,##0.00")
    Select Case ml_idTipoPrecioParaNiNs
    Case 1
       'grdProductos.Bands(0).Columns("Precio").Header.Caption = "Pr.Compra"
       lblPrecios.Caption = "<Se usará Precio de Compra>"
    Case 2
       lblPrecios.Caption = "<Se usará Precio de Distribución>"
    Case 3
       lblPrecios.Caption = "<Se usará Precio de Venta>"
    Case 4
       lblPrecios.Caption = "<Se usará Precio de Donación>"
    Case 5
       Select Case ml_idTipoConcepto
       Case 1, 2, 8, 9, 20 'compra, encargo,transferencias, Trasnfercia UE, ajuste inventario
           lblPrecios.Caption = "<Se usará Precio de Compra>"
       Case 3  'donaciones
           lblPrecios.Caption = "<Se usará Precio de Donación>"
       End Select
    End Select

End Sub



Private Sub cmbOrden_Click()
  If cmbOrden.ListIndex = 0 Then
        mRs_Productos.Sort = "codigo"
    Else
        mRs_Productos.Sort = "NombreProducto"
    End If
    grdProductos.Refresh
End Sub



Private Sub grdProductos_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
  Totalizar
End Sub

'Eventos de la grilla de servicios
Private Sub grdProductos_AfterRowActivate()
        If mb_CargandoProductos Then
            Exit Sub
        End If
End Sub


Private Sub grdProductos_AfterRowsDeleted()
    If ml_ultimoProductoEliminado > 0 Then
        mo_ProductosEliminados.Add ml_ultimoProductoEliminado
        ml_ultimoProductoEliminado = 0
        Totalizar
    Else
        Totalizar
        Set grdProductos.DataSource = mRs_Productos
    End If

End Sub



Private Sub grdProductos_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
        If mb_FilaEditable Then
            'Si la fila es editable y estamos en la celda de codigo se completa los datos
            'del producto
            Select Case grdProductos.ActiveCell.Column.Key
            Case "Codigo"
                ConfigurarProductoPorCodigo grdProductos
            Case "Cantidad"
                Dim oRow As SSRow
                Set oRow = grdProductos.ActiveCell.Row
                If oRow.Cells("cantidad").Value > 0 Then
                   Totalizar
                Else
                   MsgBox "La cantidad debe ser Mayor a CERO  ", vbInformation, "Mensaje"
                   oRow.Cells("cantidad").Value = 0
                End If
            Case "FechaVencimiento"
                If 1 = 1 Then
                    If ItemYaExiste(grdProductos.ActiveRow.Cells("idproducto").Value, grdProductos.ActiveRow.Cells("Lote").Value, grdProductos.ActiveRow.Cells("FechaVencimiento").Value, grdProductos.ActiveRow.Cells("idTipoSalidaBienInsumo").Value, grdProductos.ActiveRow.Bookmark) Then
                        grdProductos.ActiveRow.Cells("codigo").Value = ""
                        grdProductos.ActiveRow.Cells("idproducto").Value = 0
                        grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
                        grdProductos.ActiveRow.Cells("Lote").Value = ""
                    End If
                End If
            End Select
            grillaBusqueda.Visible = False
        End If
        
End Sub

Private Sub grdProductos_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    'Si la fila no es editable, cancela cualquier cambio en la fila
    If Not mb_FilaEditable Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub grdProductos_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    mb_FilaEditable = True
End Sub







Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdProductos
End Sub

Private Sub grdProductos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    OnKeyDown grdProductos, KeyCode
End Sub

Private Sub grdProductos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    OnKeyPress grdProductos, KeyAscii
End Sub


Sub ConfigurarProductoPorCodigo(oGrilla As SSUltraGrid)
Dim rs As Recordset
Dim oRow As SSRow
Dim lcFiltro As String
Dim lnPrecioUnitario As Double
    If ml_idTipoPrecioParaNiNs = 0 Then
       MsgBox "Debe elegir el Concepto antes de Registrar Productos", vbInformation, "Farmacia"
       Exit Sub
    End If
    Set oRow = oGrilla.ActiveCell.Row
    If IsNull(oRow.Cells("codigo").Value) Or Trim(oRow.Cells("codigo").Value) = "" Then
        Exit Sub
    End If
    lcFiltro = oRow.Cells("codigo").Value
 
    Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb(oRow.Cells("codigo").Value, 1, 5)
    If rs.RecordCount > 0 Then
        'Busca si ya existe el producto
        If Not ItemYaExiste(rs.Fields("idproducto").Value, "debb", Date, 1, 0) Then
            lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(rs.Fields("idproducto").Value, ml_idTipoPrecioParaNiNs)
            oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
            oRow.Cells("NombreProducto").Value = rs.Fields("NombreProducto").Value
            If mb_EsUnaDonacionOestrategico > 0 Then
               oRow.Cells("idTipoSalidaBienInsumo").Value = mb_EsUnaDonacionOestrategico
            Else
               oRow.Cells("idTipoSalidaBienInsumo").Value = SIGHEntidades.ElijeSiEsEstrategicoDevuelveId(rs.Fields("idTipoSalidaBienInsumo").Value)
            End If
            oRow.Cells("precio").Value = lnPrecioUnitario
            If IsNull(rs!esPaquete) Then
                oRow.Cells("lote").Value = ""
                oRow.Cells("FechaVencimiento").Value = Date
                oRow.Cells("esPaquete").Value = False
            Else
                oRow.Cells("lote").Value = IIf(rs!esPaquete = True, WxLOTEpaquete, "")
                oRow.Cells("FechaVencimiento").Value = IIf(rs!esPaquete = True, CDate(WxFVENCIMIENTOpaquete), Date)
                oRow.Cells("esPaquete").Value = IIf(rs!esPaquete = True, 1, 0)
                oRow.Cells("RegistroSanitario").Value = IIf(rs!esPaquete = True, WxREGSANITARIOpaquete, "")
            End If
            oRow.Cells("Total").Value = 0
            oRow.Cells("cantidad").Value = 0
            SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Cod:" & Trim(rs!codigo)

        End If
    End If

End Sub

'***************daniel barrantes**************
'***************Verifica si YA SE REGISTRO el ITEM (al momento de registrar)
'***************
Function ItemYaExiste(lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date, idTipoSalidaBienInsumo As Long, lnFila As Integer) As Boolean
        Dim lbExiste As Boolean
        Dim oRsTmp As New ADODB.Recordset
        Dim lnRow As Integer
        Set oRsTmp = mRs_Productos.Clone
        ItemYaExiste = False
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           lnRow = 1
           Do While Not oRsTmp.EOF
                If oRsTmp.Fields!idProducto = lnIdProducto And Trim(oRsTmp.Fields!Lote) = Trim(lcLote) And oRsTmp.Fields!FechaVencimiento = ldFechaVencimiento And idTipoSalidaBienInsumo = oRsTmp.Fields!idTipoSalidaBienInsumo And lnFila <> lnRow Then
                   ItemYaExiste = True
                   MsgBox "Este Producto/Tipo/Lote/FechaVencimiento ya está registrado", vbInformation, "Farmacia"
                   Exit Do
                End If
                oRsTmp.MoveNext
                lnRow = lnRow + 1
           Loop
        End If
        oRsTmp.Close
End Function


Sub OnKeyDown(oGrilla As SSUltraGrid, KeyCode As UltraGrid.SSReturnShort)
        If Not oGrilla.ActiveCell Is Nothing Then
            Select Case oGrilla.ActiveCell.Column.Key
            Case "Cantidad"
                Select Case Val(Chr(KeyCode))
                Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
                Case Else
                    KeyCode = 0
                End Select
            Case "NombreProducto"
                Select Case KeyCode
                Case vbKeyBack
                Case vbKeyEscape
                    Set grillaBusqueda.DataSource = Nothing
                    grillaBusqueda.Visible = False
                Case vbKeyReturn
                Case vbKeyDown
                    On Error Resume Next
                    grillaBusqueda.SetFocus
                Case vbKeyLeft
                End Select
            End Select
        End If
        Select Case KeyCode
        Case vbKeyF10
            AgregaProducto (True)
        End Select
End Sub

Sub OnKeyPress(oGrilla As SSUltraGrid, KeyAscii As UltraGrid.SSReturnShort)
                
        'Si la fila no es editable, cancela cualquier cambio en la fila
        If Not mb_FilaEditable Then
            Exit Sub
        End If
        
        If oGrilla.ActiveCell Is Nothing Then
            Exit Sub
        End If

        If oGrilla.ActiveCell.Column.Key = "Codigo" And KeyAscii = 13 Then
            SendKeys "{Tab}"
            If Trim(oGrilla.ActiveCell.GetText) <> "" Then
                SendKeys "{Tab}"
                SendKeys "{Tab}"
            End If
            Exit Sub
        End If
        If oGrilla.ActiveCell.Column.Key = "Lote" Then
            If KeyAscii = 13 Then
               SendKeys "{Tab}"
            End If
            Exit Sub
        End If
        If oGrilla.ActiveCell.Column.Key = "FechaVencimiento" Then
            If KeyAscii = 13 Then
               SendKeys "{Tab}"
            End If
            Exit Sub
        End If
        If oGrilla.ActiveCell.Column.Key = "Cantidad" Then
            If KeyAscii = 13 Then
               SendKeys "{Tab}"
            End If
            Exit Sub
        End If


        If oGrilla.ActiveCell.Column.Key = "NombreProducto" Then
            If ml_idTipoPrecioParaNiNs = 0 Then
               MsgBox "Debe elegir el Concepto antes de Registrar Productos", vbInformation, "Farmacia"
               Exit Sub
            End If
            Select Case KeyAscii
            Case vbKeyEscape
                If Trim(oGrilla.ActiveCell.GetText) = "" Then
                    grillaBusqueda.Visible = False
                    Set grillaBusqueda.DataSource = Nothing
                End If
            Case vbKeyReturn
            Case vbKeyDown
            Case vbKeyLeft
            Case Else
                Dim sNombre As String
                Select Case KeyAscii
                Case vbKeyBack
                    sNombre = oGrilla.ActiveCell.GetText
                Case Else
                    sNombre = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
                End Select
                
               
                Dim rs As New Recordset
                Set rs = mo_AdminCaja.BienesFiltrarParaCajero(sNombre, 1, 5)
                Set grillaBusqueda.DataSource = rs
                'grillaBusqueda.Left = oGrilla.Left
                If mRs_Productos.RecordCount < 7 Then
                   grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.Rect.Bottom * Screen.TwipsPerPixelY
                Else
                   grillaBusqueda.Top = 0
                End If
                grillaBusqueda.Visible = True
                grillaBusqueda.Enabled = True
                
            End Select
        End If
        If oGrilla.ActiveCell.Column.Key = "Precio" Then
            If KeyAscii = 13 Then
               SendKeys "{Tab}"
               AgregaProducto (False)
            End If
            Exit Sub
        End If
        
        

End Sub


Sub GenerarRecordsetProductos()
    If mRs_Productos.State = 1 Then Set mRs_Productos = Nothing
    With mRs_Productos
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "Codigo", adVarChar, 20
          .Fields.Append "NombreProducto", adChar, 300
          .Fields.Append "idTipoSalidaBienInsumo", adInteger
          .Fields.Append "Lote", adVarChar, 15
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "MovNumeroS", adChar, 9, adFldIsNullable
          .Fields.Append "RegistroSanitario", adVarChar, 50, adFldIsNullable
          .Fields.Append "NumeroDocumento", adVarChar, 20, adFldIsNullable 'Frank 07082015
          .Fields.Append "esPaquete", adBoolean
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    'Set grdProductos.DataSource = mRs_Productos
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
     On Error GoTo ConfigEstilo
     grdProductos.Bands(0).Columns("esPaquete").Hidden = True
     grdProductos.Bands(0).Columns("IdProducto").Hidden = True
     grdProductos.Bands(0).Columns("MovNumeroS").Hidden = True
     grdProductos.Bands(0).Columns("IdProducto").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("codigo").Width = 1000
     grdProductos.Bands(0).Columns("NombreProducto").Width = 7000
     '
     grdProductos.ValueLists.Add "TipoSalida"
     grdProductos.ValueLists("TipoSalida").ValueListItems.Add 1, "Ventas"
     grdProductos.ValueLists("TipoSalida").ValueListItems.Add 2, "Interv.Sanit"
     grdProductos.ValueLists("TipoSalida").ValueListItems.Add 3, "Vta/IntervSanit"
     grdProductos.ValueLists("TipoSalida").ValueListItems.Add 4, "Donaciones"
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").ValueList = "TipoSalida"
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").Style = ssStyleDropDownList
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").Width = 800
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").Header.Caption = "Tipo"
     '
     grdProductos.Bands(0).Columns("FechaVencimiento").Width = 1500
     grdProductos.Bands(0).Columns("FechaVencimiento").Format = SIGHEntidades.DevuelveFechaSoloFormato_DMY
     grdProductos.Bands(0).Columns("cantidad").Width = 800
     grdProductos.Bands(0).Columns("cantidad").Format = "###0"
     grdProductos.Bands(0).Columns("Precio").Width = 700
     grdProductos.Bands(0).Columns("Precio").Format = "#0.000"
     grdProductos.Bands(0).Columns("Total").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Total").Format = "#0.00"
     grdProductos.Bands(0).Columns("NumeroDocumento").Hidden = True 'Frank 07082015a
    
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHEntidades.GrillaConFilasBicolor
    
End Sub





Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, SIGHEntidades.GrillaConFilasBicolor
End Sub
Private Sub InicializarLaGrillaBusqueda(oGrilla As SSUltraGrid)
    On Error GoTo errInic
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    oGrilla.Bands(0).Columns("Activo").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Nombre").Width = 7800
    
    oGrilla.Bands(0).Columns("preciounitario").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHEntidades.GrillaConFilasBicolor
errInic:
End Sub
Private Sub grillaBusqueda_DblClick()
Dim fila As New Record
Dim lnIdProductoBusqueda As Long
    'debb-hra-ya en version Polsalud
    On Error GoTo ErrGrillaBusqueda
    lnIdProductoBusqueda = grillaBusqueda.ActiveRow.Cells("idproducto").Value
    If ItemYaExiste(lnIdProductoBusqueda, "debb", Date, 1, 0) Then
        grdProductos.ActiveRow.Cells("codigo").Value = ""
        grdProductos.ActiveRow.Cells("idproducto").Value = 0
        grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
        grdProductos.ActiveRow.Cells("Lote").Value = ""
        grdProductos.ActiveRow.Cells("FechaVencimiento").Value = Date
    Else
        RefrescarDatos
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
        SendKeys "{Tab}"
        SendKeys "{Tab}"
    End If
ErrGrillaBusqueda:
End Sub
Sub RefrescarDatos()
    Dim fila As New Record
    Dim lnPrecioUnitario As Double
    
    If Not grillaBusqueda.ActiveRow Is Nothing Then
               lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(grillaBusqueda.ActiveRow.Cells("idproducto").Value, ml_idTipoPrecioParaNiNs)
               grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
               grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
               grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
               If mb_EsUnaDonacionOestrategico > 0 Then
                  grdProductos.ActiveRow.Cells("idTipoSalidaBienInsumo").Value = mb_EsUnaDonacionOestrategico
               Else
                  grdProductos.ActiveRow.Cells("idTipoSalidaBienInsumo").Value = SIGHEntidades.ElijeSiEsEstrategicoDevuelveId(grillaBusqueda.ActiveRow.Cells("idTipoSalidaBienInsumo").Value)
               End If
               If IsNull(grillaBusqueda.ActiveRow.Cells("esPaquete").Value) Then
                    grdProductos.ActiveRow.Cells("lote").Value = ""
                    grdProductos.ActiveRow.Cells("fechaVencimiento").Value = Date
                    grdProductos.ActiveRow.Cells("esPaquete").Value = False
               Else
                    grdProductos.ActiveRow.Cells("lote").Value = IIf(grillaBusqueda.ActiveRow.Cells("esPaquete").Value = True, WxLOTEpaquete, "")
                    grdProductos.ActiveRow.Cells("fechaVencimiento").Value = IIf(grillaBusqueda.ActiveRow.Cells("esPaquete").Value = True, CDate(WxFVENCIMIENTOpaquete), Date)
                    grdProductos.ActiveRow.Cells("esPaquete").Value = grillaBusqueda.ActiveRow.Cells("esPaquete").Value
                    grdProductos.ActiveRow.Cells("RegistroSanitario").Value = WxREGSANITARIOpaquete
               End If
               grdProductos.ActiveRow.Cells("precio").Value = lnPrecioUnitario
               grdProductos.ActiveRow.Cells("Total").Value = 0
               grdProductos.ActiveRow.Cells("cantidad").Value = 0
               SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Cod:" & Trim(grillaBusqueda.ActiveRow.Cells("CODIGO").Value)
               Totalizar
    End If

End Sub

Private Sub grillaBusqueda_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)

    Select Case KeyCode
    Case vbKeyEscape
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
    Case vbKeyReturn
        grillaBusqueda_DblClick
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
    End Select
    
End Sub





Private Sub txtBusca_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtBusca.Text <> "" Then
            txtBusca.Text = Trim(txtBusca.Text)
            mRs_Productos.MoveFirst
            If cmbOrden.ListIndex = 0 Then
               mRs_Productos.Find "codigo='" & txtBusca.Text & "'"
            Else
               Do While Not mRs_Productos.EOF
                  If Left(mRs_Productos!nombreProducto, Len(txtBusca.Text)) = UCase(txtBusca.Text) Then
                     Exit Do
                  End If
                  mRs_Productos.MoveNext
               Loop
            End If
            grdProductos.Refresh
      End If
   End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   grdProductos.Height = UserControl.Height - UserControl.Label1.Height - 100
   
   Label1.Top = UserControl.Height - UserControl.Label1.Height - 50
   Label2.Top = UserControl.Height - UserControl.Label1.Height - 50
   cmbOrden.Top = UserControl.Height - UserControl.Label1.Height - 100
   txtBusca.Top = UserControl.Height - UserControl.Label1.Height - 100
   lblPrecios.Top = UserControl.Height - UserControl.Label1.Height - 100
   lblTotal.Top = UserControl.Height - UserControl.Label1.Height - 50
   lblTotal.Left = UserControl.Width - Len(lblTotal) * 150
End Sub

Sub LimpiarGrilla()
        If mRs_Productos Is Nothing Then
            Exit Sub
        End If

        Set grdProductos.DataSource = Nothing

        If mRs_Productos.RecordCount > 0 Then
            mRs_Productos.MoveFirst
            Do While Not mRs_Productos.EOF
                mRs_Productos.Delete
                mRs_Productos.Update
                mRs_Productos.MoveNext
            Loop
        End If
        CargaProductosPorMovNumero
End Sub



Property Get DevuelveProductos() As Recordset
    'Se debe utilizar un clon del recrdset, ya que si se trabaja directamente con el recordset
    'que esta asociado a la grilla ocurre errores en los metodos movenext, movefirst, etc.
    Set DevuelveProductos = mRs_Productos.Clone()
End Property
Property Get DevuelveTotal() As Double
    DevuelveTotal = dTotalIngresado
End Property

'****solo para que puedan registrar MEDICAMENTOS - FASE DE PRUEBAS HRA-AGOS-2009
Sub AgregaProductoParaSaldosDePruebas(lnIdProducto As Long, lcCodigo As String, lcNombreProducto As String, lcLote As String, ldFechaVencimiento As Date, lnCantidad As Long, lnPrecio As Double)
    With mRs_Productos
        .AddNew
        .Fields!idProducto = lnIdProducto
        .Fields!codigo = lcCodigo
        .Fields!nombreProducto = lcNombreProducto
        .Fields!Lote = lcLote
        .Fields!FechaVencimiento = ldFechaVencimiento
        .Fields!Cantidad = lnCantidad
        .Fields!Precio = lnPrecio
        .Fields!total = Round(lnCantidad * lnPrecio, 2)
    End With
End Sub

