VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucCatalogos 
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3120
   ScaleWidth      =   13110
   Begin VB.CommandButton btnagregar 
      Caption         =   "Agregar"
      Height          =   360
      Left            =   9240
      TabIndex        =   8
      Top             =   2640
      Width           =   990
   End
   Begin VB.CommandButton btneliminar 
      Caption         =   "Eliminar"
      Height          =   360
      Left            =   10320
      TabIndex        =   6
      Top             =   2640
      Width           =   990
   End
   Begin UltraGrid.SSUltraGrid grillacod 
      Height          =   1755
      Left            =   1440
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   3096
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
      Appearance      =   "ucCatalogos.ctx":0000
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grillaitem 
      Height          =   1755
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   3096
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
      Appearance      =   "ucCatalogos.ctx":003C
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grillahis 
      Height          =   1755
      Left            =   8760
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   3096
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
      Appearance      =   "ucCatalogos.ctx":0078
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grilladx 
      Height          =   1755
      Left            =   6840
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   3096
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
      Appearance      =   "ucCatalogos.ctx":00B4
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   1755
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   3096
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
      Appearance      =   "ucCatalogos.ctx":00F0
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   4471
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
      Caption         =   "Servicios"
   End
   Begin VB.Label lblTeclasDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teclas de ayuda: <F10> = Agregar      <Supr>  = Eliminar       <Espace> Habilita Descripción, IdSubClas "
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
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   8670
   End
End
Attribute VB_Name = "ucCatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Control para Items de la Nota de Salida
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim gridInfra As New GridInfragistic
Dim mb_CargandoProductos As Boolean
Dim mRs_Productos As New Recordset
Dim mRs_Dx As New Recordset
Dim mRs_His As New Recordset

Dim oRsSaldosEnEsteMomento As New Recordset
Dim oRsSaldosEnEsteMoment1 As New Recordset
Dim oRsSaldosEnEsteMoment2 As New Recordset
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim lnMaximoNroItems As Long
Dim ml_movNumero As String
Dim lcSql As String
'Dim lcwhere As Long

Dim ml_desc As String

Dim ml_IdProducto As Long
Dim ml_IdAlmacen As Long
Dim dTotalIngresado  As Double
Dim ml_idTipoPrecioParaNiNs As Long
Dim mb_MuestraLoteParaDespachoNS As Boolean
Dim LdFechaMinimaDespacho As Date
Dim mo_Formulario As New sighEntidades.Formulario
Dim lbDesdeCargarExcel As Boolean
Const lnTopGrilla = 3000


Dim lcwhere As Long


Dim lidatencionC As String

Property Let cidatencion(lValue As String)
   lidatencionC = lValue
End Property

Property Let FechaMinimaDespacho(lValue As Date)
   LdFechaMinimaDespacho = lValue
End Property

Property Let MuestraLoteParaDespachoNS(lValue As Boolean)
   mb_MuestraLoteParaDespachoNS = lValue
End Property
Property Let TipoPrecioParaNiNs(lValue As Long)
   ml_idTipoPrecioParaNiNs = lValue
End Property

Property Let IdAlmacen(lValue As Long)
   ml_IdAlmacen = lValue
End Property
Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property


Sub Inicializar()
    Set mRs_Productos = New Recordset
    GenerarRecordsetProductos
    lnMaximoNroItems = BuscarMaximoItemsEnParametros()
End Sub

Function BuscarMaximoItemsEnParametros() As Long
        BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
End Function

Sub AgregaProducto(lbPulsaF10 As Boolean)
    On Error GoTo AddP
    grdProductos.SetFocus
    If lbPulsaF10 Then
       SendKeys "{Tab}"
       SendKeys "{Tab}"
    End If
    mb_CargandoProductos = True
    AgregaRegistro
    mb_CargandoProductos = False
    'Totalizar
    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
AddP:
End Sub


Sub AgregaRegistro()
    On Error GoTo errAR
    With mRs_Productos
        .AddNew
        .Fields!ID = 0
        .Fields!idAtencion = ""
        .Fields!Descripcion_Tipo_Item = ""
        .Fields!Fg_Tipo = ""
        .Fields!Codigo = ""
        .Fields!NombreProducto = ""
        .Fields!IdSubClasificacion = ""
        .Fields!labConfHIS = ""
       
    End With
errAR:
End Sub


Sub CargaProductosPorIdAtencion()
   Dim rs As Recordset
   Dim oConexion As New ADODB.Connection
   oConexion.Open sighEntidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   Set rs = mo_ReglasFarmacia.usp_listarcatalogoidat(oConexion, lidatencionC)
   CargarItemsALaGrilla rs
   oConexion.Close
   Set oConexion = Nothing
End Sub

Sub CargarItemsALaGrilla(rs As Recordset)
    Dim oRsTmp As New ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oFarmMovimientoDetalle.Conexion = oConexion
    mb_CargandoProductos = True
    Do While Not rs.EOF
        mRs_Productos.AddNew
        mRs_Productos!ID = rs!ID
        mRs_Productos!idAtencion = rs!idAtencion
        mRs_Productos!Descripcion_Tipo_Item = rs!Descripcion_Tipo_Item
        mRs_Productos!Fg_Tipo = rs!Fg_Tipo
        mRs_Productos!Codigo = rs!idProducto
        mRs_Productos!NombreProducto = rs!NombreMINSA
        mRs_Productos!IdSubClasificacion = rs!IdSubclasificacionDx
        mRs_Productos!labConfHIS = rs!valores
        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    Set grdProductos.DataSource = mRs_Productos
End Sub

Private Sub btnagregar_Click()

            AgregaProducto (True)
        
End Sub

Private Sub btneliminar_Click()

Dim msgvalue As Integer
 
msgvalue = MsgBox("¿Está seguro de eliminar permanentemente el registro?", vbInformation + vbYesNo, "Mensaje de Alerta")
 
Select Case msgvalue
 
Case 6 'Yes
  Dim rs As Recordset
   Dim oConexion As New ADODB.Connection
   oConexion.Open sighEntidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   Set rs = mo_ReglasFarmacia.usp_eliminacatalogoidat(oConexion, lcwhere)
   LimpiarGrilla
   oConexion.Close
   Set oConexion = Nothing
 
Case 7 'No
 
    Exit Sub
 
End Select
 
End Sub


  
   



'Eventos de la grilla de servicios
Private Sub grdProductos_AfterRowActivate()
        If mb_CargandoProductos Then
            Exit Sub
        End If
End Sub


Sub Totalizar()
    Dim rsProductos As New ADODB.Recordset
    Set rsProductos = mRs_Productos.Clone
    dTotalIngresado = 0
    If rsProductos.RecordCount > 0 Then
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
              
                rsProductos.MoveNext
            Loop
        End If
    End If
    
End Sub


Private Sub grdProductos_AfterRowsDeleted()
   
 If lcwhere > 0 Then
    
    MsgBox "Para eliminar este registro clic el boton eliminar", vbInformation, "Atenciones"
    LimpiarGrilla
 Else
   
    If ml_ultimoProductoEliminado > 0 Then
        mo_ProductosEliminados.Add ml_ultimoProductoEliminado
        ml_ultimoProductoEliminado = 0
        'CargaProductosPorIdAtencion
        
    Else
    
        Set grdProductos.DataSource = mRs_Productos
        'CargaProductosPorIdAtencion
    End If
    
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

Private Sub grdProductos_Click()

        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = mRs_Productos.Clone
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                lcwhere = grdProductos.ActiveRow.Cells("Id").Value
               
                   Exit Do
                oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close

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

'***************daniel barrantes**************
'***************Verifica si YA SE REGISTRO el ITEM (al momento de registrar)
'***************
Function ItemYaExiste(lnIdProducto As String) As Boolean
        Dim lbExiste As Boolean
        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = mRs_Productos.Clone
        ItemYaExiste = False
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                If Trim(oRsTmp.Fields!Codigo) = Trim(lnIdProducto) Then
                   ItemYaExiste = True
                   MsgBox "Este Servicio ya está registrado", vbInformation, "Atenciones"
                   Exit Do
                End If
                oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close
End Function

Function ItemYaExisteC(lnIdProducto As String) As Boolean
        Dim lbExiste As Boolean
        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = mRs_Productos.Clone
        ItemYaExisteC = False
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                If Trim(oRsTmp.Fields!NombreProducto) = Trim(lnIdProducto) Then
                   ItemYaExisteC = True
                   MsgBox "Este Servicio ya está registrado", vbInformation, "Atenciones"
                   Exit Do
                End If
                oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close
End Function

Function ItemYaExisteSinLote(lnIdProducto As Long, lnIdTipoSalidaBienInsumoSaldo As Long) As Boolean
        Dim lbExiste As Boolean
        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = mRs_Productos.Clone
        ItemYaExisteSinLote = False
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                If oRsTmp.Fields!idProducto = lnIdProducto And oRsTmp.Fields!idTipoSalidaBienInsumo = lnIdTipoSalidaBienInsumoSaldo Then
                   ItemYaExisteSinLote = True
                   MsgBox "Este Producto/Tipo ya está registrado", vbInformation, "Farmacia"
                   Exit Do
                End If
                oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close
End Function

Sub OnKeyDown(oGrilla As SSUltraGrid, KeyCode As UltraGrid.SSReturnShort)
        If Not oGrilla.ActiveCell Is Nothing Then
            Select Case oGrilla.ActiveCell.Column.Key
            Case "Descripcion_Tipo_Item"
                Select Case KeyCode
                Case vbKeyBack
                Case vbKeyEscape
                    Set grillaitem.DataSource = Nothing
                    grillaitem.Visible = False
                Case vbKeyReturn
                Case vbKeyDown
                    On Error Resume Next
                    grillaitem.SetFocus
                Case vbKeyLeft
                End Select
            Case "Codigo"
                Select Case KeyCode
                Case vbKeyBack
                Case vbKeyEscape
                    Set grillacod.DataSource = Nothing
                    grillacod.Visible = False
                Case vbKeyReturn
                Case vbKeyDown
                    On Error Resume Next
                    grillacod.SetFocus
                Case vbKeyLeft
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
            Case "IdSubClasificacion"
                Select Case KeyCode
                Case vbKeyBack
                Case vbKeyEscape
                    Set grilladx.DataSource = Nothing
                    grilladx.Visible = False
                Case vbKeyReturn
                Case vbKeyDown
                    On Error Resume Next
                    grilladx.SetFocus
                Case vbKeyLeft
                End Select
            Case "LabConfHis"
                Select Case KeyCode
                Case vbKeyBack
                Case vbKeyEscape
                    Set grillahis.DataSource = Nothing
                    grillahis.Visible = False
                Case vbKeyReturn
                Case vbKeyDown
                    On Error Resume Next
                    grillahis.SetFocus
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

      

        If oGrilla.ActiveCell.Column.Key = "NombreProducto" Then
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
                Dim sTitem As String
                Select Case KeyAscii
                Case vbKeyBack
                    sNombre = oGrilla.ActiveCell.GetText
                    sTitem = grdProductos.ActiveRow.Cells("Descripcion_Tipo_Item").Value
                Case Else
                    sNombre = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
                    sTitem = grdProductos.ActiveRow.Cells("Descripcion_Tipo_Item").Value
                End Select
                Dim rs As New Recordset
                Dim oConexion As New ADODB.Connection
                Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                Set oFarmMovimientoDetalle.Conexion = oConexion
                   Set rs = mo_ReglasFarmacia.usp_listarcatalogoservicio(sNombre, sTitem)
                Set grillaBusqueda.DataSource = rs
                If mRs_Productos.RecordCount < 7 Then
                   grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY
                Else
                   grillaBusqueda.Top = 0
                End If
                grillaBusqueda.Visible = True
                grillaBusqueda.Enabled = True
                oConexion.Close
                Set oConexion = Nothing
                Set oFarmMovimientoDetalle = Nothing
            End Select
        End If
        
     
        If oGrilla.ActiveCell.Column.Key = "Codigo" Then
            Select Case KeyAscii
            Case vbKeyEscape
                If Trim(oGrilla.ActiveCell.GetText) = "" Then
                    grillacod.Visible = False
                    Set grillacod.DataSource = Nothing
                End If
            Case vbKeyReturn
            Case vbKeyDown
            Case vbKeyLeft
            Case Else
                Dim sCod As String
                Dim sTitem1 As String
                Select Case KeyAscii
                Case vbKeyBack
                    sCod = oGrilla.ActiveCell.GetText
                     sTitem1 = grdProductos.ActiveRow.Cells("Descripcion_Tipo_Item").Value
                Case Else
                    sCod = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
                      sTitem1 = grdProductos.ActiveRow.Cells("Descripcion_Tipo_Item").Value
                End Select
                Dim rs4 As New Recordset
                Dim oConexion4 As New ADODB.Connection
                Dim oFarmMovimientoDetalle4 As New farmMovimientoDetalle
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                Set oFarmMovimientoDetalle.Conexion = oConexion
                   Set rs = mo_ReglasFarmacia.usp_listarcatalogoservicioc(sCod, Trim(sTitem1))
                Set grillacod.DataSource = rs
                If mRs_Productos.RecordCount < 7 Then
                   grillacod.Top = oGrilla.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY
                Else
                   grillacod.Top = 0
                End If
                grillacod.Visible = True
                grillacod.Enabled = True
                oConexion.Close
                Set oConexion4 = Nothing
                Set oFarmMovimientoDetalle4 = Nothing
            End Select
        End If
        
   
        If oGrilla.ActiveCell.Column.Key = "IdSubClasificacion" Then
             Select Case KeyAscii
            Case vbKeyEscape
                If Trim(oGrilla.ActiveCell.GetText) = "" Then
                    grilladx.Visible = False
                    Set grilladx.DataSource = Nothing
                End If
            Case vbKeyReturn
            Case vbKeyDown
            Case vbKeyLeft
            Case Else
                Dim sDx As String
                Select Case KeyAscii
                Case vbKeyBack
                    sDx = oGrilla.ActiveCell.GetText
                Case Else
                    sDx = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
                End Select
                Dim rs1 As New Recordset
                Dim oConexion1 As New ADODB.Connection
                Dim oFarmMovimientoDetalle1 As New farmMovimientoDetalle
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                Set oFarmMovimientoDetalle1.Conexion = oConexion
                   Set rs1 = mo_ReglasFarmacia.usp_listarsubclasdx(sDx)
                Set grilladx.DataSource = rs1
                If mRs_Productos.RecordCount < 7 Then
                   grilladx.Top = oGrilla.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY
                Else
                   grilladx.Top = 0
                End If
                grilladx.Visible = True
                grilladx.Enabled = True
                oConexion.Close
                Set oConexion1 = Nothing
                Set oFarmMovimientoDetalle1 = Nothing
            End Select
           
        End If
        
        If oGrilla.ActiveCell.Column.Key = "LabConfHis" Then
             Select Case KeyAscii
            Case vbKeyEscape
                If Trim(oGrilla.ActiveCell.GetText) = "" Then
                    grillahis.Visible = False
                    Set grillahis.DataSource = Nothing
                End If
            Case vbKeyReturn
                Set grillahis.DataSource = Nothing
                 grillahis.Visible = False
            Case vbKeyDown
            Case vbKeyLeft
            Case Else
                Dim sHis As String
                Select Case KeyAscii
                Case vbKeyBack
                    sHis = oGrilla.ActiveCell.GetText
                Case Else
                    sHis = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
                End Select
                Dim rs2 As New Recordset
                Dim oConexion2 As New ADODB.Connection
                Dim oFarmMovimientoDetalle2 As New farmMovimientoDetalle
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                Set oFarmMovimientoDetalle2.Conexion = oConexion
                   Set rs2 = mo_ReglasFarmacia.usp_listarhissituacio(sHis)
                Set grillahis.DataSource = rs2
                If mRs_Productos.RecordCount < 7 Then
                   grillahis.Top = oGrilla.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY
                Else
                   grillahis.Top = 0
                End If
                grillahis.Visible = True
                grillahis.Enabled = True
                oConexion.Close
                Set oConexion2 = Nothing
                Set oFarmMovimientoDetalle2 = Nothing
            End Select
        End If
        
        If oGrilla.ActiveCell.Column.Key = "Descripcion_Tipo_Item" Then
             Select Case KeyAscii
            Case vbKeyEscape
                If Trim(oGrilla.ActiveCell.GetText) = "" Then
                    grillaitem.Visible = False
                    Set grillaitem.DataSource = Nothing
                End If
            Case vbKeyReturn
                Set grillaitem.DataSource = Nothing
                 grillaitem.Visible = False
            Case vbKeyDown
            Case vbKeyLeft
            Case Else
                Dim sItem As String
                Select Case KeyAscii
                Case vbKeyBack
                    sItem = oGrilla.ActiveCell.GetText
                Case Else
                    sItem = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
                End Select
                Dim rs3 As New Recordset
                Dim oConexion3 As New ADODB.Connection
                Dim oFarmMovimientoDetalle3 As New farmMovimientoDetalle
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                Set oFarmMovimientoDetalle3.Conexion = oConexion
                   Set rs3 = mo_ReglasFarmacia.usp_listartipoitem(sItem)
                Set grillaitem.DataSource = rs3
                If mRs_Productos.RecordCount < 7 Then
                   grillaitem.Top = oGrilla.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY
                Else
                   grillaitem.Top = 0
                End If
                grillaitem.Visible = True
                grillaitem.Enabled = True
                oConexion.Close
                Set oConexion3 = Nothing
                Set oFarmMovimientoDetalle3 = Nothing
            End Select
        End If

End Sub


Sub GenerarRecordsetProductos()
    With mRs_Productos
          .Fields.Append "Id", adInteger, 4
          .Fields.Append "IdAtencion", adVarChar, 20
          .Fields.Append "Descripcion_Tipo_Item", adVarChar, 100
          .Fields.Append "Fg_Tipo", adVarChar, 2
          .Fields.Append "Codigo", adVarChar, 10
          .Fields.Append "NombreProducto", adVarChar, 300
          .Fields.Append "IdSubClasificacion", adVarChar, 20
          .Fields.Append "LabConfHis", adVarChar, 5
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdProductos.DataSource = mRs_Productos
End Sub





Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
     On Error GoTo ConfigEstilo
     grdProductos.Bands(0).Columns("Id").Width = 500
    ' grdProductos.Bands(0).Columns("Id").Hidden = True
     grdProductos.Bands(0).Columns("Id").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("IdAtencion").Width = 800
      grdProductos.Bands(0).Columns("IdAtencion").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Descripcion_Tipo_Item").Header.Caption = "Descripción"
     grdProductos.Bands(0).Columns("Fg_Tipo").Header.Caption = "Fg_tipo"
     grdProductos.Bands(0).Columns("Fg_Tipo").Hidden = True
     grdProductos.Bands(0).Columns("Descripcion_Tipo_Item").Width = 1300
     grdProductos.Bands(0).Columns("Codigo").Width = 700
     grdProductos.Bands(0).Columns("NombreProducto").Width = 7000
     grdProductos.Bands(0).Columns("IdSubClasificacion").Width = 800
     grdProductos.Bands(0).Columns("LabConfHis").Width = 800
   
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
    
End Sub



Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, sighEntidades.GrillaConFilasBicolor
End Sub


Private Sub grilladx_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilladx grilladx
    gridInfra.ConfigurarFilasBiColores grilladx, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub grillahis_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillahis grillahis
    gridInfra.ConfigurarFilasBiColores grillahis, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub grillaitem_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaitem grillaitem
    gridInfra.ConfigurarFilasBiColores grillaitem, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub grillacod_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaCod grillacod
    gridInfra.ConfigurarFilasBiColores grillacod, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub grillaitem_DblClick()
Dim fila3 As New Record
    On Error GoTo ErrGrillaBusqueda
            grdProductos.ActiveRow.Cells("Descripcion_Tipo_Item").Value = grillaitem.ActiveRow.Cells("Descripcion_Tipo_Item").Value
            grdProductos.ActiveRow.Cells("Fg_Tipo").Value = grillaitem.ActiveRow.Cells("Fg_Tipo").Value
            
            Set grillaitem.DataSource = Nothing
            grillaitem.Visible = False
            'SendKeys "{Tab}"
ErrGrillaBusqueda:
End Sub



Private Sub InicializarLaGrillaBusqueda(oGrilla As SSUltraGrid)
    On Error GoTo errInic
    oGrilla.Bands(0).Columns("Codigo_Item").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo_Item").Width = 800
    oGrilla.Bands(0).Columns("Codigo_Item").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Descripcion_Item").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Descripcion_Item").Width = 7500
     oGrilla.Bands(0).Columns("Descripcion_Item").Activation = ssActivationActivateNoEdit
     
     oGrilla.Bands(0).Columns("Fg_Tipo").Header.Caption = "Fg_Tipo"
    oGrilla.Bands(0).Columns("Fg_Tipo").Width = 900
    oGrilla.Bands(0).Columns("Fg_Tipo").Activation = ssActivationActivateNoEdit
    
     oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Header.Caption = "Descripcion Item"
    oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Width = 900
    oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Activation = ssActivationActivateNoEdit
    
 
    
    oGrilla.Bands(0).Columns("Fg_Estado").Header.Caption = "Fg Estado"
    oGrilla.Bands(0).Columns("Fg_Estado").Width = 900
    oGrilla.Bands(0).Columns("Fg_Estado").Activation = ssActivationActivateNoEdit
     
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
errInic:
End Sub

Private Sub InicializarLaGrillaCod(oGrilla As SSUltraGrid)
    On Error GoTo errInic
    oGrilla.Bands(0).Columns("Codigo_Item").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo_Item").Width = 800
    oGrilla.Bands(0).Columns("Codigo_Item").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Descripcion_Item").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Descripcion_Item").Width = 7500
     oGrilla.Bands(0).Columns("Descripcion_Item").Activation = ssActivationActivateNoEdit
     
     oGrilla.Bands(0).Columns("Fg_Tipo").Header.Caption = "Fg_Tipo"
    oGrilla.Bands(0).Columns("Fg_Tipo").Width = 900
    oGrilla.Bands(0).Columns("Fg_Tipo").Activation = ssActivationActivateNoEdit
    
     oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Header.Caption = "Descripcion Item"
    oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Width = 900
    oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Fg_Estado").Header.Caption = "Fg Estado"
    oGrilla.Bands(0).Columns("Fg_Estado").Width = 900
    oGrilla.Bands(0).Columns("Fg_Estado").Activation = ssActivationActivateNoEdit
     
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
errInic:
End Sub


Private Sub InicializarLaGrilladx(oGrilla As SSUltraGrid)
    On Error GoTo errInic
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
   ' oGrilla.Bands(0).Columns("Codigo").Width = 800
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
   ' oGrilla.Bands(0).Columns("Descripcion").Width = 7500
     oGrilla.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit
     
  
     
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
errInic:
End Sub


Private Sub InicializarLaGrillahis(oGrilla As SSUltraGrid)
    On Error GoTo errInic
    
    oGrilla.Bands(0).Columns("Valores").Header.Caption = "Valores"
   ' oGrilla.Bands(0).Columns("Codigo").Width = 800
    oGrilla.Bands(0).Columns("Valores").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Descripcio").Header.Caption = "Descripción"
   ' oGrilla.Bands(0).Columns("Descripcion").Width = 7500
     oGrilla.Bands(0).Columns("Descripcio").Activation = ssActivationActivateNoEdit
 
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
errInic:
End Sub


Private Sub InicializarLaGrillaitem(oGrilla As SSUltraGrid)
    On Error GoTo errInic

    oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Width = 1600
     oGrilla.Bands(0).Columns("Descripcion_Tipo_Item").Activation = ssActivationActivateNoEdit
     
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
errInic:
End Sub

Private Sub grillaBusqueda_DblClick()
Dim fila As New Record
Dim codser As String
Dim desser As String
Dim lnIdProductoBusqueda As String
    On Error GoTo ErrGrillaBusqueda
    lnIdProductoBusqueda = grillaBusqueda.ActiveRow.Cells("Codigo_Item").Value
        If ItemYaExiste(lnIdProductoBusqueda) Then
           'grdProductos.ActiveRow.Cells("Id").Value = ""
            grdProductos.ActiveRow.Cells("IdAtencion").Value = ""
            grdProductos.ActiveRow.Cells("Codigo").Value = ""
            grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
            grdProductos.ActiveRow.Cells("IdSubClasificacion").Value = ""
            grdProductos.ActiveRow.Cells("LabConfHis").Value = ""
        Else
            RefrescarDatos
            Set grillaBusqueda.DataSource = Nothing
            grillaBusqueda.Visible = False
            SendKeys "{Tab}"
        End If
ErrGrillaBusqueda:
End Sub
Sub RefrescarDatos()
    Dim fila As New Record
    If Not grillaBusqueda.ActiveRow Is Nothing Then
       grdProductos.ActiveRow.Cells("IdAtencion").Value = lidatencionC
       grdProductos.ActiveRow.Cells("Codigo").Value = grillaBusqueda.ActiveRow.Cells("Codigo_Item").Value
       grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("Descripcion_Item").Value
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



Private Sub grillacod_DblClick()
Dim fila As New Record
Dim codser As String
Dim desser As String
Dim lnIdProductoBusqueda As String
    On Error GoTo ErrGrillaBusqueda
    lnIdProductoBusqueda = grillacod.ActiveRow.Cells("Descripcion_Item").Value
        If ItemYaExisteC(lnIdProductoBusqueda) Then
           ' grdProductos.ActiveRow.Cells("Id").Value = ""
            grdProductos.ActiveRow.Cells("IdAtencion").Value = ""
            grdProductos.ActiveRow.Cells("Codigo").Value = ""
            grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
            grdProductos.ActiveRow.Cells("IdSubClasificacion").Value = ""
            grdProductos.ActiveRow.Cells("LabConfHis").Value = ""
        Else
            RefrescarDatosC
            Set grillacod.DataSource = Nothing
            grillacod.Visible = False
            SendKeys "{Tab}"
            SendKeys "{Tab}"
        End If
ErrGrillaBusqueda:
End Sub
Sub RefrescarDatosC()
    Dim fila As New Record
    If Not grillacod.ActiveRow Is Nothing Then
       grdProductos.ActiveRow.Cells("IdAtencion").Value = lidatencionC
       grdProductos.ActiveRow.Cells("Codigo").Value = grillacod.ActiveRow.Cells("Codigo_Item").Value
       grdProductos.ActiveRow.Cells("NombreProducto").Value = grillacod.ActiveRow.Cells("Descripcion_Item").Value
    End If
End Sub

Private Sub grillaitem_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Set grillaitem.DataSource = Nothing
        grillaitem.Visible = False
    Case vbKeyReturn
        grillaitem_DblClick
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
    End Select
End Sub

Private Sub grillacod_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Set grillacod.DataSource = Nothing
        grillacod.Visible = False
    Case vbKeyReturn
        grillacod_DblClick
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
    End Select
End Sub

Private Sub grilladx_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Set grilladx.DataSource = Nothing
        grilladx.Visible = False
    Case vbKeyReturn
        grilladx_DblClick
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
    End Select
    
End Sub

Private Sub grillahis_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)

    Select Case KeyCode
    Case vbKeyEscape
        Set grillahis.DataSource = Nothing
        grillahis.Visible = False
    Case vbKeyReturn
        grillahis_DblClick
   
    
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
    End Select
    
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
        CargaProductosPorIdAtencion
End Sub

Property Get DevuelveProductos() As Recordset
    'Se debe utilizar un clon del recrdset, ya que si se trabaja directamente con el recordset
    'que esta asociado a la grilla ocurre errores en los metodos movenext, movefirst, etc.
    Set DevuelveProductos = mRs_Productos.Clone()
End Property
Property Get DevuelveTotal() As Double
    DevuelveTotal = dTotalIngresado
End Property



Private Sub grilladx_DblClick()
Dim fila1 As New Record
Dim lnIdProductoBusqueda As Long
    On Error GoTo ErrGrillaBusqueda
   
            'RefrescarDatos
            
            grdProductos.ActiveRow.Cells("IdSubClasificacion").Value = grilladx.ActiveRow.Cells("Codigo").Value
            
            Set grilladx.DataSource = Nothing
            grilladx.Visible = False
            SendKeys "{Tab}"
            SendKeys "{Tab}"
            SendKeys "{Tab}"
            SendKeys "{Tab}"
   
   
ErrGrillaBusqueda:
End Sub

Private Sub grillahis_DblClick()
Dim fila2 As New Record
Dim lnIdProductoBusqueda As Long
    On Error GoTo ErrGrillaBusqueda
   
            'RefrescarDatos
            
            grdProductos.ActiveRow.Cells("LabConfHis").Value = grillahis.ActiveRow.Cells("Valores").Value
            
            Set grillahis.DataSource = Nothing
            grillahis.Visible = False
            SendKeys "{Tab}"
            SendKeys "{Tab}"
            SendKeys "{Tab}"
            SendKeys "{Tab}"
   
   
ErrGrillaBusqueda:
End Sub

