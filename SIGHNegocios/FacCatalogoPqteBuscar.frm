VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FacCatalogoPqteBuscar 
   Caption         =   "Busqueda Procedimientos"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   Icon            =   "FacCatalogoPqteBuscar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   50
      TabIndex        =   4
      Top             =   0
      Width           =   10650
      Begin VB.CheckBox chkFiltroIzq 
         Alignment       =   1  'Right Justify
         Caption         =   "Filtro desde la IZQUIERDA"
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
         Left            =   8055
         TabIndex        =   13
         Top             =   930
         Width           =   2505
      End
      Begin VB.CheckBox chkSaldosFmayoresAcero 
         Caption         =   "Sólo Muestra Medicamentos e Insumos con Saldos mayores a CERO"
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
         Left            =   8790
         TabIndex        =   11
         Top             =   195
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7170
         Picture         =   "FacCatalogoPqteBuscar.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8535
         Picture         =   "FacCatalogoPqteBuscar.frx":2C55
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   5955
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblFarmacia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Height          =   210
         Left            =   180
         TabIndex        =   12
         Top             =   900
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Left            =   150
         TabIndex        =   10
         Top             =   810
         Width           =   7635
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código       Nombre"
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
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   6825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   50
      TabIndex        =   0
      Top             =   7200
      Width           =   10665
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacCatalogoPqteBuscar.frx":5831
         DownPicture     =   "FacCatalogoPqteBuscar.frx":5CF5
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
         Left            =   5430
         Picture         =   "FacCatalogoPqteBuscar.frx":61E1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacCatalogoPqteBuscar.frx":66CD
         DownPicture     =   "FacCatalogoPqteBuscar.frx":6B2D
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
         Left            =   3900
         Picture         =   "FacCatalogoPqteBuscar.frx":6FA2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdPreVentaDet 
      Height          =   5775
      Left            =   50
      TabIndex        =   3
      Top             =   1350
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   10186
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
      Caption         =   ".."
   End
   Begin SIGHNegocios.ucFacturacionFarm ucFacturacionFarm1 
      Height          =   615
      Left            =   1455
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1085
   End
   Begin SIGHNegocios.ucFacturacionServ ucFacturacionServ1 
      Height          =   870
      Left            =   45
      TabIndex        =   15
      Top             =   75
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
   End
End
Attribute VB_Name = "FacCatalogoPqteBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Paquete para Farmacia o Caja
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim oRsPreVentaDet As New Recordset, oRsItemsMasivosElegidos As New Recordset
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim lcSql As String
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim oRs As New Recordset
Dim mi_idProducto As Long
Dim mi_idPuntoCarga As Long
Dim mi_Producto As String
Dim mi_Precio As Double
Dim mi_tipoProducto As Long
Dim mi_Codigo As String
Dim mi_Descripcion As String
Dim lnIdFarmaciaElegida As Long
Dim lcFarmaciaElegida As String
Dim ml_IdTipoFinanciamiento As Long
Dim ml_RegistraTodosLosItems As Boolean
Dim lnIdPuntoCargaOtrosCpt As Long
Property Let RegistraTodosLosItems(lValue As Boolean)
   ml_RegistraTodosLosItems = lValue
End Property

Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property

Property Let FarmaciaElegida(iValue As String)
  lcFarmaciaElegida = iValue
End Property
Property Let IdFarmaciaElegida(iValue As Long)
  lnIdFarmaciaElegida = iValue
End Property
Property Get DevuelveTodosLosItemsServ() As Recordset
    Set DevuelveTodosLosItemsServ = Me.ucFacturacionServ1.DevuelveProductos
End Property
Property Get DevuelveTodosLosItemsFarm() As Recordset
    Set DevuelveTodosLosItemsFarm = Me.ucFacturacionFarm1.DevuelveProductos
End Property

Property Get ItemsMasivosElegidos() As Recordset
    Dim oRow As SSRow
    Dim lnCantidad As Integer
    lnCantidad = 0
    grdPreVentaDet.Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    'grdPreVentaDet.SortFilter.AfterSortEnd
    If grdPreVentaDet.Bands(0).Columns.Exists("Agregar") = False Then
        Set ItemsMasivosElegidos = oRsItemsMasivosElegidos.Clone
        Exit Property
    End If
    grdPreVentaDet.Bands(0).Columns("Agregar").SortIndicator = ssSortIndicatorAscending
    Set oRow = Me.grdPreVentaDet.GetRow(ssChildRowFirst)
     
    If Not oRow Is Nothing Then
        If oRow.Cells.Count > 3 Then
            If oRow.Cells("Agregar").value = True Then
                oRsItemsMasivosElegidos.AddNew
                oRsItemsMasivosElegidos.Fields!idProducto = oRow.Cells("idProducto").value
                oRsItemsMasivosElegidos.Fields!Producto = oRow.Cells("Codigo").value & "//" & oRow.Cells("Nombre").value
                If mi_idPuntoCarga = 2600 Then  'OtrosCpts
                   oRsItemsMasivosElegidos.Fields!IdPuntoCarga = oRow.Cells("IdPuntoCarga").value
                End If
                oRsItemsMasivosElegidos.Update
                lnCantidad = lnCantidad + 1
            End If

            Do While oRow.HasNextSibling
                Set oRow = oRow.GetSibling(ssSiblingRowNext)
                If oRow.Cells("Agregar").value = True Then
                    oRsItemsMasivosElegidos.AddNew
                    oRsItemsMasivosElegidos.Fields!idProducto = oRow.Cells("idProducto").value
                    oRsItemsMasivosElegidos.Fields!Producto = oRow.Cells("Codigo").value & "//" & oRow.Cells("Nombre").value
                    If mi_idPuntoCarga = 2600 Then  'OtrosCpts
                       oRsItemsMasivosElegidos.Fields!IdPuntoCarga = oRow.Cells("IdPuntoCarga").value
                    End If
                    oRsItemsMasivosElegidos.Update
                    lnCantidad = lnCantidad + 1
                End If
            Loop
        End If
    End If
    If lnCantidad = 0 And mi_Producto <> "" Then
        oRsItemsMasivosElegidos.AddNew
        oRsItemsMasivosElegidos.Fields!idProducto = mi_idProducto
        oRsItemsMasivosElegidos.Fields!Producto = mi_Producto
        If mi_idPuntoCarga = 2600 Then  'OtrosCpts
           oRsItemsMasivosElegidos.Fields!IdPuntoCarga = lnIdPuntoCargaOtrosCpt
        End If
        oRsItemsMasivosElegidos.Update

    End If
    Set oRow = Nothing
    Set ItemsMasivosElegidos = oRsItemsMasivosElegidos.Clone  'grdPreVentaDet.DataSource
End Sub

Property Get Descripcion() As String
  Descripcion = mi_Descripcion
End Property
Property Get Codigo() As String
  Codigo = mi_Codigo
End Property

Property Get Precio() As String
  Precio = mi_Precio
End Property


Property Get Producto() As String
  Producto = mi_Producto
End Property


Property Let IdPuntoCarga(iValue As Long)
  mi_idPuntoCarga = iValue
End Property

Property Get IdPuntoCarga() As Long
  IdPuntoCarga = mi_idPuntoCarga
End Property


Property Let BotonPresionado(iValue As sghBotonDetallePresionado)
  mi_BotonPresionado = iValue
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
  BotonPresionado = mi_BotonPresionado
End Property

Property Let idProducto(iValue As Long)
  mi_idProducto = iValue
End Property

Property Get idProducto() As Long
  idProducto = mi_idProducto
End Property
Property Get TipoProducto() As Long
  TipoProducto = mi_tipoProducto
End Property




Private Sub btnAceptar_Click()
    On Error Resume Next
    If Me.ucFacturacionFarm1.Visible = True Or Me.ucFacturacionServ1.Visible = True Then
        
    Else
        Select Case mi_idPuntoCarga
        Case 0
           mi_Producto = oRsPreVentaDet.Fields!nombre
        Case Else
           mi_Producto = oRsPreVentaDet.Fields!Codigo & "//" & oRsPreVentaDet.Fields!nombre
           oRsPreVentaDet.Fields!Agregar = True
           oRsPreVentaDet.Update
           'Set oRsItemsMasivosElegidos = grdProductos.DataSource
           'Set oRs = Me.grdPreVentaDet.DataSource
           'oRs.Filter = "agregar=false"
        End Select
    End If
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnBuscar_Click()
    If Me.txtCodigo.Text = "" And Me.txtDescripcion.Text = "" Then
       MsgBox "Ingrese Código o Nombre", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lcFiltro As String
    If Me.txtCodigo.Text <> "" Then
       'Por codigo
       On Error Resume Next
       oRsPreVentaDet.MoveFirst
       oRsPreVentaDet.Find "codigo='" & Trim(Me.txtCodigo.Text) & "'"
    Else
       'Por Nombre
        If mi_idPuntoCarga = sghPtoCargaFarmacia Then
           If Me.chkSaldosFmayoresAcero.value = 1 Then
               If chkFiltroIzq.value = 1 Then
                  lcFiltro = "left(nombre," & Trim(Str(Len(Me.txtDescripcion.Text))) & ")='" & Me.txtDescripcion.Text & "'"
               Else
                  lcFiltro = "Nombre like '%" & Trim(Me.txtDescripcion.Text) & "%'"
               End If
               'debb-23/11/2016
               If lnIdFarmaciaElegida > 1 Then
                  lcFiltro = lcFiltro & " and dbo.farmAlmacen.idAlmacen=" & Trim(Str(lnIdFarmaciaElegida))
               End If
               Set oRsPreVentaDet = mo_ReglasFarmacia.farmSaldoTotalesSoloMayoresAcero(lcFiltro, sghPorDescripcion)
           Else
               If chkFiltroIzq.value = 1 Then
                  lcFiltro = "left(nombre," & Trim(Str(Len(Me.txtDescripcion.Text))) & ")='" & Me.txtDescripcion.Text & "'"
               Else
                  lcFiltro = "Nombre like '%" & Trim(Me.txtDescripcion.Text) & "%'"
               End If
               'debb-23/11/2016
               If lnIdFarmaciaElegida > 1 Then
                  lcFiltro = lcFiltro & " and dbo.farmAlmacen.idAlmacen=" & Trim(Str(lnIdFarmaciaElegida))
               End If
               Set oRsPreVentaDet = mo_reglasComunes.CatalogoBienesInsumosResumenSeleccionarPorFiltro(lcFiltro, sghPorDescripcion)
           End If
           Set Me.grdPreVentaDet.DataSource = oRsPreVentaDet
        ElseIf mi_idPuntoCarga = 2501 Then
           'Farmacia - solo SIS
           If chkFiltroIzq.value = 1 Then
               lcFiltro = "Nombre like '" & Trim(Me.txtDescripcion.Text) & "%'"
           Else
               lcFiltro = "Nombre like '%" & Trim(Me.txtDescripcion.Text) & "%'"
           End If
           oRsPreVentaDet.Filter = lcFiltro
        Else
           If chkFiltroIzq.value = 1 Then
               lcFiltro = "idestado=1 and Nombre like '" & Trim(Me.txtDescripcion.Text) & "%'"
           Else
               lcFiltro = "idestado=1 and Nombre like '%" & Trim(Me.txtDescripcion.Text) & "%'"
           End If
           oRsPreVentaDet.Filter = lcFiltro
        End If
    End If
    Me.grdPreVentaDet.SetFocus
End Sub


Private Sub btnCancelar_Click()
        mi_BotonPresionado = sghCancelar
        Me.Visible = False
End Sub

Private Sub btnLimpiar_Click()
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
End Sub

Private Sub chkSaldosFmayoresAcero_Click()
    Me.MousePointer = 11
    If Me.chkSaldosFmayoresAcero.value = 1 Then
       Set oRsPreVentaDet = mo_ReglasFarmacia.farmSaldoTotalesSoloMayoresAcero("", sghPorDescripcion)
    Else
       Set oRsPreVentaDet = mo_reglasComunes.CatalogoBienesInsumosResumenSeleccionarPorFiltro("", sghPorDescripcion)
    End If
    Set Me.grdPreVentaDet.DataSource = oRsPreVentaDet
    On Error Resume Next
    Me.txtDescripcion.SetFocus
    Me.MousePointer = 1
End Sub



Private Sub Form_Activate()
    On Error Resume Next
    If ucFacturacionServ1.Visible = True Then
       ucFacturacionServ1.TabEnDescripcion
    ElseIf ucFacturacionFarm1.Visible = True Then
       ucFacturacionFarm1.TabEnDescripcion
    Else
        txtDescripcion.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If oRsItemsMasivosElegidos.State = 1 Then Set oRsItemsMasivosElegidos = Nothing
    With oRsItemsMasivosElegidos
          .Fields.Append "idProducto", adInteger
          .Fields.Append "Producto", adVarChar, 300
          .Fields.Append "idPuntoCarga", adInteger, 0, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    chkSaldosFmayoresAcero.Visible = False
    Select Case mi_idPuntoCarga
    Case 0      'tabla especialidades
        Set oRsPreVentaDet = mo_reglasComunes.EspecialidadesSeleccionarPorFiltro("")
    Case sghPtoCargaFarmacia
        Me.Caption = "Buscar Medicamentos e Insumos"    'debb2014b
        'chkSaldosFmayoresAcero_Click
        'chkSaldosFmayoresAcero.Visible = True
        '
        lblFarmacia.Visible = True
        lblFarmacia.Caption = lcFarmaciaElegida
        Select Case lnIdFarmaciaElegida
        Case 0      'Todos los ITEMS
             Set oRsPreVentaDet = mo_reglasComunes.CatalogoBienesInsumosResumenSeleccionarPorFiltro("", sghPorDescripcion)
             Me.chkSaldosFmayoresAcero.value = 0    'debb-18/08/2016
        Case 1      'Los q tienen saldos >0
             Set oRsPreVentaDet = mo_ReglasFarmacia.farmSaldoTotalesSoloMayoresAcero("", sghPorDescripcion)
             
        Case Else   'eligiò alguna FARMACIA
             Set oRsPreVentaDet = mo_ReglasFarmacia.farmSaldoTotalesSoloMayoresAcero("dbo.farmAlmacen.idAlmacen=" & Trim(Str(lnIdFarmaciaElegida)), sghPorDescripcion)
        End Select
        '
        
    Case 1500   'Procedimientos Administrativos
        Set oRsPreVentaDet = mo_reglasComunes.CatalogoServiciosSeleccionarSoloAdministrativos
        oRsPreVentaDet.Filter = "idEstado=1"
    Case 2500   'Todos los Procedimientos (Imagenes/Laboratorio/Operaciones) - solo SIS
        Set oRsPreVentaDet = mo_reglasComunes.CatalogoServiciosSeleccionarSoloConPreciosEnSIS
        oRsPreVentaDet.Filter = "idEstado=1"
    Case 2501   'Medicamentos/Insumos - solo SIS
        Me.Caption = "Buscar Medicamentos e Insumos"    'debb2014b
        Set oRsPreVentaDet = mo_reglasComunes.CatalogoBienesInsumosSeleccionarSoloConPreciosEnSIS
    Case Else
        '*******/Laboratorio/imagenes/OtrosServicios(2600)/*****
         
        'If ml_IdTipoFinanciamiento > 0 Then
        '   Set oRsPreVentaDet = mo_ReglasComunes.CatalogoServiciosSeleccionarSoloPtoCargaYtipoFinanciamiento(mi_idPuntoCarga, ml_IdTipoFinanciamiento)
        'Else
           Set oRsPreVentaDet = mo_reglasComunes.CatalogoServiciosSeleccionarSoloConPreciosEnParticular(mi_idPuntoCarga)
        'End If
        oRsPreVentaDet.Filter = "idEstado=1"
    End Select
    Set Me.grdPreVentaDet.DataSource = oRsPreVentaDet
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPreVentaDet, sighentidades.GrillaConFilasBicolor
    'kike 2017
    ucFacturacionFarm1.Visible = False
    ucFacturacionFarm1.inicializar
    ucFacturacionServ1.Visible = False
    ucFacturacionServ1.inicializar
    If ml_RegistraTodosLosItems = True Then
        fraBusqueda.Visible = False
        grdPreVentaDet.Visible = False
        
        If mi_idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaFarmacia Then
            
            ucFacturacionFarm1.Visible = True
            ucFacturacionFarm1.Top = 0
            ucFacturacionFarm1.Left = 0
            ucFacturacionFarm1.Width = fraBusqueda.Width
            ucFacturacionFarm1.Height = fraBusqueda.Height + grdPreVentaDet.Height
            ucFacturacionFarm1.movNumero = ""
            ucFacturacionFarm1.IdAlmacen = lnIdFarmaciaElegida
            'ucFacturacionFarm1.inicializar
            ucFacturacionFarm1.TipoPrecioParaNiNs = 3    'precio de venta
            ucFacturacionFarm1.movNumero = ""
            ucFacturacionFarm1.LimpiarGrilla
            ucFacturacionFarm1.CargaProductosPorMovNumero
            ucFacturacionFarm1.AgregaProducto True
        
        Else
            ucFacturacionServ1.Visible = True
            ucFacturacionServ1.Top = 0
            ucFacturacionServ1.Left = 0
            ucFacturacionServ1.Width = fraBusqueda.Width
            ucFacturacionServ1.Height = fraBusqueda.Height + grdPreVentaDet.Height
            
            ucFacturacionServ1.TipoProducto = sghServicio
            ucFacturacionServ1.idUsuario = sighentidades.Usuario
            'ucFacturacionServ1.inicializar
            Set ucFacturacionServ1.Atencion = Nothing
            ucFacturacionServ1.idCuentaAtencion = 0
            ucFacturacionServ1.IdTipoFinanciamiento = 1
            ucFacturacionServ1.IdPuntoCarga = 99
            ucFacturacionServ1.PermiteAgregarItems = True
            'ucFacturacionServ1.AgregaProducto
            ucFacturacionServ1.LimpiarGrilla
            ucFacturacionServ1.AgregaProducto
            'ucFacturacionServ1.TabEnDescripcion
            Select Case mi_idPuntoCarga
            Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2
                 ucFacturacionServ1.FiltraCpt = sghFiltraCpt.sghCptSoloLabAnatomiaP
            Case sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2
                 ucFacturacionServ1.FiltraCpt = sghFiltraCpt.sghCptSoloLabBancoS
            Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral
                 ucFacturacionServ1.FiltraCpt = sghFiltraCpt.sghCptSoloEcografiaGeneral
            Case sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica
                 ucFacturacionServ1.FiltraCpt = sghFiltraCpt.sghCptSoloEcografiaObstetrica
            Case sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica
                 ucFacturacionServ1.FiltraCpt = sghFiltraCpt.sghCptSoloLabPatologiaC
            Case sghPuntosCargaBasicos.sghPtoCargaRayosX
                 ucFacturacionServ1.FiltraCpt = sghFiltraCpt.sghCptSoloRayosX
            Case sghPuntosCargaBasicos.sghPtoCargaTomografia
                 ucFacturacionServ1.FiltraCpt = sghFiltraCpt.sghCptSoloTomografia
            End Select
        End If
    End If
End Sub


Private Sub grdPreVentaDet_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = oRsPreVentaDet.DataSource
    On Error Resume Next
    lnIdPuntoCargaOtrosCpt = 0
    Select Case mi_idPuntoCarga
    Case 0
       mi_idProducto = rsRecordset("idEspecialidad")
    Case 2500, 2501
       mi_idProducto = rsRecordset("idProducto")
       mi_Precio = rsRecordset("PrecioUnitario")
       mi_Codigo = rsRecordset("codigo")
       mi_Descripcion = rsRecordset("nombre")
       If mi_idPuntoCarga = 2501 Then
          mi_tipoProducto = rsRecordset("tipoProducto")
       End If
    Case Else
       mi_idProducto = rsRecordset("idProducto")
       If mi_idPuntoCarga = 2600 Then  'OtrosCpts
          lnIdPuntoCargaOtrosCpt = rsRecordset("idPuntoCarga")
       End If
    End Select
End Sub

Private Sub grdPreVentaDet_DblClick()
       mi_idProducto = 0
       On Error GoTo errDet
       Select Case mi_idPuntoCarga
       Case 0
          mi_idProducto = oRsPreVentaDet.Fields!IdEspecialidad
       Case 2500, 2501
          mi_idProducto = oRsPreVentaDet.Fields!idProducto
          mi_Precio = oRsPreVentaDet.Fields!PrecioUnitario
          mi_Codigo = oRsPreVentaDet.Fields!Codigo
          mi_Descripcion = oRsPreVentaDet.Fields!nombre
          If mi_idPuntoCarga = 2501 Then
               mi_tipoProducto = oRsPreVentaDet("tipoProducto")
          End If
       Case Else
          mi_idProducto = oRsPreVentaDet.Fields!idProducto
          If mi_idPuntoCarga = 2600 Then  'OtrosCpts
            lnIdPuntoCargaOtrosCpt = oRsPreVentaDet("idPuntoCarga")
          End If
       End Select
       btnAceptar_Click
errDet:
End Sub

Private Sub grdPreVentaDet_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    'Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    Select Case mi_idPuntoCarga
    Case 0
        grdPreVentaDet.Bands(0).Columns("idEspecialidad").Hidden = True
        grdPreVentaDet.Bands(0).Columns("idEspecialidad").Width = 800
        grdPreVentaDet.Bands(0).Columns("idEspecialidad").Activation = ssActivationActivateNoEdit
        grdPreVentaDet.Bands(0).Columns("Nombre").Width = 9000
        grdPreVentaDet.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
        grdPreVentaDet.Bands(0).Columns.Add "Agregar", "Agregar"
        grdPreVentaDet.Bands(0).Columns("Agregar").DataType = ssDataTypeBoolean
        grdPreVentaDet.Bands(0).Columns("Agregar").Header.Caption = "¿Agregar?"
        grdPreVentaDet.Bands(0).Columns("Agregar").Width = 800
        grdPreVentaDet.Bands(0).Columns("Agregar").Style = ssStyleCheckBox
    Case Else
        grdPreVentaDet.Bands(0).Columns("idProducto").Hidden = True
        grdPreVentaDet.Bands(0).Columns("Codigo").Width = 800
        grdPreVentaDet.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
        grdPreVentaDet.Bands(0).Columns("Nombre").Width = 8500
        grdPreVentaDet.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
        grdPreVentaDet.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
        If mi_idPuntoCarga = 2500 Or mi_idPuntoCarga = 2501 Then
            grdPreVentaDet.Bands(0).Columns("precioUnitario").Hidden = True
            If mi_idPuntoCarga = 2501 Then
               grdPreVentaDet.Bands(0).Columns("tipoProducto").Hidden = True
            End If
            grdPreVentaDet.Bands(0).Columns.Add "Agregar", "Agregar"
            grdPreVentaDet.Bands(0).Columns("Agregar").DataType = ssDataTypeBoolean
            grdPreVentaDet.Bands(0).Columns("Agregar").Header.Caption = "¿Agregar?"
            grdPreVentaDet.Bands(0).Columns("Agregar").Width = 800
            grdPreVentaDet.Bands(0).Columns("Agregar").Style = ssStyleCheckBox
        Else
            grdPreVentaDet.Bands(0).Columns.Add "Agregar", "Agregar"
            grdPreVentaDet.Bands(0).Columns("Agregar").DataType = ssDataTypeBoolean
            grdPreVentaDet.Bands(0).Columns("Agregar").Header.Caption = "¿Agregar?"
            grdPreVentaDet.Bands(0).Columns("Agregar").Width = 800
            grdPreVentaDet.Bands(0).Columns("Agregar").Style = ssStyleCheckBox
            grdPreVentaDet.Bands(0).Columns("Nombre").Width = 8400
            If sghPuntosCargaBasicos.sghPtoCargaFarmacia <> mi_idPuntoCarga Then
               grdPreVentaDet.Bands(0).Columns("idEstado").Hidden = True
            End If
        End If
    End Select
    On Error Resume Next
    grdPreVentaDet.Bands(0).Columns("idTipoLocales").Hidden = True
End Sub






Private Sub grdPreVentaDet_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
     If KeyAscii = 13 Then
        grdPreVentaDet_DblClick
     End If

End Sub



Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.txtCodigo.Text <> "" Then
       btnBuscar_Click
    End If
End Sub



Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.txtDescripcion.Text <> "" Then
       btnBuscar_Click
    End If

End Sub




Private Sub ucFacturacionFarm1_SePresionoTeclaEspecial(KeyCode As Integer)
    If KeyCode = vbKeyF2 Then
       btnAceptar_Click
    End If
End Sub


Private Sub ucFacturacionServ1_SePresionoTeclaEspecial(KeyCode As Integer)
    If KeyCode = vbKeyF2 Then
       btnAceptar_Click
    End If
End Sub

