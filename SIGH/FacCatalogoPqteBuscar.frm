VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FacCatalogoPqteBuscar 
   Caption         =   "Busqueda Procedimientos"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   Icon            =   "FacCatalogoPqteBuscar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10785
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
      Height          =   945
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   10650
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
      Left            =   30
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
      Height          =   6135
      Left            =   30
      TabIndex        =   3
      Top             =   990
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   10821
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
End
Attribute VB_Name = "FacCatalogoPqteBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim oRsPreVentaDet As New Recordset, oRsItemsMasivosElegidos As New Recordset
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim lcSql As String
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mi_idProducto As Long
Dim mi_idPuntoCarga As Long
Dim mi_Producto As String
Dim mi_Precio As Double
Dim mi_tipoProducto As Long
Dim mi_Codigo As String
Dim mi_Descripcion As String

Property Get ItemsMasivosElegidos() As Recordset
    Dim oRow As SSRow
    Dim lnCantidad As Integer
    lnCantidad = 0
    Set oRow = Me.grdPreVentaDet.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        If oRow.Cells("Agregar").Value = True Then
            oRsItemsMasivosElegidos.AddNew
            oRsItemsMasivosElegidos.Fields!idProducto = oRow.Cells("idProducto").Value
            oRsItemsMasivosElegidos.Fields!Producto = oRow.Cells("Codigo").Value & "//" & oRow.Cells("Nombre").Value
            oRsItemsMasivosElegidos.Update
            lnCantidad = lnCantidad + 1
        End If
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            If oRow.Cells("Agregar").Value = True Then
                oRsItemsMasivosElegidos.AddNew
                oRsItemsMasivosElegidos.Fields!idProducto = oRow.Cells("idProducto").Value
                oRsItemsMasivosElegidos.Fields!Producto = oRow.Cells("Codigo").Value & "//" & oRow.Cells("Nombre").Value
                oRsItemsMasivosElegidos.Update
                lnCantidad = lnCantidad + 1
            End If
        Loop
    End If
    If lnCantidad = 0 And mi_Producto <> "" Then
        oRsItemsMasivosElegidos.AddNew
        oRsItemsMasivosElegidos.Fields!idProducto = mi_idProducto
        oRsItemsMasivosElegidos.Fields!Producto = mi_Producto
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


Property Let idPuntoCarga(iValue As Long)
  mi_idPuntoCarga = iValue
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
    Select Case mi_idPuntoCarga
    Case 0
       mi_Producto = oRsPreVentaDet.Fields!Nombre
    Case Else
       mi_Producto = oRsPreVentaDet.Fields!Codigo & "//" & oRsPreVentaDet.Fields!Nombre
       oRsPreVentaDet.Fields!agregar = True
       oRsPreVentaDet.Update
       'Set oRsItemsMasivosElegidos = grdProductos.DataSource
       
    End Select
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnBuscar_Click()
    If Me.txtCodigo.Text = "" And Me.txtDescripcion.Text = "" Then
       MsgBox "Ingrese Código o Nombre", vbCritical, Me.Caption
       Exit Sub
    End If
    If Me.txtCodigo.Text <> "" Then
       'Por codigo
       oRsPreVentaDet.MoveFirst
       oRsPreVentaDet.Find "codigo='" & Trim(Me.txtCodigo.Text) & "'"
    Else
       'Por Nombre
        oRsPreVentaDet.Filter = "Nombre like '%" & Trim(Me.txtDescripcion.Text) & "%'"
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

Private Sub Form_Activate()
    On Error Resume Next
    txtDescripcion.SetFocus
End Sub

Private Sub Form_Load()
    If oRsItemsMasivosElegidos.State = 1 Then Set oRsItemsMasivosElegidos = Nothing
    With oRsItemsMasivosElegidos
          .Fields.Append "idProducto", adInteger
          .Fields.Append "Producto", adVarChar, 300
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With

    Select Case mi_idPuntoCarga
    Case 0      'tabla especialidades
        Set oRsPreVentaDet = mo_ReglasComunes.EspecialidadesSeleccionarPorFiltro("")
    Case sghPtoCargaFarmacia
        Set oRsPreVentaDet = mo_ReglasComunes.CatalogoBienesInsumosSeleccionarPorFiltro("", sghPorDescripcion)
    Case 1500   'Procedimientos Administrativos
        Set oRsPreVentaDet = mo_ReglasComunes.CatalogoServiciosSeleccionarSoloAdministrativos
    Case 2500   'Todos los Procedimientos (Imagenes/Laboratorio/Operaciones) - solo SIS
        Set oRsPreVentaDet = mo_ReglasComunes.CatalogoServiciosSeleccionarSoloConPreciosEnSIS
    Case 2501   'Medicamentos/Insumos - solo SIS
        Set oRsPreVentaDet = mo_ReglasComunes.CatalogoBienesInsumosSeleccionarSoloConPreciosEnSIS
    Case Else
        Set oRsPreVentaDet = mo_ReglasComunes.CatalogoServiciosSeleccionarSoloConPreciosEnParticular(mi_idPuntoCarga)
    End Select
    Set Me.grdPreVentaDet.DataSource = oRsPreVentaDet
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPreVentaDet, SIGHEntidades.GrillaConFilasBicolor
End Sub


Private Sub grdPreVentaDet_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = oRsPreVentaDet.DataSource
    On Error Resume Next
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
          mi_Descripcion = oRsPreVentaDet.Fields!Nombre
          If mi_idPuntoCarga = 2501 Then
               mi_tipoProducto = oRsPreVentaDet("tipoProducto")
          End If
       Case Else
          mi_idProducto = oRsPreVentaDet.Fields!idProducto
       End Select
       btnAceptar_Click
errDet:
End Sub

Private Sub grdPreVentaDet_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Select Case mi_idPuntoCarga
    Case 0
        grdPreVentaDet.Bands(0).Columns("idEspecialidad").Hidden = True
        grdPreVentaDet.Bands(0).Columns("idEspecialidad").Width = 800
        grdPreVentaDet.Bands(0).Columns("idEspecialidad").Activation = ssActivationActivateNoEdit
        grdPreVentaDet.Bands(0).Columns("Nombre").Width = 9000
        grdPreVentaDet.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
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
        Else
            grdPreVentaDet.Bands(0).Columns.Add "Agregar", "Agregar"
            grdPreVentaDet.Bands(0).Columns("Agregar").DataType = ssDataTypeBoolean
            grdPreVentaDet.Bands(0).Columns("Agregar").Header.Caption = "¿Agregar?"
            grdPreVentaDet.Bands(0).Columns("Agregar").Width = 800
            grdPreVentaDet.Bands(0).Columns("Agregar").Style = ssStyleCheckBox
        End If
    End Select
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
