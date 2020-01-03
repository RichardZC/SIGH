VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form ProductosDetalle 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   Icon            =   "ProductosDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Planes disponibles"
      Height          =   3225
      Left            =   60
      TabIndex        =   16
      Top             =   1830
      Width           =   7995
      Begin VB.CommandButton btnEliminar 
         DisabledPicture =   "ProductosDetalle.frx":08CA
         DownPicture     =   "ProductosDetalle.frx":0C55
         Height          =   315
         Left            =   6840
         Picture         =   "ProductosDetalle.frx":0FE8
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton btnAgregar 
         DisabledPicture =   "ProductosDetalle.frx":1379
         DownPicture     =   "ProductosDetalle.frx":1762
         Height          =   315
         Left            =   5775
         Picture         =   "ProductosDetalle.frx":1B6E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox txtPrecioPlan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1230
         TabIndex        =   6
         Top             =   600
         Width           =   1065
      End
      Begin UltraGrid.SSUltraGrid grdPlanProducto 
         Height          =   1935
         Left            =   180
         TabIndex        =   7
         Top             =   990
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   3413
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "Planes disponibles"
      End
      Begin MSDataListLib.DataCombo cmbIdPlan 
         Height          =   315
         Left            =   1245
         TabIndex        =   5
         Top             =   255
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lblIdFuenteFinanciamiento 
         Caption         =   "Precio"
         Height          =   315
         Left            =   210
         TabIndex        =   18
         Top             =   690
         Width           =   795
      End
      Begin VB.Label lblIdTipoFinanciamiento 
         Caption         =   "Plan"
         Height          =   315
         Left            =   210
         TabIndex        =   17
         Top             =   330
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   15
      Top             =   5070
      Width           =   7995
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   2640
         Picture         =   "ProductosDetalle.frx":1F7A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   4215
         Picture         =   "ProductosDetalle.frx":23EF
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   60
      TabIndex        =   10
      Top             =   90
      Width           =   7995
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   0
         Top             =   210
         Width           =   1000
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   1230
         MaxLength       =   250
         TabIndex        =   1
         Top             =   570
         Width           =   6615
      End
      Begin VB.TextBox txtPrecioBase 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1245
         TabIndex        =   2
         Top             =   1260
         Width           =   1000
      End
      Begin VB.CheckBox chkBloqueado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Bloqueado"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6690
         TabIndex        =   4
         Top             =   1260
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo cmbIdCategoriaProducto 
         Height          =   315
         Left            =   3870
         TabIndex        =   3
         Top             =   1260
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código"
         Height          =   315
         Left            =   225
         TabIndex        =   14
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
         Height          =   315
         Left            =   225
         TabIndex        =   13
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblPrecioBase 
         Caption         =   "Precio base"
         Height          =   315
         Left            =   210
         TabIndex        =   12
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label lblIdCategoriaProducto 
         Caption         =   "Categoría"
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Top             =   1320
         Width           =   1305
      End
   End
End
Attribute VB_Name = "ProductosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POProductos
'        Autor: William Castro Grijalva
'        Fecha: 30/08/2004 08:02:52 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim mo_Productos As New DOProducto
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim ml_IdProducto As Long
Dim mo_PlanProducto As New Collection
Dim mrs_PlanProducto As New Recordset

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

       cmbIdCategoriaProducto.BoundColumn = "IdCategoriaProducto"
       cmbIdCategoriaProducto.ListField = "DescripcionLarga"
       'Set cmbIdCategoriaProducto.RowSource = mo_AdminFacturacion.CategoriasProductoSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminFacturacion.MensajeError
       
       cmbIdPlan.BoundColumn = "IdPlan"
       cmbIdPlan.ListField = "DescripcionLarga"
       Set cmbIdPlan.RowSource = mo_AdminFacturacion.PlanesSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminFacturacion.MensajeError
       
       If sMensaje <> "" Then
           MsgBox sMensaje, vbCritical, Me.Caption
       End If

End Sub
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property

Private Sub cmbIdCategoriaProducto_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCategoriaProducto
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdCategoriaProducto_LostFocus()
   If cmbIdCategoriaProducto.Text <> "" Then
       cmbIdCategoriaProducto.BoundText = Val(Split(cmbIdCategoriaProducto.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdCategoriaProducto
End Sub

Private Sub cmbIdCategoriaProducto_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub chkBloqueado_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, chkBloqueado
AdministrarKeyPreview KeyCode
End Sub

Private Sub chkBloqueado_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtPrecioBase_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrecioBase
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrecioBase_LostFocus()
   mo_Formulario.MarcarComoVacio txtPrecioBase
End Sub

Private Sub txtPrecioBase_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombre_LostFocus()
   mo_Formulario.MarcarComoVacio txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigo_LostFocus()
   mo_Formulario.MarcarComoVacio txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Productos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Productos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()

        GenerarRecordsetTemporal
        
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Productos"
       Case sghModificar
           Me.Caption = "Modificar Productos"
       Case sghConsultar
           Me.Caption = "Consultar Productos"
       Case sghEliminar
           Me.Caption = "Eliminar Productos"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Productos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   'If IdProducto = 0 Then
   '    sMensaje = sMensaje + "Ingrese el valor de IdProducto" + Chr(13)
   'End If
   
   If Me.cmbIdCategoriaProducto.BoundText = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de IdCategoriaProducto" + Chr(13)
   End If
   If Me.txtPrecioBase.Text = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de PrecioBase" + Chr(13)
   End If
   If Me.txtNombre.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de Nombre" + Chr(13)
   End If
   If Me.txtCodigo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de Codigo" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Productos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
Dim oDOPlanProducto As DOPlanProducto

   With mo_Productos
           .IdProducto = Me.IdProducto
           .IdCategoriaProducto = Me.cmbIdCategoriaProducto.BoundText
           .Bloqueado = Me.chkBloqueado.Value
           .PrecioBase = Me.txtPrecioBase.Text
           .Nombre = Me.txtNombre.Text
           .Codigo = Me.txtCodigo.Text
   End With
   
    'Busca IdPrestamo que se van a excluir
    Dim oRow As SSRow
    Set oRow = Me.grdPlanProducto.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        Set oDOPlanProducto = New DOPlanProducto
        oDOPlanProducto.IdPlanProducto = 0
        oDOPlanProducto.IdProducto = Me.IdProducto
        oDOPlanProducto.IdPlan = Val(oRow.Cells("IdPlan").Value)
        oDOPlanProducto.Precio = Val(oRow.Cells("Precio").Value)
        oDOPlanProducto.IdUsuarioAuditoria = ml_IdUsuario
        mo_PlanProducto.Add oDOPlanProducto
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            Set oDOPlanProducto = New DOPlanProducto
            oDOPlanProducto.IdPlanProducto = 0
            oDOPlanProducto.IdProducto = Me.IdProducto
            oDOPlanProducto.IdPlan = Val(oRow.Cells("IdPlan").Value)
            oDOPlanProducto.Precio = Val(oRow.Cells("Precio").Value)
            oDOPlanProducto.IdUsuarioAuditoria = ml_IdUsuario
            mo_PlanProducto.Add oDOPlanProducto
        Loop
    End If
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminFacturacion.ProductosAgregar(mo_Productos, mo_PlanProducto)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminFacturacion.ProductosModificar(mo_Productos, mo_PlanProducto)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminFacturacion.ProductosEliminar(mo_Productos)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Productos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

        Set mo_Productos = mo_AdminFacturacion.ProductosSeleccionarPorId(Me.IdProducto)
        If mo_AdminFacturacion.MensajeError <> "" Then
            MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbCritical, Me.Caption"
            mb_ExistenDatos = False
            Exit Sub
        End If
        
        If Not mo_Productos Is Nothing Then
           With mo_Productos
                Me.IdProducto = .IdProducto
                Me.cmbIdCategoriaProducto.BoundText = .IdCategoriaProducto
                Me.chkBloqueado.Value = IIf(.Bloqueado, 1, 0)
                Me.txtPrecioBase.Text = .PrecioBase
                Me.txtNombre.Text = .Nombre
                Me.txtCodigo.Text = .Codigo
                
                Dim rsPlanProducto As New Recordset
                Set rsPlanProducto = mo_AdminFacturacion.PlanesProductosSeleccionarPorIdProducto(Me.IdProducto)
                Do While Not rsPlanProducto.EOF
                     With mrs_PlanProducto
                         .AddNew
                         .Fields!IdPlan = rsPlanProducto!IdPlan
                         .Fields!NombrePlan = rsPlanProducto!NombrePlan
                         .Fields!Precio = "" & rsPlanProducto!Precio
                     End With
                     rsPlanProducto.MoveNext
                Loop
                
                mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Productos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.IdProducto = 0
           Me.cmbIdCategoriaProducto.BoundText = ""
           Me.chkBloqueado.Value = 0
           Me.txtPrecioBase.Text = ""
           Me.txtNombre.Text = ""
           Me.txtCodigo.Text = ""
   
End Sub

Sub GenerarRecordsetTemporal()
    
    With mrs_PlanProducto
          .Fields.Append "IdPlan", adVarChar, 10
          .Fields.Append "NombrePlan", adVarChar, 255
          .Fields.Append "Precio", adVarChar, 10
          .CursorType = adOpenStatic
          .LockType = adLockOptimistic
          .Open
    End With
    
    Set Me.grdPlanProducto.DataSource = mrs_PlanProducto
    
End Sub

Private Sub btnAgregar_Click()
    
    On Error Resume Next
    mrs_PlanProducto.MoveFirst
    Do While Not mrs_PlanProducto.EOF
        If Me.cmbIdPlan.BoundText = mrs_PlanProducto!IdPlan Then
            MsgBox "El plan seleccionado ya existe", vbExclamation, Me.Caption
            Exit Sub
        End If
        mrs_PlanProducto.MoveNext
    Loop
    
    With mrs_PlanProducto
        .AddNew
        .Fields!IdPlan = Me.cmbIdPlan.BoundText
        .Fields!NombrePlan = Me.cmbIdPlan.Text
        .Fields!Precio = Me.txtPrecioPlan.Text
    End With
    
End Sub

Private Sub btnQuitar_Click()
    On Error Resume Next
    With mrs_PlanProducto
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With

    Set Me.grdPlanProducto.DataSource = mrs_PlanProducto

End Sub


