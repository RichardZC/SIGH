VERSION 5.00
Begin VB.Form FarmHistPreciosDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FarmHistPreciosDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1065
      Left            =   15
      TabIndex        =   11
      Top             =   2655
      Width           =   9660
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FarmHistPreciosDetalle.frx":0CCA
         DownPicture     =   "FarmHistPreciosDetalle.frx":118E
         Height          =   700
         Left            =   4957
         Picture         =   "FarmHistPreciosDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FarmHistPreciosDetalle.frx":1B66
         DownPicture     =   "FarmHistPreciosDetalle.frx":1FC6
         Height          =   700
         Left            =   3412
         Picture         =   "FarmHistPreciosDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Height          =   2595
      Left            =   15
      TabIndex        =   8
      Top             =   0
      Width           =   9675
      Begin VB.CommandButton cmbBuscaProducto 
         Caption         =   "..."
         Height          =   345
         Left            =   2925
         TabIndex        =   17
         ToolTipText     =   "Busca producto por Código o Descripción"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtPrCompra 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1005
         Width           =   1395
      End
      Begin VB.TextBox txtPrDistribucion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1395
         Width           =   1395
      End
      Begin VB.TextBox txtPrDonaciones 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2145
         Width           =   1395
      End
      Begin VB.TextBox txtPrVenta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1755
         Width           =   1395
      End
      Begin VB.CommandButton cmdCalculaPrec 
         Caption         =   "..."
         Height          =   345
         Left            =   2910
         TabIndex        =   16
         ToolTipText     =   "Calcula Precio de Distribución y Precio de Venta en base al precio de Compra"
         Top             =   1005
         Width           =   375
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   0
         Top             =   270
         Width           =   1395
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   6
         Top             =   630
         Width           =   8070
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pr Donación"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   15
         Top             =   2175
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pr Venta"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   14
         Top             =   1785
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pr Distribución"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   1410
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pr Compra"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   12
         Top             =   1035
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   660
         Width           =   645
      End
   End
End
Attribute VB_Name = "FarmHistPreciosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Almacenes
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_DoFarmHistPrecio As New DoFarmHistPrecio
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasFarmacia As New ReglasFarmacia
Dim mo_ReglasComunes As New ReglasComunes
Dim ms_MensajeError As String
Dim mb_ExistenDatos As Boolean
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdHistPrecio As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lnIdProducto As Long
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property


Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let IdHistPrecio(lValue As Long)
   ml_IdHistPrecio = lValue
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
     If mi_Opcion <> sghAgregar Then
        mo_Formulario.HabilitarDeshabilitar Me.txtCodigo, False
     End If
     mo_Formulario.HabilitarDeshabilitar Me.txtDescripcion, False
     Select Case mi_Opcion
     Case sghAgregar
         'cmbBuscaProducto.SetFocus
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
 End Select
End Sub



Private Sub cmbBuscaProducto_Click()
    Dim oBusqueda As New ListaProductos
    oBusqueda.MuestraTodosItems = True
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
        lnIdProducto = oBusqueda.IdRegistroSeleccionado
        Me.txtDescripcion.Text = oBusqueda.NombreSeleccionado
        txtCodigo.Text = oBusqueda.CodigoSeleccionado
    End If
    Set oBusqueda = Nothing
End Sub

Private Sub cmdCalculaPrec_Click()
    If Val(txtPrCompra.Text) > 0 Then
        txtPrDistribucion.Text = Round(CDbl(txtPrCompra.Text) + (CDbl(lcBuscaParametro.SeleccionaFilaParametro(307)) * CDbl(txtPrCompra.Text) / 100), 2)
        txtPrVenta.Text = Round(CDbl(txtPrCompra.Text) + ((CDbl(lcBuscaParametro.SeleccionaFilaParametro(307)) + CDbl(lcBuscaParametro.SeleccionaFilaParametro(307))) * CDbl(txtPrCompra.Text) / 100), 2)
    End If

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Precio"
       Case sghModificar
           Me.Caption = "Modificar Precio"
       Case sghConsultar
           Me.Caption = "Consultar Precio"
       Case sghEliminar
           Me.Caption = "Anular Precio"
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
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
                   MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   'LimpiarFormulario
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
'       If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
'           If ValidarReglas() Then
'               If EliminarDatos() Then
'                   MsgBox "Los datos se Anularon correctamente", vbInformation, Me.Caption
'                   Me.Visible = False
'                   LimpiarVariablesDeMemoria
'               Else
'                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
'               End If
'           End If
'        End If
       MsgBox "Solo se puede MODIFICAR"
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
   LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   sMensaje = ""
   If txtDescripcion.Text = "" Then
       sMensaje = sMensaje + "Tiene que elegir el Medicamento o Insumo" + Chr(13)
       Me.cmbBuscaProducto.SetFocus
   End If
   If CDbl(txtPrCompra.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Precio de la última Compra" + Chr(13)
       Me.txtPrCompra.SetFocus
   End If
   If Val(txtPrDistribucion.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Precio de Distribución" + Chr(13)
   End If
   If Val(txtPrVenta.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Precio de Venta" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   Dim sMensaje As String
   sMensaje = ""
   If Val(txtPrCompra.Text) > 0 And Val(txtPrDistribucion.Text) > 0 And Val(txtPrVenta.Text) > 0 Then
       If Not (CDbl(txtPrVenta.Text) > CDbl(txtPrDistribucion.Text) And CDbl(txtPrDistribucion.Text) > CDbl(txtPrCompra.Text)) Then
          sMensaje = sMensaje + "Se tiene que seguir el orden: Pr.Venta>Pr.Distribución>Pr.Compra" + Chr(13)
       End If
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   With mo_DoFarmHistPrecio
        If mi_Opcion = sghAgregar Then
           .fecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
           .idProducto = lnIdProducto
           .idUsuario = ml_idUsuario
        End If
        '.idHistPrecio
        .IdUsuarioAuditoria = ml_idUsuario
        .PrecioCompra = IIf(Me.txtPrCompra.Text = "", 0, CDbl(Me.txtPrCompra.Text))
        .PrecioDistribucion = IIf(Me.txtPrDistribucion.Text = "", 0, CDbl(Me.txtPrDistribucion.Text))
        .PrecioDonacion = IIf(Me.txtPrDonaciones.Text = "", 0, Val(Me.txtPrDonaciones.Text))
        .PrecioVenta = IIf(Me.txtPrVenta.Text = "", 0, CDbl(Me.txtPrVenta.Text))
   End With
End Sub



'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_ReglasFarmacia.FarmHistPrecioAgregar(mo_DoFarmHistPrecio, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_ReglasFarmacia.FarmHistPrecioModificar(mo_DoFarmHistPrecio, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_ReglasFarmacia.FarmHistPrecioEliminar(mo_DoFarmHistPrecio, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    Set mo_DoFarmHistPrecio = mo_ReglasFarmacia.farmHistPrecioSeleccionarPorId(ml_IdHistPrecio)
    If Not mo_DoFarmHistPrecio Is Nothing Then
         Dim oRsTmp1 As New Recordset
         Set oRsTmp1 = mo_ReglasComunes.CatalogoBienesInsumosSeleccionarXid(mo_DoFarmHistPrecio.idProducto)
         If oRsTmp1.RecordCount > 0 Then
            Me.txtCodigo = oRsTmp1!Codigo
            Me.txtDescripcion = oRsTmp1!Nombre
         End If
         oRsTmp1.Close
         Set oRsTmp1 = Nothing
         With mo_DoFarmHistPrecio
             lnIdProducto = .idProducto
             Me.txtPrCompra.Text = .PrecioCompra
             Me.txtPrDistribucion.Text = .PrecioDistribucion
             Me.txtPrDonaciones.Text = .PrecioDonacion
             Me.txtPrVenta.Text = .PrecioVenta
            
        End With
        mb_ExistenDatos = True
        If Format(lcBuscaParametro.RetornaFechaHoraServidorSQL, sighentidades.DevuelveFechaSoloFormato_DMY) > Format(mo_DoFarmHistPrecio.fecha, sighentidades.DevuelveFechaSoloFormato_DMY) Then
           MsgBox "No se podrá Modificar/Eliminar porque es de una Fecha pasada", vbInformation, "Precios"
           Me.btnAceptar.Visible = False
        End If
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
End Sub

Sub CargarComboBoxes()
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub




Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.txtCodigo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> "" And mi_Opcion = sghAgregar Then
        Dim rs As New ADODB.Recordset
        txtCodigo.Text = Trim(txtCodigo.Text)
        Set rs = mo_ReglasFarmacia.FactCatalogoBienesInsumosSeleccionarXDescripYcodigo(txtCodigo.Text, "")
        If rs.RecordCount > 0 Then
            lnIdProducto = rs.Fields("idproducto").Value
            Me.txtDescripcion.Text = rs.Fields("NombreProducto").Value
        Else
            Me.txtDescripcion.Text = ""
            txtCodigo.Text = ""
            lnIdProducto = 0
        End If
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_DoFarmHistPrecio = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasComunes = Nothing
End Sub

Private Sub txtPrCompra_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.txtPrCompra
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrCompra_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtPrDistribucion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrDistribucion
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtPrDistribucion_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtPrDonaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrDonaciones
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtPrDonaciones_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtPrVenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrVenta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtPrVenta_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub
