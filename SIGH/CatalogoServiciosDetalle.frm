VERSION 5.00
Begin VB.Form CatalogoBaseServicioDetalle 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   Icon            =   "CatalogoServiciosDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   20
      Top             =   4020
      Width           =   5820
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoServiciosDetalle.frx":0CCA
         DownPicture     =   "CatalogoServiciosDetalle.frx":118E
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
         Picture         =   "CatalogoServiciosDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoServiciosDetalle.frx":1B66
         DownPicture     =   "CatalogoServiciosDetalle.frx":1FC6
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
         Left            =   1395
         Picture         =   "CatalogoServiciosDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   60
      TabIndex        =   16
      Top             =   0
      Width           =   5850
      Begin VB.ComboBox cmbIdClasificacionProducto 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3885
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   1
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox txtNombre 
         Height          =   345
         Left            =   1860
         MaxLength       =   250
         TabIndex        =   2
         Top             =   1020
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Clasificacion"
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
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         TabIndex        =   18
         Top             =   660
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         TabIndex        =   17
         Top             =   1020
         Width           =   1365
      End
   End
   Begin VB.Frame fraGrupoFarmacologico 
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   60
      TabIndex        =   13
      Top             =   1500
      Width           =   5850
      Begin VB.ComboBox cmbIdServicioGrupo 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   3885
      End
      Begin VB.ComboBox cmbIdServicioSubSeccion 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1020
         Width           =   3885
      End
      Begin VB.ComboBox cmbIdServicioSeccion 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   630
         Width           =   3885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo"
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Sub Sección"
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
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sección"
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
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1365
      End
   End
   Begin VB.Frame fraPresupuesto 
      Caption         =   "Presupuesto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   10
      Top             =   2940
      Width           =   5850
      Begin VB.ComboBox cmbIdPartida 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   630
         Width           =   3825
      End
      Begin VB.ComboBox cmbIdCentroCosto 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   210
         Width           =   3825
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Partida"
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
         TabIndex        =   12
         Top             =   660
         Width           =   1365
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
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
         TabIndex        =   11
         Top             =   270
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CatalogoBaseServicioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: PODiagnosticos
'        Autor: William Castro Grijalva
'        Fecha: 30/08/2004 12:17:18 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim mo_CatalogoServicios As New DOCatalogoServicios
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProducto As Long
Dim mo_AdminComun As New ReglasComunes

Dim mo_cmbIdCategoriaProducto As New SIGHComun.ListaDespleglable
Dim mo_cmbIdCentroCosto As New SIGHComun.ListaDespleglable
Dim mo_cmbIdPartida As New SIGHComun.ListaDespleglable
Dim mo_cmbIdServicioGrupo As New SIGHComun.ListaDespleglable
Dim mo_cmbIdServicioSeccion As New SIGHComun.ListaDespleglable
Dim mo_cmbIdServicioSubSeccion As New SIGHComun.ListaDespleglable


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

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
         
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         fraDatosGenerales.Enabled = False
         fraGrupoFarmacologico.Enabled = False
         fraPresupuesto.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         fraDatosGenerales.Enabled = False
         fraGrupoFarmacologico.Enabled = False
         fraPresupuesto.Enabled = False
         CargarDatosALosControles
 End Select
End Sub
Private Sub cmbIdCentroCosto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdCentroCosto
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdServicioGrupo_Click()
    'Recuperamos las secciones
    mo_cmbIdServicioSeccion.BoundColumn = "IdServicioSeccion"
    mo_cmbIdServicioSeccion.ListField = "Descripcion"
    Set mo_cmbIdServicioSeccion.RowSource = mo_AdminComun.CatalogoServiciosSeccionSeleccionarPorGrupo(Val(mo_cmbIdServicioGrupo.BoundText))
End Sub

Private Sub cmbIdServicioSeccion_Click()
    'Recuperamos los  SubGrupos
    mo_cmbIdServicioSubSeccion.BoundColumn = "IdServicioSubSeccion"
    mo_cmbIdServicioSubSeccion.ListField = "Descripcion"
    Set mo_cmbIdServicioSubSeccion.RowSource = mo_AdminComun.CatalogoServiciosSubSeccionSeleccionarPorSeccion(Val(mo_cmbIdServicioSeccion.BoundText))
End Sub
Private Sub cmbIdServicioGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioGrupo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdServicioSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioSeccion
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdPartida_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdPartida
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdServicioSubSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioSubSeccion
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdCategoriaProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdCategoriaProducto
    AdministrarKeyPreview KeyCode
End Sub


Private Sub Form_Initialize()
    Set mo_cmbIdCentroCosto.MiComboBox = cmbIdCentroCosto
    Set mo_cmbIdServicioGrupo.MiComboBox = cmbIdServicioGrupo
    Set mo_cmbIdServicioSeccion.MiComboBox = cmbIdServicioSeccion
    Set mo_cmbIdPartida.MiComboBox = cmbIdPartida
    Set mo_cmbIdServicioSubSeccion.MiComboBox = cmbIdServicioSubSeccion
    Set mo_cmbIdCategoriaProducto.MiComboBox = cmbIdCategoriaProducto
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Servicio"
       Case sghModificar
           Me.Caption = "Modificar Servicio"
       Case sghConsultar
           Me.Caption = "Consultar Servicio"
       Case sghEliminar
           Me.Caption = "Eliminar Servicio"
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
                   LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
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
   
   If mo_cmbIdCategoriaProducto.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese la categoria del producto" + Chr(13)
   End If
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If Trim(Me.txtNombre) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre" + Chr(13)
   End If
   If Trim(Me.txtPrecioUnitario) = "" Then
       sMensaje = sMensaje + "Ingrese el precio" + Chr(13)
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
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   Me.txtPrecioUnitario = Replace(Me.txtPrecioUnitario, ".", ",")
   With mo_CatalogoServicios
        .Codigo = Me.txtCodigo.Text
        .Nombre = Me.txtNombre.Text
        .PrecioUnitario = CCur(Me.txtPrecioUnitario.Text)
        .IdServicioGrupo = Val(mo_cmbIdServicioGrupo.BoundText)
        .IdCategoriaProducto = Val(mo_cmbIdCategoriaProducto.BoundText)
        .IdServicioSeccion = Val(mo_cmbIdServicioSeccion.BoundText)
        .IdServicioSubSeccion = Val(mo_cmbIdServicioSubSeccion.BoundText)
        .IdPartida = Val(mo_cmbIdPartida.BoundText)
        .IdCentroCosto = Val(mo_cmbIdCentroCosto.BoundText)
        
        .IdUsuarioAuditoria = Me.IdUsuario
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminComun.CatalogoServiciosAgregar(mo_CatalogoServicios)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminComun.CatalogoServiciosModificar(mo_CatalogoServicios)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminComun.CatalogoServiciosEliminar(mo_CatalogoServicios)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_CatalogoServicios = mo_AdminComun.CatalogoServiciosSeleccionarPorId(Me.IdProducto)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminComun.MensajeError, vbCritical, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CatalogoServicios Is Nothing Then
        With mo_CatalogoServicios
            mo_cmbIdCentroCosto.BoundText = .IdCentroCosto
            mo_cmbIdServicioGrupo.BoundText = .IdServicioGrupo
            mo_cmbIdServicioSeccion.BoundText = .IdServicioSeccion
            mo_cmbIdPartida.BoundText = .IdPartida
            mo_cmbIdServicioSubSeccion.BoundText = .IdServicioSubSeccion
            mo_cmbIdCategoriaProducto.BoundText = .IdCategoriaProducto
            
            Me.txtNombre = .Nombre
            Me.txtCodigo = .Codigo
            Me.txtPrecioUnitario = .PrecioUnitario
            
            mb_ExistenDatos = True
        End With
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

    Me.IdProducto = 0
    mo_cmbIdCentroCosto.BoundText = ""
    mo_cmbIdServicioGrupo.BoundText = ""
    mo_cmbIdServicioSeccion.BoundText = ""
    mo_cmbIdPartida.BoundText = ""
    mo_cmbIdServicioSubSeccion.BoundText = ""
    mo_cmbIdCategoriaProducto.BoundText = ""
    
    Me.txtNombre = ""
    Me.txtCodigo = ""
    
    Me.txtPrecioUnitario = ""
End Sub

Sub CargarComboBoxes()
       
    mo_cmbIdCategoriaProducto.BoundColumn = "IdCategoriaProducto"
    mo_cmbIdCategoriaProducto.ListField = "Descripcion"
    Set mo_cmbIdCategoriaProducto.RowSource = mo_AdminComun.CategoriasProductoSeleccionarTodos()

    mo_cmbIdCentroCosto.BoundColumn = "IdCentroCosto"
    mo_cmbIdCentroCosto.ListField = "Descripcion"
    Set mo_cmbIdCentroCosto.RowSource = mo_AdminComun.CentrosCostoSeleccionarTodos

    mo_cmbIdPartida.BoundColumn = "IdPartidaPresupuestal"
    mo_cmbIdPartida.ListField = "Descripcion"
    Set mo_cmbIdPartida.RowSource = mo_AdminComun.PartidasPresupuestalesSeleccionarTodos

    mo_cmbIdServicioGrupo.BoundColumn = "IdServicioGrupo"
    mo_cmbIdServicioGrupo.ListField = "Descripcion"
    Set mo_cmbIdServicioGrupo.RowSource = mo_AdminComun.CatalogoServiciosGrupoSeleccionarTodos
    

    
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtPrecioUnitario_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrecioUnitario
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtPrecioUnitario_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


