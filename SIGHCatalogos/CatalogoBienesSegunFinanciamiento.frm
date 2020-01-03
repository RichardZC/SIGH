VERSION 5.00
Begin VB.Form CatalogoBienesSegunFinanciamiento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10005
   Icon            =   "CatalogoBienesSegunFinanciamiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFarmacia 
      Caption         =   "Farmacia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   11
      Top             =   1260
      Width           =   9900
      Begin VB.CheckBox chkActualizaPV 
         Alignment       =   1  'Right Justify
         Caption         =   "Actualiza el Precio de Venta en los demás IAFA"
         Height          =   315
         Left            =   5730
         TabIndex        =   3
         Top             =   210
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.ComboBox cmbEstado 
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
         ItemData        =   "CatalogoBienesSegunFinanciamiento.frx":0CCA
         Left            =   2430
         List            =   "CatalogoBienesSegunFinanciamiento.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   570
         Width           =   1425
      End
      Begin VB.TextBox txtPrVenta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2430
         MaxLength       =   20
         TabIndex        =   2
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   13
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimo Precio de Venta"
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
         Left            =   150
         TabIndex        =   12
         Top             =   270
         Width           =   1890
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Caption         =   "Datos Generales"
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
      Height          =   1170
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   9900
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2460
         MaxLength       =   250
         TabIndex        =   1
         Top             =   660
         Width           =   7245
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   2460
         MaxLength       =   20
         TabIndex        =   0
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label9 
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
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label8 
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
         Left            =   180
         TabIndex        =   9
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   7
      Top             =   2430
      Width           =   9870
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoBienesSegunFinanciamiento.frx":0CEA
         DownPicture     =   "CatalogoBienesSegunFinanciamiento.frx":114A
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
         Left            =   3547
         Picture         =   "CatalogoBienesSegunFinanciamiento.frx":15BF
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoBienesSegunFinanciamiento.frx":1A34
         DownPicture     =   "CatalogoBienesSegunFinanciamiento.frx":1EF8
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
         Left            =   5092
         Picture         =   "CatalogoBienesSegunFinanciamiento.frx":23E4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CatalogoBienesSegunFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Muestra Precio Venta de Medamento e Insumo
'        Programado por: Castro W
'        Fecha: Agosto 2004
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_CatalogoBienesInsumos As New DOCatalogoBienesInsumos
Dim oCatalogoBienesPrecios As New FinanciamientoCatalogoBien
Dim oDoCatalogoBienesPrecios As New DoFinanciamientoCatalogoBien
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProducto As Long
Dim mo_AdminComun As New ReglasComunes
Dim mo_cmbIdClasificacionBienInsumo As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim ml_IdPlanCatalogo As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get idProducto() As Long
   idProducto = ml_IdProducto
End Property
Property Let IdPlanCatalogo(lValue As Long)
   ml_IdPlanCatalogo = lValue
End Property
Property Get IdPlanCatalogo() As Long
   IdPlanCatalogo = ml_IdPlanCatalogo
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
         CargarDatosALosControles
     Case sghEliminar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
 End Select
End Sub






Private Sub cmbEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbEstado
    AdministrarKeyPreview KeyCode

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Bien e Insumo"
       Case sghModificar
           Me.Caption = "Modificar Bien e Insumo"
       Case sghConsultar
           Me.Caption = "Consultar Bien e Insumo"
       Case sghEliminar
           Me.Caption = "Eliminar Bien e Insumo"
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub
Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
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
   
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If Trim(Me.txtNombre) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre" + Chr(13)
   End If
   If CDbl(txtPrVenta.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Precio" + Chr(13)
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
   'Me.txtPrecioUnitario = Replace(Me.txtPrecioUnitario, ".", ",")
   With oDoCatalogoBienesPrecios
        .PrecioUnitario = CDbl(txtPrVenta.Text)
        .Activo = cmbEstado.ListIndex
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminComun.CatalogoBienesInsumosAgregar(mo_CatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

    CargaDatosAlObjetosDeDatos
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oCatalogoBienesPrecios.Conexion = oConexion
    If oDoCatalogoBienesPrecios.Activo = False Then
        ModificarDatos = oCatalogoBienesPrecios.Eliminar(oDoCatalogoBienesPrecios)
    Else
        If chkActualizaPV.Value = 0 Then
               ModificarDatos = oCatalogoBienesPrecios.Modificar(oDoCatalogoBienesPrecios)
        Else
               mo_ReglasComunes.CatalogoBienesInsumosHospActualizaPrecioSegunIdProducto oDoCatalogoBienesPrecios.idProducto, oDoCatalogoBienesPrecios.PrecioUnitario
               ModificarDatos = True
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminComun.CatalogoBienesInsumosEliminar(mo_CatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
   Set mo_CatalogoBienesInsumos = mo_AdminComun.CatalogoBienesInsumosSeleccionarPorId(Me.idProducto)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos " + Chr(13) + mo_AdminComun.MensajeError, vbInformation, Me.Caption
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CatalogoBienesInsumos Is Nothing Then
        With mo_CatalogoBienesInsumos
            Me.txtNombre = .Nombre
            Me.txtCodigo = .Codigo
            mb_ExistenDatos = True
        End With
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
    
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oCatalogoBienesPrecios.Conexion = oConexion
    oDoCatalogoBienesPrecios.IdPlanCatalogo = ml_IdPlanCatalogo
    If oCatalogoBienesPrecios.SeleccionarPorId(oDoCatalogoBienesPrecios) Then
       txtPrVenta.Text = oDoCatalogoBienesPrecios.PrecioUnitario
       cmbEstado.ListIndex = IIf(oDoCatalogoBienesPrecios.Activo, 1, 0)
       mb_ExistenDatos = True
    End If
    oConexion.Close
    Set oConexion = Nothing
    
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.idProducto = 0
    
    Me.txtNombre = ""
    Me.txtCodigo = ""
End Sub

Sub CargarComboBoxes()
       
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode

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
