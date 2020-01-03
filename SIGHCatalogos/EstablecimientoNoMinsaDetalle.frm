VERSION 5.00
Begin VB.Form EstablecimientoNoMinsaDetalle 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EstablecimientoNoMinsaDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1455
         MaxLength       =   150
         TabIndex        =   9
         Top             =   600
         Width           =   3900
      End
      Begin VB.ComboBox cmbIdTipoSubsector 
         Height          =   330
         Left            =   1455
         TabIndex        =   8
         Top             =   990
         Width           =   3900
      End
      Begin VB.ComboBox cmbIdDepartamento 
         Height          =   330
         ItemData        =   "EstablecimientoNoMinsaDetalle.frx":0CCA
         Left            =   1455
         List            =   "EstablecimientoNoMinsaDetalle.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1350
         Width           =   3900
      End
      Begin VB.ComboBox cmbIdProvincia 
         Height          =   330
         Left            =   1455
         TabIndex        =   6
         Top             =   1710
         Width           =   3885
      End
      Begin VB.ComboBox cmbIdDistrito 
         Height          =   330
         Left            =   1455
         TabIndex        =   5
         Top             =   2070
         Width           =   3885
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1455
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label lblAdvertencia 
         Caption         =   "No esta Habilitada la busqueda en la Web de Renaes"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   5175
      End
      Begin VB.Label lblIdTipoSubsector 
         Caption         =   "Subsector"
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Nombre"
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   615
         Width           =   1005
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   1770
         Width           =   705
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   12
         Top             =   2100
         Width           =   570
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   11
         Top             =   1395
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Código Renaes"
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   255
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   60
      TabIndex        =   2
      Top             =   2970
      Width           =   5535
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EstablecimientoNoMinsaDetalle.frx":0CCE
         DownPicture     =   "EstablecimientoNoMinsaDetalle.frx":1192
         Height          =   700
         Left            =   2910
         Picture         =   "EstablecimientoNoMinsaDetalle.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EstablecimientoNoMinsaDetalle.frx":1B6A
         DownPicture     =   "EstablecimientoNoMinsaDetalle.frx":1FCA
         Height          =   700
         Left            =   1365
         Picture         =   "EstablecimientoNoMinsaDetalle.frx":243F
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "EstablecimientoNoMinsaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Establecimientos NO MINSA
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_EstablecimientosNoMinsa As New DOEstablecimientoNoMinsa
Dim mo_AdminReglasCOmunes As New ReglasComunes
Dim mo_AdminReglasGeograficas As New ReglasServGeograf
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdEstablecimientoNoMinsa As Long
Dim mo_cmbIdTipoSubsector As New ListaDespleglable
Dim mo_cmbIdDepartamento As New ListaDespleglable
Dim mo_cmbIdProvincia As New ListaDespleglable
Dim mo_cmbIdDistrito As New ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ms_CodigoRenaesAnterior As String
Dim mb_BuequedaEnRenaesHabilitada As Boolean
Dim oBuscaEnSUNASA As New SIGHNegocios.SunasaConsumoWeb
Dim mb_CodigoValidado As Boolean
Dim mb_esEstablecimientoMinsa As Boolean

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

       mo_cmbIdTipoSubsector.BoundColumn = "IdTipoSubsector"
       mo_cmbIdTipoSubsector.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoSubsector.RowSource = mo_AdminReglasCOmunes.TiposSubsectorSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminReglasCOmunes.MensajeError
       If sMensaje <> "" Then
           MsgBox mo_AdminReglasCOmunes.MensajeError, vbInformation, Me.Caption
       End If
       
        mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
        mo_cmbIdDepartamento.ListField = "Nombre"
        Set mo_cmbIdDepartamento.RowSource = mo_AdminReglasGeograficas.DepartamentosSeleccionarTodos()
        mo_cmbIdDepartamento.BoundText = Trim(Str(Val(Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2))))
        sMensaje = sMensaje + mo_AdminReglasGeograficas.MensajeError


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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdEstablecimientoNoMinsa(lValue As Long)
   ml_IdEstablecimientoNoMinsa = lValue
End Property
Property Get IdEstablecimientoNoMinsa() As Long
   IdEstablecimientoNoMinsa = ml_IdEstablecimientoNoMinsa
End Property

Private Sub cmbIdDepartamento_Click()
       
       mo_cmbIdProvincia.BoundColumn = "IdProvincia"
       mo_cmbIdProvincia.ListField = "Nombre"
       On Error Resume Next
       Set mo_cmbIdProvincia.RowSource = mo_AdminReglasGeograficas.ProvinciasSeleccionarPorDepartamento(Val(mo_cmbIdDepartamento.BoundText))
       
       mo_cmbIdProvincia.BoundText = ""
       mo_cmbIdDistrito.BoundText = ""
End Sub
Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDepartamento_LostFocus()
   mo_Formulario.MarcarComoVacio cmbIdDepartamento
End Sub

Private Sub cmbIdDepartamento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDistrito
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDistrito_LostFocus()
   mo_Formulario.MarcarComoVacio cmbIdDistrito
End Sub

Private Sub cmbIdDistrito_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdProvincia
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdProvincia_LostFocus()
   
   mo_Formulario.MarcarComoVacio cmbIdProvincia
   
End Sub

Private Sub cmbIdProvincia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdProvincia_Click()
       
       mo_cmbIdDistrito.BoundColumn = "IdDistrito"
       mo_cmbIdDistrito.ListField = "Nombre"
       Set mo_cmbIdDistrito.RowSource = mo_AdminReglasGeograficas.DistritoSeleccionarPorProvincia(Val(mo_cmbIdProvincia.BoundText))

       If mo_AdminReglasGeograficas.MensajeError <> "" Then
            MsgBox mo_AdminReglasGeograficas.MensajeError, vbInformation, Me.Caption
       End If
       
       mo_cmbIdDistrito.BoundText = ""

End Sub

Private Sub cmbIdTipoSubsector_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSubsector
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoSubsector_LostFocus()
   If cmbIdTipoSubsector.Text <> "" Then
       mo_cmbIdTipoSubsector.BoundText = Val(Split(cmbIdTipoSubsector.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoSubsector
End Sub

Private Sub cmbIdTipoSubsector_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub






Private Sub Form_Initialize()
    Set mo_cmbIdTipoSubsector.MiComboBox = cmbIdTipoSubsector
    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdProvincia.MiComboBox = cmbIdProvincia
    Set mo_cmbIdDistrito.MiComboBox = cmbIdDistrito

End Sub



Private Sub txtCodigo_GotFocus()
    ms_CodigoRenaesAnterior = txtCodigo.Text
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> ms_CodigoRenaesAnterior And txtCodigo.Text <> "" Then
        Call BuscarEstablecimientoEnSIS(txtCodigo.Text)
    End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombre_LostFocus()
    txtNombre = mo_Teclado.CapitalizarNombres(txtNombre)
   mo_Formulario.MarcarComoVacio txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla EstablecimientosNoMinsa
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
    mb_CodigoValidado = True
    mb_esEstablecimientoMinsa = False
    Select Case mi_Opcion
        Case sghAgregar
            mb_CodigoValidado = False
        Case sghModificar
            CargarDatosALosControles
        Case sghConsultar
            CargarDatosALosControles
        Case sghEliminar
            CargarDatosALosControles
    End Select
    mb_BuequedaEnRenaesHabilitada = oBuscaEnSUNASA.HabilitadoParaBusquedaEnWebRenaes
    lblAdvertencia.Visible = Not mb_BuequedaEnRenaesHabilitada
 
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla EstablecimientosNoMinsa
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar EstablecimientosNoMinsa"
       Case sghModificar
           Me.Caption = "Modificar EstablecimientosNoMinsa"
       Case sghConsultar
           Me.Caption = "Consultar EstablecimientosNoMinsa"
           Me.fraDatos.Enabled = False
       Case sghEliminar
           Me.Caption = "Eliminar EstablecimientosNoMinsa"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla EstablecimientosNoMinsa
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
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminReglasCOmunes.MensajeError, vbExclamation, Me.Caption
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
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminReglasCOmunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminReglasCOmunes.MensajeError, vbExclamation, Me.Caption
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
   If Me.txtCodigo.Text = "" Then
      sMensaje = sMensaje + "Ingrese el valor de Código RENAES" + Chr(13)
   End If
   If Me.txtNombre.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de Nombre" + Chr(13)
   End If
   If Val(mo_cmbIdTipoSubsector.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de IdTipoSubsector" + Chr(13)
   End If
   If Val(mo_cmbIdDistrito.BoundText) = 0 Then
      sMensaje = sMensaje + "Elija el Distrito" + Chr(13)
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   
   If VerificarValidezCodigo() = False Then
        Exit Function
   End If
   If mb_esEstablecimientoMinsa = True Then
        MsgBox "Establecimiento Pertenece al MINSA, proceda a ingresarlo en la lista de establecimientos", vbInformation, Me.Caption
        Exit Function
   End If
   
    Dim oDOEstablecimiento As New DOEstablecimiento
    If mo_AdminReglasCOmunes.EstablecimientosSeleccionarPorCodigo(FormatoCodigoRENAES(Me.txtCodigo.Text, GALENHOS), oDOEstablecimiento) = True Then
        MsgBox "Código pertenece a establecimiento que ha sido registrado como perteneciente al MINSA", vbInformation, Me.Caption
        Exit Function
    End If
   
   Dim oRsTmp1 As New Recordset
   Set oRsTmp1 = mo_AdminReglasCOmunes.EstablecimientosNoMinsaSeleccionarPorCodigo(Me.txtCodigo.Text)
   Select Case mi_Opcion
   Case sghAgregar
        If oRsTmp1.RecordCount > 0 Then
           MsgBox "Ya existe ese CODIGO RENAES para: " & oRsTmp1.Fields!Nombre
           Exit Function
        End If
   Case sghModificar
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              If oRsTmp1.Fields!IdEstablecimientoNoMinsa <> mo_EstablecimientosNoMinsa.IdEstablecimientoNoMinsa Then
                 MsgBox "Ya existe ese CODIGO RENAES para: " & oRsTmp1.Fields!Nombre
                 Exit Function
              End If
              oRsTmp1.MoveNext
           Loop
        End If
   End Select
   
   Set oRsTmp1 = Nothing
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla EstablecimientosNoMinsa
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_EstablecimientosNoMinsa
           .IdEstablecimientoNoMinsa = Me.IdEstablecimientoNoMinsa
           .Nombre = Me.txtNombre.Text
           .IdTipoSubsector = mo_cmbIdTipoSubsector.BoundText
           .idDistrito = Val(mo_cmbIdDistrito.BoundText)
           .IdUsuarioAuditoria = ml_idUsuario
           .Codigo = Me.txtCodigo.Text
           'mgaray201503
           '.Codigo = formatocodigorenaes(Me.txtCodigo.Text, RENAESNORMA)
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Private Function VerificarValidezCodigo() As Boolean
    VerificarValidezCodigo = True
    If mb_CodigoValidado = False Then
        If MsgBox("Código ingresado no ha sido validado por ninguna fuente confiable, " & _
                        "desea agregar el registro de todas maneras", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            VerificarValidezCodigo = False
        End If
    End If
End Function

Function AgregarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    AgregarDatos = mo_AdminReglasCOmunes.EstablecimientosNoMinsaAgregar(mo_EstablecimientosNoMinsa, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminReglasCOmunes.EstablecimientosNoMinsaModificar(mo_EstablecimientosNoMinsa, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminReglasCOmunes.EstablecimientosNoMinsaEliminar(mo_EstablecimientosNoMinsa, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla EstablecimientosNoMinsa
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
Dim oDOProvincia As New DOProvincia

        Set mo_EstablecimientosNoMinsa = mo_AdminReglasCOmunes.EstablecimientosNoMinsaSeleccionarPorId(Me.IdEstablecimientoNoMinsa)
        If mo_AdminReglasCOmunes.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminReglasComunes.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        
       If Not mo_EstablecimientosNoMinsa Is Nothing Then
           With mo_EstablecimientosNoMinsa
                Me.IdEstablecimientoNoMinsa = .IdEstablecimientoNoMinsa
                Me.txtNombre.Text = .Nombre
                mo_cmbIdTipoSubsector.BoundText = .IdTipoSubsector
                Set oDOProvincia = mo_AdminReglasGeograficas.DistritoSeleccionarProvincia(.idDistrito)
                mo_cmbIdDepartamento.BoundText = oDOProvincia.IdDepartamento
                mo_cmbIdProvincia.BoundText = oDOProvincia.IdProvincia
                mo_cmbIdDistrito.BoundText = .idDistrito
                Me.txtCodigo.Text = .Codigo
                mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla EstablecimientosNoMinsa
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.IdEstablecimientoNoMinsa = 0
    Me.txtNombre.Text = ""
    mo_cmbIdTipoSubsector.BoundText = ""
    mo_cmbIdDepartamento.BoundText = ""
    mo_cmbIdProvincia.BoundText = ""
    mo_cmbIdDistrito.BoundText = ""
    Me.txtCodigo.Text = ""
End Sub

Private Function BuscarEstablecimientoEnSIS(ByVal cCodigo As String) As Recordset
    Dim oBuscaEnSIS As New SIGHNegocios.SisConsumoWeb
    Dim ODom_eess As Dom_eess
    Dim oRsTmp As Recordset
    Dim lcCodigoRenaes As String
    
    lcCodigoRenaes = sighEntidades.FormatoCodigoRENAES(cCodigo, SIS)
    Set oRsTmp = oBuscaEnSIS.m_eessEleccionarPorCodigoRenaes(lcCodigoRenaes)
    mb_esEstablecimientoMinsa = False
    If oRsTmp.RecordCount > 0 Then
        Set ODom_eess = New Dom_eess
        
        ODom_eess.pre_IdEESS = oRsTmp.Fields!pre_IdEESS
        ODom_eess.pre_CodigoRENAES = oRsTmp.Fields!pre_CodigoRENAES
        ODom_eess.pre_idCategoriaEESS = oRsTmp.Fields!pre_esmn
        'puede estar nulo
        ODom_eess.pre_IdDisa = oRsTmp.Fields!pre_IdDisa
        ODom_eess.pre_IdOdsis = oRsTmp.Fields!pre_IdOdsis
        
        ODom_eess.pre_IdEstado = oRsTmp.Fields!pre_IdEstado
        ODom_eess.pre_IdUbigeo = oRsTmp.Fields!pre_IdUbigeo
        ODom_eess.pre_Nombre = oRsTmp.Fields!pre_Nombre
        ODom_eess.Entidad = ""
        If MsgBox("Código RENAES encontrado en la base de datos SIS, desea ver su información", _
                        vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            Call AsignaDatosAControlesDesdeSIS(ODom_eess)
        End If
    Else
        If mb_BuequedaEnRenaesHabilitada = True Then
            Set ODom_eess = oBuscaEnSUNASA.ConsultarServicioBuscarEESSxCodigo(lcCodigoRenaes, Nothing)
            If Not (ODom_eess Is Nothing) Then
                If oBuscaEnSUNASA.EsEstablecimientoMinsa(ODom_eess) = True Then
                    mb_esEstablecimientoMinsa = True
                End If
                If MsgBox("Código RENAES encontrado en la WEB de RENAES, desea ver su información", _
                            vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                
                    Call AsignaDatosAControlesDesdeSIS(ODom_eess)
                
                End If
            End If
        End If
    End If
    'Estabecer valor para saber si el codigo a sido validado por alguna fuente externa
    mb_CodigoValidado = True
    If ODom_eess Is Nothing Then
        mb_CodigoValidado = False
    End If
End Function

Private Function AsignaDatosAControlesDesdeSIS(oDomEESs As Dom_eess) As Boolean
    Dim oDOProvincia As New DOProvincia
    txtNombre.Text = oDomEESs.pre_Nombre
    'mo_cmbIdTipoSubsector.BoundText = .IdTipoSubsector
    Set oDOProvincia = mo_AdminReglasGeograficas.DistritoSeleccionarProvincia(oDomEESs.pre_IdUbigeo)
    mo_cmbIdDepartamento.BoundText = oDOProvincia.IdDepartamento
    mo_cmbIdProvincia.BoundText = oDOProvincia.IdProvincia
    mo_cmbIdDistrito.BoundText = oDomEESs.pre_IdUbigeo
    If oDomEESs.Entidad <> "" Then
        MsgBox "Establecimiento pertenece a " & oDomEESs.Entidad, vbInformation, Me.Caption
    End If
End Function
