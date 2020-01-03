VERSION 5.00
Begin VB.Form InteoIntegracionSistema 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2730
   ClientLeft      =   10635
   ClientTop       =   6240
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   9180
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9000
      Begin VB.CheckBox chkEsProveedorActual 
         Caption         =   "Es Proveedor Actual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   10
         Top             =   1080
         Width           =   2625
      End
      Begin VB.ComboBox cmbIdProveedorSistema 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   7485
      End
      Begin VB.ComboBox cmbIdTipoSistema 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   7485
      End
      Begin VB.TextBox txtClaveUsuario 
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
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox txtNombreUsuario 
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
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label lblCodigoCIE2004 
         Caption         =   "Tipo Sistema"
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
         Left            =   135
         TabIndex        =   7
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   690
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   120
      TabIndex        =   0
      Top             =   1590
      Width           =   9000
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "InteoIntegracionSistema.frx":0000
         DownPicture     =   "InteoIntegracionSistema.frx":0460
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
         Left            =   2565
         Picture         =   "InteoIntegracionSistema.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "InteoIntegracionSistema.frx":0D4A
         DownPicture     =   "InteoIntegracionSistema.frx":120E
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
         Left            =   4080
         Picture         =   "InteoIntegracionSistema.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "InteoIntegracionSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de tipo de modalidad de sala
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_IntegracionSistema As New DOInteoIntegracionSistema
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdIntegracionSistema As Long
Dim mo_ReglasIntegracionSistema As New ReglasIntegracion
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mo_cmbIdTipoSistema As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdProveedorSistema As New SIGHEntidades.ListaDespleglable
Dim mb_RegistroInactivo As Boolean
Dim mb_ResultadoOperacion As Boolean

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
Property Let IdIntegracionSistema(lValue As Long)
   ml_IdIntegracionSistema = lValue
End Property
Property Get IdIntegracionSistema() As Long
   IdIntegracionSistema = ml_IdIntegracionSistema
End Property

Property Let ResultadoOperacion(bValue As Boolean)
   mb_ResultadoOperacion = bValue
End Property

Property Get ResultadoOperacion() As Boolean
   ResultadoOperacion = mb_ResultadoOperacion
End Property

Private Sub chkEsProveedorActual_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkEsProveedorActual
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdProveedorSistema_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdProveedorSistema
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoSistema_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSistema
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoSistema.MiComboBox = cmbIdTipoSistema
    Set mo_cmbIdProveedorSistema.MiComboBox = cmbIdProveedorSistema

End Sub


'SE creo estos campos para poder hacer uso de credenciales al brindar servicios a otros sistemas
Private Sub txtNombreUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombreUsuario
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNombreUsuario_KeyPress(KeyAscii As Integer)
    'IMPLEMENTAR REGLAS DE NOMBRE DE USUARIO
'    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
End Sub
'SE creo estos campos para poder hacer uso de credenciales al brindar servicios a otros sistemas
Private Sub txtClaveUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtClaveUsuario
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtClaveUsuario_KeyPress(KeyAscii As Integer)
'IMPLEMENTAR REGLAS DE SEGURIDADA EN CONTRASEÑAS
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
End Sub

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
         Frame1.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         Frame1.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
    CargaComboBoxes
       
    Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agregar Integración con tipo de sistema"
        Case sghModificar
            Me.Caption = "Modificar Integración con tipo de sistema"
        Case sghConsultar
            Me.Caption = "Consultar Integración con tipo de sistema"
        Case sghEliminar
            Me.Caption = "Eliminar Integración con tipo de sistema"
    End Select
    CargarDatosAlFormulario
    mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

Sub CargaComboBoxes()
    mo_cmbIdTipoSistema.BoundColumn = "IdTipoSistema"
    mo_cmbIdTipoSistema.ListField = "TipoSistema"
    Set mo_cmbIdTipoSistema.RowSource = mo_ReglasIntegracionSistema.TipoSistemaSeleccionarTodos()
        
    mo_cmbIdProveedorSistema.ListField = "ProveedorSistema"
    mo_cmbIdProveedorSistema.BoundColumn = "IdProveedorSistema"
    Set mo_cmbIdProveedorSistema.RowSource = mo_ReglasIntegracionSistema.ProveedorSistemaSeleccionarTodos()
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
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   mb_ResultadoOperacion = False
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            mb_RegistroInactivo = False
           If ValidarReglas() Then
                If mb_RegistroInactivo = False Then
                    mb_ResultadoOperacion = AgregarDatos()
                Else
                    mb_ResultadoOperacion = ModificarDatos()
                End If
               If mb_ResultadoOperacion Then
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
                   Me.cmbIdTipoSistema.SetFocus
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasIntegracionSistema.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   mb_ResultadoOperacion = True
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasIntegracionSistema.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   mb_ResultadoOperacion = True
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasIntegracionSistema.MensajeError, vbExclamation, Me.Caption
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
  
   If Val(mo_cmbIdTipoSistema.BoundText) = 0 Then
       sMensaje = sMensaje + "Seleccione Tipo de Sistema" + Chr(13)
   End If
   If Val(mo_cmbIdProveedorSistema.BoundText) = 0 Then
       sMensaje = sMensaje + "Seleccione Proveedor del sistema" + Chr(13)
   End If
'   If Me.txtNombreUsuario.Text = "" Then
'       sMensaje = sMensaje + "Ingrese el código " + Chr(13)
'   End If
'   If Me.txtClaveUsuario = "" Then
'       sMensaje = sMensaje + "Ingrese la descripción" + Chr(13)
'   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   'si existe y esta inactivo dejar que siga su curso
   CargaDatosAlObjetosDeDatos
   Dim oRs As ADODB.Recordset
   
    If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
        If mi_Opcion = sghAgregar Then
            mo_IntegracionSistema.IdIntegracionSistema = 0
        Else
            mo_IntegracionSistema.IdIntegracionSistema = ml_IdIntegracionSistema
        End If
        Set oRs = mo_ReglasIntegracionSistema.InteoIntegracionSistemaVerificarDuplicado(mo_IntegracionSistema)
         If oRs Is Nothing Then
             MsgBox mo_ReglasIntegracionSistema.MensajeError, vbInformation, "Intergración de Sistema"
             Exit Function
         Else
             Dim lNumeroError As Long
             lNumeroError = ValidarDuplicadoRegistro(mi_Opcion, oRs)
             If lNumeroError > 0 Then
                MsgBox mapearMensageErrorDupplicado(lNumeroError), vbInformation, "Intergración de Sistema"
                Exit Function
             End If
         End If
    End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   With mo_IntegracionSistema
        .IdTipoSistema = mo_cmbIdTipoSistema.BoundText
        .IdProveedorSistema = mo_cmbIdProveedorSistema.BoundText
        .EsProveedorActual = Me.chkEsProveedorActual.Value
        .NombreUsuario = Me.txtNombreUsuario.Text
        .ClaveUsuario = UCase(Me.txtClaveUsuario.Text)
        .IdUsuarioAuditoria = Me.idUsuario
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_ReglasIntegracionSistema.InteoIntegracionSistemaAgregar(mo_IntegracionSistema, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, cmbIdTipoSistema.Text)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_ReglasIntegracionSistema.InteoIntegracionSistemaModificar(mo_IntegracionSistema, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, cmbIdTipoSistema.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_ReglasIntegracionSistema.InteoIntegracionSistemaEliminar(mo_IntegracionSistema, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, cmbIdTipoSistema.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    Set mo_IntegracionSistema = mo_ReglasIntegracionSistema.InteoIntegracionSistemaSeleccionarPorId(Me.IdIntegracionSistema)
    If mo_ReglasIntegracionSistema.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_ReglasIntegracionSistema.MensajeError, vbInformation, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_IntegracionSistema Is Nothing Then
        With mo_IntegracionSistema
            mo_cmbIdTipoSistema.BoundText = .IdTipoSistema
            mo_cmbIdProveedorSistema.BoundText = .IdProveedorSistema
            Me.txtNombreUsuario = .NombreUsuario
            Me.txtClaveUsuario = .ClaveUsuario
            Me.chkEsProveedorActual.Value = IIf(.EsProveedorActual, 1, 0)
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
    Me.IdIntegracionSistema = 0
    Me.txtNombreUsuario = ""
    Me.txtClaveUsuario = ""
    mo_cmbIdTipoSistema.BoundText = ""
    mo_cmbIdProveedorSistema.BoundText = ""
    chkEsProveedorActual.Value = 0
    Err = 0
End Sub


Private Function ValidarDuplicadoRegistro(iOpcion As sghOpciones, oRs As ADODB.Recordset) As Integer
    Dim lTotalRegistros As Long
    ValidarDuplicadoRegistro = 0
    
    If oRs.RecordCount > 0 Then
        lTotalRegistros = oRs.RecordCount
        
        oRs.MoveFirst
        While oRs.EOF = False
            If oRs.Fields!EsActivo = True Then
                'duplicado activo
                ValidarDuplicadoRegistro = 1
                Exit Function
            Else
                If iOpcion = sghModificar Then
                    'registro duplicado e inactivo si se permite la edicion ocasionaria dos registros uno activo y otro inactivo
                    ValidarDuplicadoRegistro = 2
                    Exit Function
                ElseIf iOpcion = sghAgregar Then
                    'si hay mas de un registro inactivo duplicado no se puede permitir activar uno al azar
                    If lTotalRegistros > 1 Then
                        ValidarDuplicadoRegistro = 3
                        Exit Function
                    Else
'                        ml_IdIntegracionSistema = oRs.Fields!IdIntegracionSistema
                        mo_IntegracionSistema.IdIntegracionSistema = oRs.Fields!IdIntegracionSistema
                        mb_RegistroInactivo = True
                        Exit Function
                    End If
                End If
            End If
            oRs.MoveNext
        Wend
    End If
End Function

Private Function mapearMensageErrorDupplicado(lNumeroError As Long) As String
    Dim sMessage As String
    Select Case lNumeroError
        Case 1:
            sMessage = "Existe un registro para este tipo de sistema y proveedor"
        Case 2:
            sMessage = "Existe un registro para este tipo de sistema y proveedor que ha sido eliminado"
        Case 3:
            sMessage = "Existe mas de un registro para este tipo de sistema y proveedor que han sido eliminados"
    End Select
    
    mapearMensageErrorDupplicado = sMessage
End Function


