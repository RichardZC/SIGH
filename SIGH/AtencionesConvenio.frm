VERSION 5.00
Begin VB.Form AtencionesConvenio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13350
   Icon            =   "AtencionesConvenio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2250
      Left            =   75
      TabIndex        =   10
      Top             =   60
      Width           =   13215
      Begin VB.TextBox txtAnio 
         Height          =   300
         Left            =   2325
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1890
         Width           =   630
      End
      Begin VB.TextBox txtMes 
         Height          =   300
         Left            =   1875
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1890
         Width           =   375
      End
      Begin VB.TextBox txtDia 
         Height          =   300
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1890
         Width           =   375
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
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
         Left            =   1425
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1065
         Width           =   1000
      End
      Begin VB.TextBox txtCarta 
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
         Left            =   1410
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1485
         Width           =   1000
      End
      Begin VB.TextBox txtCodServicio 
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
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   1
         Text            =   "77781"
         Top             =   675
         Width           =   1000
      End
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   1440
         TabIndex        =   0
         Top             =   285
         Width           =   1000
      End
      Begin VB.Label lblNombreServicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2505
         TabIndex        =   17
         Top             =   690
         Width           =   10575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Oficio"
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
         Left            =   105
         TabIndex        =   16
         Top             =   1545
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe Sesión"
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
         Left            =   105
         TabIndex        =   15
         Top             =   1110
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Sesión"
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
         Left            =   105
         TabIndex        =   14
         Top             =   1890
         Width           =   1065
      End
      Begin VB.Label lblCodigoCIE2004 
         AutoSize        =   -1  'True
         Caption         =   "Nro Historia"
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
         Left            =   105
         TabIndex        =   13
         Top             =   330
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cod.Servicio"
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
         Left            =   105
         TabIndex        =   12
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lblNombrePaciente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2505
         TabIndex        =   11
         Top             =   285
         Width           =   10575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   75
      TabIndex        =   9
      Top             =   2355
      Width           =   13215
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AtencionesConvenio.frx":0CCA
         DownPicture     =   "AtencionesConvenio.frx":112A
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
         Left            =   5228
         Picture         =   "AtencionesConvenio.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AtencionesConvenio.frx":1A14
         DownPicture     =   "AtencionesConvenio.frx":1ED8
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
         Left            =   6758
         Picture         =   "AtencionesConvenio.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "AtencionesConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MZD 19/06/2005 [Todo el Archivo]
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase:
'        Autor: Daniel Barrantes
'        Fecha: 05/09/2007
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:Registro de SESIONES COBALTO TERAPIA para CONVENIO MINSA-ESSALUD
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim mo_AtencionesConvenio As New DOAtencionesConvenio
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdCajero As Long
Dim ml_idAtencionConvenio As Long
Dim ml_idPaciente As Long
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_IdProducto As Long
Dim sFecha As String
Const IdTipoFinanciamiento As Long = 4                'Convenios
Const lnCodServicioInicial As String = "77781"
Const lnCodServicioFinal As String = "77785"

Property Let idProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get idProducto() As Long
   idProducto = ml_IdProducto
End Property

Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let idAtencionConvenio(lValue As Long)
   ml_idAtencionConvenio = lValue
End Property
Property Get idAtencionConvenio() As Long
   idAtencionConvenio = ml_idAtencionConvenio
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

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Inicializar()

End Function
Sub CargarDatosAlFormulario()

    Select Case mi_Opcion
        Case sghAgregar
            txtDia.Text = Day(Date)
            txtMes.Text = Month(Date)
            txtAnio.Text = Year(Date)
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
    'GenerarRecordsetTemporal
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Atención Convenio"
       Case sghModificar
           Me.Caption = "Modificar Atención Convenio"
       Case sghConsultar
           Me.Caption = "Consultar Atención Convenio"
       Case sghEliminar
           Me.Caption = "Eliminar Atención Convenio"
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
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   txtDia.Text = ""
                   txtDia.SetFocus
               Else
                   MsgBox "No se pudo agregar los datos", vbExclamation, Me.Caption
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
                   MsgBox "No se pudo modificar los datos", vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos", vbExclamation, Me.Caption
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
   sMensaje = ""
   
   If lblNombrePaciente.Caption = "" Then
       sMensaje = sMensaje + "Ingrese Nro de Historia Clínica del Paciente " + Chr(13)
   End If
   If lblNombreServicio.Caption = "" Then
       sMensaje = sMensaje + "Sólo puede usar los Códigos de Servicios entre:  " + lnCodServicioInicial & " hasta " + lnCodServicioFinal + Chr(13)
   End If
   If txtCarta.Text = "" Then
       sMensaje = sMensaje + "Ingrese Nro de Carta " + Chr(13)
   End If
   If Val(txtImporte.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Importe "
   End If
   sFecha = Right("0" & txtDia.Text, 2) & "/" & Right("0" & txtMes.Text, 2) & "/" & Right(txtAnio.Text, 4)
   If Not SIGHComun.EsFecha(sFecha, "DD/MM/AAAA") Then
      sMensaje = sMensaje + "La Fecha registrada no es correcta"
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   Dim sMensaje As String
   Dim oRsBuscar As New ADODB.Recordset
   Dim oConexion As New ADODB.Connection
   ValidarReglas = False
   sMensaje = ""
   If Not (Val(txtCodServicio.Text) >= Val(lnCodServicioInicial) And Val(txtCodServicio.Text) <= Val(lnCodServicioFinal)) Then
       sMensaje = sMensaje + "Sólo puede usar los Códigos de Servicios entre:  " + lnCodServicioInicial + " hasta " + lnCodServicioFinal
   End If
   oConexion.Open SIGHComun.CadenaConexion
   oRsBuscar.Open "select * from atencionesConvenio where idPaciente=" & idPaciente & " and fechaSesion='" & sFecha & "' and idProducto=" & idProducto, oConexion, adOpenKeyset, adLockOptimistic
   Select Case mi_Opcion
   Case sghAgregar
        If oRsBuscar.RecordCount > 0 Then
           sMensaje = sMensaje + "Esa sesión Ya fue registrada (Paciente/Servicio/Fecha)"
        End If
   Case sghModificar
        If oRsBuscar.RecordCount > 0 Then
           oRsBuscar.MoveFirst
           Do While Not oRsBuscar.EOF
              If ml_idAtencionConvenio <> oRsBuscar.Fields!IdAtencionesConvenio And oRsBuscar.Fields!FechaSesion = sFecha Then
                 sMensaje = sMensaje + "Esa sesión Ya fue registrada (Paciente/Servicio/Fecha)"
                 Exit Do
              End If
              oRsBuscar.MoveNext
           Loop
        End If
   End Select
   oRsBuscar.Close
   
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
   With mo_AtencionesConvenio
        .NroOficio = txtCarta.Text
        .FechaSesion = sFecha
        .IdAtencionesConvenio = idAtencionConvenio
        .IdUsuarioAuditoria = ml_idUsuario
        .ImporteSesion = txtImporte.Text
        .idPaciente = idPaciente
        .NombrePaciente = lblNombrePaciente
        .NroHistoria = txtNroHistoria.Text
        .NombreProducto = lblNombreServicio.Caption
        .idProducto = idProducto
        .CodProducto = txtCodServicio.Text
   End With
   'Cargamos los detalles
   'CargarSupervisoresAlObjetoDatos mo_Supervisores
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oAtencionesConvenio As New SIGHDatos.AtencionesConvenio
    CargaDatosAlObjetosDeDatos
    AgregarDatos = False
    oConexion.Open SIGHComun.CadenaConexion
    Set oAtencionesConvenio.Conexion = oConexion
    If oAtencionesConvenio.Insertar(mo_AtencionesConvenio) Then
        AgregarDatos = True
    Else
        MsgBox oAtencionesConvenio.MensajeError
    End If
    oConexion.Close
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oAtencionesConvenio As New SIGHDatos.AtencionesConvenio
    CargaDatosAlObjetosDeDatos
    ModificarDatos = False
    oConexion.Open SIGHComun.CadenaConexion
    Set oAtencionesConvenio.Conexion = oConexion
    If oAtencionesConvenio.Modificar(mo_AtencionesConvenio) Then
        ModificarDatos = True
    Else
        MsgBox oAtencionesConvenio.MensajeError
    End If
    oConexion.Close
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oAtencionesConvenio As New SIGHDatos.AtencionesConvenio
    CargaDatosAlObjetosDeDatos
     EliminarDatos = False
    oConexion.Open SIGHComun.CadenaConexion
    Set oAtencionesConvenio.Conexion = oConexion
    If oAtencionesConvenio.Eliminar(mo_AtencionesConvenio) Then
        EliminarDatos = True
    Else
        MsgBox oAtencionesConvenio.MensajeError
    End If
    oConexion.Close
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    Dim oConexion As New ADODB.Connection
    Dim oAtencionesConvenio As New SIGHDatos.AtencionesConvenio
    'CargaDatosAlObjetosDeDatos
    mo_AtencionesConvenio.IdAtencionesConvenio = ml_idAtencionConvenio
    oConexion.Open SIGHComun.CadenaConexion
    Set oAtencionesConvenio.Conexion = oConexion
    If oAtencionesConvenio.SeleccionarPorId(mo_AtencionesConvenio) Then
       With mo_AtencionesConvenio
           txtDia.Text = Day(.FechaSesion)
           txtMes.Text = Month(.FechaSesion)
           txtAnio.Text = Year(.FechaSesion)
           sFecha = .FechaSesion
           idPaciente = .idPaciente
           txtImporte.Text = .ImporteSesion
           lblNombrePaciente.Caption = .NombrePaciente
           txtCarta.Text = .NroOficio
           lblNombreServicio.Caption = .NombreProducto
           idProducto = .idProducto
           txtCodServicio.Text = .CodProducto
           txtNroHistoria.Text = .NroHistoria
       End With
       mb_ExistenDatos = True
    Else
       mb_ExistenDatos = False
    End If
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
           txtDia.Text = Day(Date)
           txtMes.Text = Month(Date)
           txtAnio.Text = Year(Date)
           txtImporte.Text = 0
           idPaciente = 0
           txtNroHistoria.Text = ""
           lblNombrePaciente.Caption = ""
           txtCarta.Text = ""
           lblNombreServicio.Caption = ""
           txtCodServicio.Text = ""
           idProducto = 0
End Sub


Sub CargarComboBoxes()
       
End Sub





Private Sub txtCarta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCarta
   AdministrarKeyPreview KeyCode

End Sub


Private Sub txtCodServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodServicio
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodServicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtCodServicio.Text <> "" Then
        Dim oBuscaServicio As New SIGHDatos.CatalogoServicios
        Dim rsTmp As New ADODB.Recordset
        Set rsTmp = mo_ReglasFacturacion.FacturacionServicioPorCodigo(txtCodServicio.Text, IdTipoFinanciamiento)
        If rsTmp.RecordCount > 0 Then
           ml_IdProducto = rsTmp.Fields!idProducto
           lblNombreServicio.Caption = Trim(rsTmp.Fields!NombreProducto)
           txtImporte.Text = rsTmp.Fields!PrecioUnitario
        Else
           lblNombreServicio.Caption = ""
           ml_IdProducto = 0
           txtImporte.Text = 0
        End If
      End If
   End If
End Sub





Private Sub txtDia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys "{tab}"
   End If
End Sub

Private Sub txtImporte_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtImporte
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtNroHistoria.Text <> "" Then
        Dim oBuscaHistoria As New SIGHDatos.Pacientes
        Dim oConexion As New ADODB.Connection
        Dim oDOPaciente As New SIGHComun.doPaciente
        oDOPaciente.NroHistoriaClinica = txtNroHistoria.Text
        oConexion.Open SIGHComun.CadenaConexion
        Set oBuscaHistoria.Conexion = oConexion
        If oBuscaHistoria.SeleccionarPorHistoriaClinicaDefinitiva(oDOPaciente) Then
           ml_idPaciente = oDOPaciente.idPaciente
           lblNombrePaciente.Caption = Trim(oDOPaciente.apellidoPaterno) & " " & Trim(oDOPaciente.apellidoMaterno) & " " & oDOPaciente.PrimerNombre
        Else
           ml_idPaciente = 0
           lblNombrePaciente.Caption = ""
        End If
      End If
   End If
End Sub
