VERSION 5.00
Begin VB.Form Triaje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   Icon            =   "Triaje.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SIGHCatalogos.ucTriaje ucTriaje1 
      Height          =   2445
      Left            =   45
      TabIndex        =   4
      Top             =   1530
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4313
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   30
      TabIndex        =   9
      Top             =   4155
      Width           =   11025
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "Triaje.frx":0CCA
         DownPicture     =   "Triaje.frx":118E
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
         Left            =   5655
         Picture         =   "Triaje.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   255
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "Triaje.frx":1B66
         DownPicture     =   "Triaje.frx":1FC6
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
         Left            =   4155
         Picture         =   "Triaje.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame frmCabecera 
      Height          =   1440
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   11025
      Begin VB.TextBox txtOtrosDatos 
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
         Left            =   3030
         TabIndex        =   12
         Top             =   990
         Width           =   7845
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
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
         Left            =   2580
         Picture         =   "Triaje.frx":28B0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtProcedencia 
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
         Left            =   3030
         TabIndex        =   3
         Top             =   600
         Width           =   7845
      End
      Begin VB.TextBox txtNcuenta 
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
         MaxLength       =   30
         TabIndex        =   0
         Top             =   210
         Width           =   1095
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3030
         TabIndex        =   1
         Top             =   180
         Width           =   4395
      End
      Begin VB.TextBox txtPlan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7470
         TabIndex        =   2
         Top             =   180
         Width           =   3405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
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
         TabIndex        =   8
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "*  La información de color rojo define los valores NORMALES de los datos del triaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   150
      TabIndex        =   10
      Top             =   3960
      Width           =   6735
   End
End
Attribute VB_Name = "Triaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Triaje o Signos Vitales
'        Programado por: Barrantes D
'        Fecha: Febrero 2011
'
'------------------------------------------------------------------------------------
'debb-jamo
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_idAtencion As Long
Dim mo_DOAtencionesCE As New DOAtencionesCE
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_idPaciente As Long
Dim ml_NroHCPaciente As Long
Dim FCita As String
Dim FServidor As String
Dim NroHCPaciente As Long
Dim ml_triajeOrigen As sightriajeorigen
Dim ml_idCuentaAtencion As String
Dim mb_cerrarAlGuardarNuevo As Boolean
Dim mb_EsAtencionCRED As Boolean
Dim noEjecutarAccion As Boolean
Dim mb_GuardoTriaje As Boolean
Dim lbTienePermisoParaRegistrarAtencionesPasadas As Boolean, lnIdTipoServicio As Long

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property
Property Let NroHistoria(lValue As Long)
    ml_NroHCPaciente = lValue
End Property
Property Get NroHistoria() As Long
    NroHistoria = ml_NroHCPaciente
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property

Property Let TriajeOrigen(lValue As sightriajeorigen)
   ml_triajeOrigen = lValue
End Property

Property Get TriajeOrigen() As sightriajeorigen
   TriajeOrigen = ml_triajeOrigen
End Property

Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property

Property Let EsAtencionCRED(bValue As Boolean)
   mb_EsAtencionCRED = bValue
End Property

Property Let GuardoTriaje(bValue As Boolean)
   mb_GuardoTriaje = bValue
End Property
Property Get GuardoTriaje() As Boolean
   GuardoTriaje = mb_GuardoTriaje
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
    mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
    mo_Formulario.HabilitarDeshabilitar txtProcedencia, False
    mo_Formulario.HabilitarDeshabilitar Me.txtOtrosDatos, False
    Select Case mi_Opcion
     Case sghAgregar
         If ml_idCuentaAtencion <> 0 Then
            txtNcuenta.Text = ml_idCuentaAtencion
            txtNcuenta_LostFocus
            mo_Formulario.HabilitarDeshabilitar txtNcuenta, False
            mo_Formulario.HabilitarDeshabilitar cmdBuscaCuentaPorApellidos, False
            mb_cerrarAlGuardarNuevo = True
         End If
         'mgaray20141003
         btnAceptar.Enabled = False
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
    End Select
 
    Select Case mi_Opcion
         Case sghAgregar
         Case sghModificar
         Case sghConsultar
            Me.btnAceptar.Enabled = False
         Case sghEliminar
     End Select
     
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_ReglasAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(oDOPaciente.IdPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDOPaciente = Nothing
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       FServidor = lcBuscaParametro.RetornaFechaServidorSQL
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Triaje"
       Case sghModificar
           Me.Caption = "Modificar Triaje"
       Case sghConsultar
           Me.Caption = "Consultar Triaje"
       Case sghEliminar
           Me.Caption = "Eliminar Triaje"
       End Select

       lbTienePermisoParaRegistrarAtencionesPasadas = mo_ReglasAdmision.TienePermisosParaModificarAtencionesPasadas
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       
End Sub

Function SeGrabaTriajePorPermisoEspecial() As Boolean
    SeGrabaTriajePorPermisoEspecial = False
    If lbTienePermisoParaRegistrarAtencionesPasadas = True And lnIdTipoServicio = sghTipoServicio.sghConsultaExterna Then
       SeGrabaTriajePorPermisoEspecial = True
    ElseIf lnIdTipoServicio <> sghTipoServicio.sghConsultaExterna Then
       SeGrabaTriajePorPermisoEspecial = True
    End If
    
End Function

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Triaje
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       Else
            On Error Resume Next
            'Me.txtPresion.SetFocus
            ucTriaje1.FocusPulso
       End If
   Else
       txtNcuenta.SetFocus
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
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    mb_GuardoTriaje = True
                    If mb_cerrarAlGuardarNuevo = True Then
                        Me.Visible = False
                    Else
                        LimpiarFormulario
                        Me.txtNcuenta.SetFocus
                        'mgaray20141003
                        Me.btnAceptar.Enabled = False
                    End If
                Else
                    If FCita <> FServidor Then
                        MsgBox "No es posible agregar los datos, debido a que la fecha de la cita es diferente a la fecha actual" + Chr(13) + mo_ReglasAdmision.MensajeError, vbExclamation, Me.Caption
                    Else
                        MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasAdmision.MensajeError, vbExclamation, Me.Caption
                    End If
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    mb_GuardoTriaje = True
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    mb_GuardoTriaje = True
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   ValidarDatosObligatorios = False
   Dim messageError As String
   
   If Me.txtNcuenta.Text = "" Or Me.txtPlan.Text = "" Then
       MsgBox "Ingrese el N° Cuenta", vbInformation, Me.Caption
       Exit Function
   End If
   'Validar campos de peso, talla y perimetro abdominal - 20/07/2020 - GLCC
'   If ucTriaje.txtPerimetroCefalico.Text = "" Then
'       MsgBox "Ingrese el Perimetros Cefalico", vbInformation, Me.Caption
'       Exit Function
'   End If
   If Me.txtNcuenta.Text = "" Or Me.txtPlan.Text = "" Then
       MsgBox "Ingrese el N° Cuenta", vbInformation, Me.Caption
       Exit Function
   End If
   If Me.txtNcuenta.Text = "" Or Me.txtPlan.Text = "" Then
       MsgBox "Ingrese el N° Cuenta", vbInformation, Me.Caption
       Exit Function
   End If
   'Para CUENTA DE EMERGENCIA no son obligatorios los datos
   If lcBuscaParametro.SeleccionaFilaParametro(520) <> "S" And InStr(txtProcedencia.Text, "Emergencia") > 0 Then
      ValidarDatosObligatorios = True
      Exit Function
   End If
   
   If ucTriaje1.validarTodosValoresTriaje() = False Then
        messageError = ucTriaje1.RetornaMensageDeErrorDatosObligatorios
        If messageError <> "" Then
            MsgBox "Debe ingresar los datos: " & vbCrLf & messageError, vbInformation, Me.Caption
            Exit Function
        End If
   End If
   
'   If Me.txtPeso.Text = "" And Me.txtPresion.Text = SIGHEntidades.PresionDevuelveVacia And Me.txtTalla.Text = "" And Me.txtTemperatura.Text = "" Then
'       MsgBox "Debe ingresar alguno de los datos: Peso, Presión, Talla, Temperatura ", vbInformation, Me.Caption
'       Exit Function
'   End If
   ValidarDatosObligatorios = True

End Function

Function ValidarReglas() As Boolean
    Dim messageError As String
   ValidarReglas = False
   
   If ucTriaje1.validarTodosValoresTriaje() = False Then
        messageError = ucTriaje1.RetornaMensageDeErrorReglas
        If messageError <> "" Then
            MsgBox "Los valores de los siguientes datos no son COHERENTES: " & vbCrLf & messageError, vbInformation, Me.Caption
            Exit Function
        End If
   End If
   If ucTriaje1.PresionVerificaSiTieneDatosYsiEstaOK() = False Then
      Exit Function
   End If
   If ucTriaje1.ValidarReglas = False Then
      Exit Function
   End If

   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Triaje
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
    ucTriaje1.CargaDatosAlObjetosDeDatos mo_DOAtencionesCE
   With mo_DOAtencionesCE
        '.idAtencion = ml_idAtencion
        .IdUsuarioAuditoria = ml_idUsuario
        If mi_Opcion = sghAgregar Then
            .TriajeIdUsuario = ml_idUsuario
            .TriajeFecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
        End If
   End With
End Sub

Sub AgregaTriajeAotrosConsultoriosDelDia(ldTriajeFecha As Date)
    On Error Resume Next
    Dim oRsTmp As New Recordset, oRsTmp1 As New Recordset
    Dim oAtencionesCE As New AtencionesCE
    Dim o_DOAtencionesCE As New DOAtencionesCE
    Dim ldFechaI As Date
    Set oRsTmp = mo_ReglasAdmision.AtencionesDePacienteElmismoDiaCE(ml_idPaciente, ldTriajeFecha)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       ldFechaI = CDate(Format(ldTriajeFecha, "dd/mm/yyyy"))
       Do While Not oRsTmp.EOF
          If mi_Opcion <> sghEliminar And ml_idAtencion <> oRsTmp.Fields!idAtencion And IsNull(oRsTmp.Fields!FechaEgreso) And ml_idPaciente = oRsTmp.Fields!IdPaciente And ldFechaI = oRsTmp.Fields!fechaingreso Then
                Select Case mi_Opcion
                'Case sghAgregar
                '     If mo_ReglasAdmision.AtencionCEAgregar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(oRsTmp.Fields!idAtencion))) = False Then
                '        MsgBox "Problemas con los OTROS CONSULTORIOS, al Agregar el Triaje", vbInformation, Me.Caption
                '     End If
                Case sghAgregar, sghModificar
                     o_DOAtencionesCE.idAtencion = oRsTmp.Fields!idAtencion
                     Set o_DOAtencionesCE = mo_ReglasAdmision.AtencionCESeleccionarPorId(oRsTmp.Fields!idAtencion)
                     ucTriaje1.CargaDatosAlObjetosDeDatos o_DOAtencionesCE
                     With o_DOAtencionesCE
                         .IdUsuarioAuditoria = ml_idUsuario
                         If mi_Opcion = sghAgregar Then
                             .TriajeIdUsuario = ml_idUsuario
                             .TriajeFecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
                         End If
                     End With
                     If o_DOAtencionesCE.NroHistoriaClinica = 0 Then
                        o_DOAtencionesCE.NroHistoriaClinica = mo_DOAtencionesCE.NroHistoriaClinica
                        o_DOAtencionesCE.idAtencion = oRsTmp.Fields!idAtencion
                        If mo_ReglasAdmision.AtencionCEAgregar(o_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(oRsTmp.Fields!idAtencion))) = False Then
                           MsgBox "Problemas con los OTROS CONSULTORIOS, al Agregar el Triaje", vbInformation, Me.Caption
                        End If
                     Else
                        mo_DOAtencionesCE.NroHistoriaClinica = mo_DOAtencionesCE.NroHistoriaClinica
                        If mo_ReglasAdmision.AtencionCEModificar(o_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(oRsTmp.Fields!idAtencion))) = False Then
                           MsgBox "Problemas con los OTROS CONSULTORIOS, al modificar el Triaje", vbInformation, Me.Caption
                        End If
                     End If
                     
'                     Set oRsTmp1 = mo_ReglasAdmision.AtencionCESeleccionarPorIdAtencion(mo_DOAtencionesCE.idAtencion)
'                     If oRsTmp1.RecordCount = 0 Then
'                        If mo_ReglasAdmision.AtencionCEAgregar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(oRsTmp.Fields!idAtencion))) = False Then
'                           MsgBox "Problemas con los OTROS CONSULTORIOS, al Agregar el Triaje", vbInformation, Me.Caption
'                        End If
'                     Else
'                        If mo_ReglasAdmision.AtencionCEModificar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(oRsTmp.Fields!idAtencion))) = False Then
'                           MsgBox "Problemas con los OTROS CONSULTORIOS, al modificar el Triaje", vbInformation, Me.Caption
'                        End If
'                     End If
                'Case sghEliminar
                '     If mo_ReglasAdmision.AtencionCEeliminar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(oRsTmp.Fields!idAtencion))) = False Then
                '        MsgBox "Problemas con los OTROS CONSULTORIOS, al eliminar el Triaje", vbInformation, Me.Caption
                '     End If
                End Select
          End If
          oRsTmp.MoveNext
       Loop
    End If
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oAtencionesCE = Nothing
    Set o_DOAtencionesCE = Nothing
End Sub
'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    mo_DOAtencionesCE.idAtencion = ml_idAtencion
    CargaDatosAlObjetosDeDatos
    With mo_DOAtencionesCE
        .CitaAntecedente = ""
        .CitaDiagMed = ""
        .CitaDniMedicoJamo = ""
        .CitaExamenClinico = ""
        .CitaFecha = 0
        .CitaFechaAtencion = 0
        .CitaIdServicio = 0
        .CitaIdUsuario = 0
        .CitaMedico = ""
        .CitaMotivo = ""
        .CitaObservaciones = ""
        .CitaServicioJamo = ""
        .CitaTratamiento = ""
    End With
    
    FServidor = lcBuscaParametro.RetornaFechaServidorSQL
    'restriccion cambiada para el caso de emergencia y hospitalizacion
    If FCita = FServidor Or (FCita <= FServidor And (ml_triajeOrigen = Emergencia Or ml_triajeOrigen = Hospitalizacion)) Or _
                                                                                        SeGrabaTriajePorPermisoEspecial = True Then
        AgregarDatos = mo_ReglasAdmision.AtencionCEAgregar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)))
        AgregaTriajeAotrosConsultoriosDelDia mo_DOAtencionesCE.TriajeFecha
    End If
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_ReglasAdmision.AtencionCEModificar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)))
    AgregaTriajeAotrosConsultoriosDelDia mo_DOAtencionesCE.TriajeFecha
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    EliminarDatos = mo_ReglasAdmision.AtencionCEeliminar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)))
    AgregaTriajeAotrosConsultoriosDelDia mo_DOAtencionesCE.TriajeFecha
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
       Dim oAtencionesCE As New AtencionesCE
       Dim oConexion As New Connection
       oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
       mo_DOAtencionesCE.idAtencion = ml_idAtencion
       Set oAtencionesCE.Conexion = oConexion
       If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = False Then
           MsgBox "No se pudo obtener los datos" + Chr(13) + oAtencionesCE.MensajeError, vbInformation, Me.Caption
           mb_ExistenDatos = False
           Exit Sub
       End If
       If Not mo_DOAtencionesCE Is Nothing Then
            With mo_DOAtencionesCE
                 Me.txtNcuenta.Text = mo_ReglasAdmision.AtencionesSeleccionarNroCuentaSegunIdAtencion(ml_idAtencion)
                 ucTriaje1.CargarDatosALosControles mo_DOAtencionesCE
                 mb_ExistenDatos = True
            End With
            mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
            cmdBuscaCuentaPorApellidos.Enabled = False
            '@comentado
'            txtNcuenta_LostFocus
            ucTriaje1.txtTemperatura_LostFocus
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
End Sub


Sub LimpiarFormulario()
    Me.idAtencion = 0
    ucTriaje1.LimpiarControles
    Me.txtNcuenta.Text = ""
    Me.txtDatosDeCuenta.Text = ""
    Me.txtProcedencia.Text = ""
    Me.txtPlan.Text = ""
    ml_idPaciente = 0
    txtOtrosDatos.Text = ""
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
End Sub

Private Sub txtNcuenta_LostFocus()
   If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
        Dim oDOPaciente As New DOPaciente
        Dim oRsTmp As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim lbSigue As Boolean, lbEsUnEPS As Boolean
        Dim oAtencionesCE As New AtencionesCE
        Dim oConexion As New Connection
        Dim oConexionExterna As New Connection
        
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        
        mo_DOAtencionesCE.idAtencion = ml_idAtencion
        
        
        Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
        lbSigue = True
        txtDatosDeCuenta.Text = ""
        txtPlan.Text = ""
        txtProcedencia.Text = ""
        If oRsTmp.RecordCount > 0 Then
            Set oAtencionesCE.Conexion = oConexionExterna
            mo_DOAtencionesCE.NroHistoriaClinica = oRsTmp.Fields!NroHistoriaClinica
            If oAtencionesCE.SeleccionarPorNroHistoria(mo_DOAtencionesCE) Then
                '@comentado
                 'txtTalla.Text = mo_DOAtencionesCE.TriajeTalla
                 ucTriaje1.setValorTalla mo_DOAtencionesCE.TriajeTalla, mo_DOAtencionesCE.TriajePeso
                 
            End If
            
            mo_DOAtencionesCE.NroHistoriaClinica = oRsTmp.Fields!NroHistoriaClinica
            NroHCPaciente = mo_DOAtencionesCE.NroHistoriaClinica
            txtDatosDeCuenta.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsTmp.Fields!NroHistoriaClinica)), False) & " - " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!PrimerNombre) & " (Edad: " & Trim(Str(oRsTmp.Fields!Edad)) & " " & Trim(oRsTmp.Fields!tedad) & ")"
            txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento
            ml_idAtencion = oRsTmp.Fields!idAtencion
            txtProcedencia.Text = "F.Ing: " & oRsTmp.Fields!fechaingreso & " - " & IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ") " & mo_ReglasFacturacion.BuscaServicioActualDelPaciente(oRsTmp.Fields!IdServicioIngreso)
            FCita = oRsTmp.Fields!fechaingreso
            ml_idPaciente = oRsTmp.Fields!IdPaciente
            lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
            
            Set oDOPaciente = mo_ReglasAdmision.PacientesSeleccionarPorId(ml_idPaciente, oConexion)
            txtOtrosDatos.Text = "Gs: " & oDOPaciente.GrupoSanguineo & "    Factor RH: " & oDOPaciente.FactorRh
            
            If mi_Opcion <> sghConsultar Then
                'If (mi_Opcion <> sghConsultar And Format(FServidor, SIGHEntidades.DevuelveFechaSoloFormato_DMY) <> Format(oRsTmp.Fields!fechaingreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY)) And ml_triajeOrigen = ConsultaExterna Then
                If Format(FServidor, sighentidades.DevuelveFechaSoloFormato_DMY) <> _
                        Format(oRsTmp.Fields!fechaingreso, sighentidades.DevuelveFechaSoloFormato_DMY) _
                        And ml_triajeOrigen = ConsultaExterna Then
                    If SeGrabaTriajePorPermisoEspecial = False Then
                    MsgBox "La CITA tiene fecha diferente a la de HOY", vbInformation, Me.Caption
                    lbSigue = False
                    End If
                End If
            
       
                If oRsTmp.Fields!idEstado <> 1 Then
                   'If mi_Opcion <> sghConsultar Then
                      MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
'                      If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
'                         btnAceptar.Enabled = False
'                      Else
                         lbSigue = False
'                      End If
                   'End If
                End If
                
                If lbSigue = True And mi_Opcion = sghAgregar And YaSeRegistroTriaje(oRsTmp.Fields!idAtencion) = True Then
                    MsgBox "Ya se registró el Triaje para esa Cuenta", vbInformation, Me.Caption
                    lbSigue = False
                End If
                'If lbSigue = True And mi_Opcion <> sghConsultar And SIGHEntidades.esfecha(Format(oRsTmp.Fields!FechaEgreso, "dd/mm/yyyy"), "DD/MM/AAAA") = True Then
                If lbSigue = True And sighentidades.EsFecha(Format(oRsTmp.Fields!FechaEgreso, "dd/mm/yyyy"), "DD/MM/AAAA") = True Then
                    If ml_triajeOrigen = sightriajeorigen.Triaje Or mi_Opcion = sghEliminar Then
                        If mi_Opcion = sghEliminar Then
                            MsgBox "Ya se atendió al Paciente, no se podrá Eliminar TRIAJE", vbInformation, Me.Caption
                        Else
                            MsgBox "Ya se atendió al Paciente, no se podrá registrar desde el Area de TRIAJE", vbInformation, Me.Caption
                        End If
                        lbSigue = False
                        
'                        btnAceptar.Enabled = False
'                        ucTriaje1.idAtencion = ml_idAtencion
'                        ucTriaje1.Inicializar
'                        ucTriaje1.BloqueoTodosLosControles
'                        Exit Sub
                     End If
                End If
                If lbSigue Then
                    lbEsUnEPS = False
                    If Not IsNull(oRsTmp!EpsPorcentaje) Then
                       If oRsTmp!EpsPorcentaje > 0 Then
                          lbEsUnEPS = True
                       End If
                    End If
                    Dim rsTempPago As New ADODB.Recordset
                    If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oRsTmp.Fields!IdFormaPago, oConexion) = True Or lbEsUnEPS = True Then
                        If mo_ReglasAdmision.EsServicioCostoCero(oRsTmp.Fields!IdServicioIngreso) = False Then
                            Set oRsTmp = mo_ReglasFacturacion.FacturacionPaquetesCEporIdPuntoCargaNrocuentaIdEspecialidad(Val(txtNcuenta.Text), oRsTmp!IdEspecialidad, 6, oConexion)
                            If oRsTmp.RecordCount = 0 Then
                                Set oRsTmp = mo_ReglasFacturacion.DevuelveSiPagoConsultaMedicaEnCaja(ml_idAtencion, "1", oConexion)
                                oRsTmp.Filter = "idTipoFinanciamiento=1"
                                If oRsTmp.RecordCount > 0 Then
                                   If oRsTmp.Fields!IdEstadoFacturacion <> 4 Then
                                      MsgBox "No pagó CITA el paciente: " & txtDatosDeCuenta.Text, vbInformation, Me.Caption
                                      LimpiarFormulario
                                      Set oRsTmp = Nothing
                                      Exit Sub
                                   End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
                
                On Error Resume Next
                ucTriaje1.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
                ucTriaje1.FechaTriaje = oRsTmp.Fields!fechaingreso
                ucTriaje1.idAtencion = ml_idAtencion
                ucTriaje1.EstadoPaciente = 0
                ucTriaje1.TriajeOrigen = IIf(ml_triajeOrigen = 0, sightriajeorigen.Triaje, ml_triajeOrigen)
                ucTriaje1.EsAtencionCRED = mb_EsAtencionCRED
                ucTriaje1.Inicializar
                If txtNcuenta.Locked = False Then
                    ucTriaje1.FocusPulso
                End If
                If lbSigue = False Then
                    btnAceptar.Enabled = False
                    ucTriaje1.BloqueoTodosLosControles
                'mgaray20141003
                ElseIf mi_Opcion <> sghConsultar Then
                    btnAceptar.Enabled = True
                End If

        'mgaray20141022
        Else
            MsgBox "Número de cuenta no encontrado", vbInformation, "Registro Triaje"
'            txtNcuenta.SetFocus
            btnAceptar.Enabled = False
        End If
        oRsTmp.Close
        Set oRsTmp = Nothing
        oConexion.Close
        Set oConexion = Nothing
        Set oDOPaciente = Nothing
    End If
    If mi_Opcion = sghConsultar Then
        ucTriaje1.BloqueoTodosLosControles
    End If
End Sub

Function YaSeRegistroTriaje(lnIdAtencion As Long) As Boolean
       Dim oAtencionesCE As New AtencionesCE
       Dim oConexion As New Connection
       YaSeRegistroTriaje = False
       oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
       Set oAtencionesCE.Conexion = oConexion
       mo_DOAtencionesCE.idAtencion = lnIdAtencion
       If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = True Then
          If mo_DOAtencionesCE.NroHistoriaClinica > 0 Then
             YaSeRegistroTriaje = True
          End If
       Else
       End If
       oConexion.Close
       Set oConexion = Nothing
       Set oAtencionesCE = Nothing
End Function

Private Sub ucTriaje1_SePresionoTeclaEspecial(KeyCode As Integer)
'    AdministrarKeyPreview KeyCode
End Sub



