VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form ArchiveroServicioDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "ArchiverosDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos/ninguno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   75
      TabIndex        =   16
      Top             =   4440
      Width           =   3240
   End
   Begin VB.CommandButton btnBuscarServicio 
      Caption         =   "..."
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   870
      Width           =   315
   End
   Begin VB.CommandButton btnBuscarRespArchivo 
      Caption         =   "..."
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   315
   End
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   45
      TabIndex        =   15
      Top             =   4815
      Width           =   6720
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ArchiverosDetalle.frx":0CCA
         DownPicture     =   "ArchiverosDetalle.frx":112A
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
         Left            =   1755
         Picture         =   "ArchiverosDetalle.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ArchiverosDetalle.frx":1A14
         DownPicture     =   "ArchiverosDetalle.frx":1ED8
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
         Left            =   3300
         Picture         =   "ArchiverosDetalle.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   60
      TabIndex        =   13
      Top             =   690
      Width           =   6705
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "ArchiverosDetalle.frx":28B0
         DownPicture     =   "ArchiverosDetalle.frx":2C99
         Height          =   315
         Left            =   1680
         Picture         =   "ArchiverosDetalle.frx":30A5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1005
      End
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "ArchiverosDetalle.frx":34B1
         DownPicture     =   "ArchiverosDetalle.frx":383C
         Height          =   315
         Left            =   2760
         Picture         =   "ArchiverosDetalle.frx":3BCF
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox txtIdServicio 
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
         Left            =   1650
         TabIndex        =   3
         Top             =   195
         Width           =   975
      End
      Begin VB.TextBox txtNombreServicio 
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
         Left            =   3090
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   195
         Width           =   3510
      End
      Begin VB.Label lblIdServicioIngreso 
         Caption         =   "Servicio"
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
         Left            =   165
         TabIndex        =   14
         Top             =   225
         Width           =   1320
      End
   End
   Begin UltraGrid.SSUltraGrid grdServicios 
      Height          =   2625
      Left            =   60
      TabIndex        =   8
      Top             =   1770
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   4630
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
      Caption         =   "Lista de servicios autorizados"
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   6705
      Begin VB.TextBox txtIdEmpleado 
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
         Left            =   1665
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtNombreEmpleado 
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
         Left            =   3075
         TabIndex        =   2
         Top             =   255
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Archivero"
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
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   1230
      End
   End
End
Attribute VB_Name = "ArchiveroServicioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Archiveros
'        Programado por: Castro W
'        Fecha: Enero 2005
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_ArchiveroServicio As New DOArchiveroServicio
Dim mo_AdminArchivoClinico As New ReglasArchivoClinico
Dim mo_AdminServiciosHosp As New ReglasServiciosHosp
Dim mo_AdminReglasCOmunes As New ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdEmpleado As Long
Dim mrs_Servicios As New Recordset
Dim mo_Archiveros As New Collection
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_EsConsultorioAsignado As Boolean

Property Let EsConsultorioAsignado(lValue As Boolean)
    ml_EsConsultorioAsignado = lValue
    If ml_EsConsultorioAsignado = True Then
       Label4.Caption = "Usuario"
       Me.lblIdServicioIngreso.Caption = "Consultorio"
    End If
    '
    Dim lcMensajeLicencia As String
'    If  False Then   'licencia
'       ml_EsConsultorioAsignado = False
'    End If
End Property


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Sub CargarComboBoxes()

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
Property Let IdEmpleado(lValue As Long)
   ml_IdEmpleado = lValue
End Property
Property Get IdEmpleado() As Long
   IdEmpleado = ml_IdEmpleado
End Property

Private Sub chkTodos_Click()
    On Error GoTo errChkT
    Dim oRsTmp1 As New Recordset
    If mrs_Servicios.RecordCount > 0 Then
        mrs_Servicios.MoveFirst
        Do While Not mrs_Servicios.EOF
            mrs_Servicios.Delete
            mrs_Servicios.Update
            mrs_Servicios.MoveNext
        Loop
    End If
    If chkTodos.Value = 1 Then
        Set oRsTmp1 = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro("idTipoServicio=1 and idEstado=1", sghPorDescripcion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
                With mrs_Servicios
                    .AddNew
                    .Fields!idServicio = oRsTmp1!idServicio
                    .Fields!NombreServicio = oRsTmp1!Nombre
                End With
                oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
    End If
errChkT:
    Set oRsTmp1 = Nothing
    grdServicios.Refresh
End Sub


Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)

    grdServicios.Bands(0).Columns("IdServicio").Hidden = True
    
    grdServicios.Bands(0).Columns("NombreServicio").Header.Caption = IIf(ml_EsConsultorioAsignado = True, "Consultorio", "Servicio")
    grdServicios.Bands(0).Columns("NombreServicio").Width = 6000
    
End Sub

Private Sub txtIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdServicio
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicio_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicio, Me.txtNombreServicio
   mo_Formulario.MarcarComoVacio txtIdServicio
End Sub

Private Sub txtIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleado
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdEmpleado_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleado, Me.txtNombreEmpleado
    mo_Formulario.MarcarComoVacio txtIdEmpleado
End Sub

Private Sub txtIdEmpleado_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla ArchiveroServicio
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

Private Sub btnAgregarDx_Click()
Dim lIdServicio As Long
Dim sNombreServicio As String

    Me.txtIdServicio = Trim(Me.txtIdServicio)
    
    If Me.txtIdServicio = "" Then
        MsgBox "Por favor ingresar el servicio", vbInformation, Me.Caption
        Exit Sub
    End If

    Dim oDOServicio As doServicio
    Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
    If Not oDOServicio Is Nothing Then
        lIdServicio = oDOServicio.idServicio
        sNombreServicio = oDOServicio.codigo + " - " + oDOServicio.Nombre
    Else
        MsgBox "El servicio ingresado no existe", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mrs_Servicios.RecordCount > 0 Then
        mrs_Servicios.MoveFirst
        Do While Not mrs_Servicios.EOF
            If mrs_Servicios!idServicio = lIdServicio Then
                MsgBox "El servicio ingresado ya se ha seleccionado", vbInformation, Me.Caption
                Exit Sub
            End If
            mrs_Servicios.MoveNext
        Loop
    End If
    
    With mrs_Servicios
        .AddNew
        .Fields!idServicio = lIdServicio
        .Fields!NombreServicio = sNombreServicio
    End With

End Sub

Private Sub btnBuscarRespArchivo_Click()
    CompletarDatosResponsable Me.txtIdEmpleado, Me.txtNombreEmpleado
End Sub

Private Sub btnBuscarServicio_Click()
    CompletarDatosDeServicio Me.txtIdServicio, Me.txtNombreServicio
End Sub

Private Sub btnQuitarDx_Click()
    On Error Resume Next
    With mrs_Servicios
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla ArchiveroServicio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       
       GenerarRecordsetTemporal
       
       Select Case mi_Opcion
       Case sghAgregar
           If ml_EsConsultorioAsignado = True Then
              Me.Caption = "Agregar Empleados para CITAS"
           Else
              Me.Caption = "Agregar asignación de servicios"
           End If
       Case sghModificar
           If ml_EsConsultorioAsignado = True Then
              Me.Caption = "Modificar Empleados para CITAS"
           Else
              Me.Caption = "Modificar asignación de servicios"
           End If
       Case sghConsultar
           If ml_EsConsultorioAsignado = True Then
              Me.Caption = "Consultar Empleados para CITAS"
           Else
              Me.Caption = "Consultar asignación de servicios"
           End If
       Case sghEliminar
           If ml_EsConsultorioAsignado = True Then
              Me.Caption = "Eliminar Empleados para CITAS"
           Else
              Me.Caption = "Eliminar asignación de servicios"
           End If
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       mo_Apariencia.ConfigurarFilasBiColores Me.grdServicios, sighentidades.GrillaConFilasBicolor
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla ArchiveroServicio
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
                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
                   LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
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
   
   If Val(Me.txtIdEmpleado.Text) = 0 Then
       sMensaje = sMensaje + "Ingrese el código del empleado" + Chr(13)
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
'   Descripción:    Seleccionar un registro unico de la tabla ArchiveroServicio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   Dim oArchiveroServicio As DOArchiveroServicio
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS MOVIMIENTOS
    '---------------------------------------------------------------------------------
    Set mo_Archiveros = New Collection
    If Not (mrs_Servicios.BOF And mrs_Servicios.EOF) Then
        
        mrs_Servicios.MoveFirst
        
        Do While Not mrs_Servicios.EOF
            Set oArchiveroServicio = New DOArchiveroServicio
            With oArchiveroServicio
                .IdArchivero = 0
                .IdEmpleado = Me.txtIdEmpleado.Tag
                .idServicio = mrs_Servicios!idServicio
                .IdUsuarioAuditoria = Me.idUsuario
                .EsConsultorioAsignado = ml_EsConsultorioAsignado
            End With
            mo_Archiveros.Add oArchiveroServicio
            mrs_Servicios.MoveNext
        Loop
    End If
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminArchivoClinico.ArchiveroServicioAgregar(mo_Archiveros, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminArchivoClinico.ArchiveroServicioModificar(mo_Archiveros, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminArchivoClinico.ArchiveroServicioEliminar(mo_Archiveros, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, ml_idUsuario)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla ArchiveroServicio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

        Dim oDOEmpleado As dOEmpleado
        Set oDOEmpleado = mo_AdminReglasCOmunes.EmpleadosSeleccionarPorId(ml_IdEmpleado)
        If Not oDOEmpleado Is Nothing Then
            Me.txtIdEmpleado.Tag = oDOEmpleado.IdEmpleado
            Me.txtIdEmpleado.Text = oDOEmpleado.CodigoPlanilla
            Me.txtNombreEmpleado = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            Dim rsServicios As New Recordset
            Set rsServicios = mo_AdminArchivoClinico.ArchiveroServicioFiltrarPorEmpleado(ml_IdEmpleado)
            rsServicios.Filter = "EsConsultorioAsignado=" & IIf(ml_EsConsultorioAsignado = True, "1", "0")
            Do While Not rsServicios.EOF
                mrs_Servicios.AddNew
                mrs_Servicios.Fields!idServicio = rsServicios!idServicio
                mrs_Servicios.Fields!NombreServicio = rsServicios!NombreServicio
                rsServicios.MoveNext
            Loop
            
            mb_ExistenDatos = True
        Else
            mb_ExistenDatos = False
        End If
    
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla ArchiveroServicio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.IdEmpleado = 0
    Me.txtIdServicio.Text = ""
    Me.txtIdEmpleado.Text = ""
    Me.txtNombreEmpleado = ""
    Me.txtNombreServicio = ""
   
    If mrs_Servicios.RecordCount > 0 Then
        mrs_Servicios.MoveFirst
        Do While Not mrs_Servicios.EOF
            mrs_Servicios.Delete
            mrs_Servicios.Update
            mrs_Servicios.MoveNext
        Loop
    End If
   
End Sub

Sub CompletarDatosResponsable(txtIdResponsable As TextBox, txtNombreResponsable As TextBox)
    Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
    Dim oDOEmpleado As New dOEmpleado
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOEmpleado = mo_AdminReglasCOmunes.EmpleadosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDOEmpleado Is Nothing Then
            txtIdResponsable.Tag = oDOEmpleado.IdEmpleado
            txtIdResponsable.Text = oDOEmpleado.CodigoPlanilla
            txtNombreResponsable = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub
Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDOServicio As New doServicio
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    If ml_EsConsultorioAsignado = True Then
       oBusqueda.idTipoServicio = sghTipoServicio.sghConsultaExterna
       oBusqueda.HabilitarTipoServicio = False
    Else
       oBusqueda.HabilitarTipoServicio = True
    End If
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOServicio Is Nothing Then
            txtIdServicio.Text = oDOServicio.codigo
            txtIdServicio.Tag = oDOServicio.idServicio
            lblDescripcionServicio = oDOServicio.Nombre
        Else
            txtIdServicio.Text = ""
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDOServicio = Nothing
End Sub

Sub CompletarDatosDeEmpleadoEnElLostFocus(txtCodigoPlanilla As TextBox, txtNombre As TextBox)
Dim oDOEmpleado As New dOEmpleado

        If mo_AdminReglasCOmunes.EmpleadosSeleccionarPorCodigo(txtCodigoPlanilla.Text, oDOEmpleado) Then
            txtCodigoPlanilla.Tag = oDOEmpleado.IdEmpleado
            txtCodigoPlanilla.Text = oDOEmpleado.CodigoPlanilla
            txtNombre = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtCodigoPlanilla.Tag = ""
            txtCodigoPlanilla = ""
            txtNombre = ""
        End If
End Sub

Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDOServicio As doServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDOServicio Is Nothing Then
            txtIdServicio.Tag = oDOServicio.idServicio
            lblDescripcionServicio.Text = oDOServicio.Nombre
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio.Text = ""
        End If
   End If

End Sub
Sub GenerarRecordsetTemporal()
    
    With mrs_Servicios
          .Fields.Append "IdServicio", adInteger
          .Fields.Append "NombreServicio", adVarChar, 255
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    
    Set Me.grdServicios.DataSource = mrs_Servicios
    
End Sub

