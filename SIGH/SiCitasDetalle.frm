VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form SiCitasDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "SiCitasDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1080
      Left            =   -30
      TabIndex        =   11
      Top             =   2640
      Width           =   6660
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SiCitasDetalle.frx":08CA
         DownPicture     =   "SiCitasDetalle.frx":0D8E
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
         Left            =   3480
         Picture         =   "SiCitasDetalle.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SiCitasDetalle.frx":1766
         DownPicture     =   "SiCitasDetalle.frx":1BC6
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
         Left            =   1935
         Picture         =   "SiCitasDetalle.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame fraProg 
      Height          =   2670
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   6645
      Begin VB.TextBox txtMotivoAnulacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1740
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1755
         Width           =   4710
      End
      Begin VB.Frame Frame 
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
         Height          =   510
         Left            =   75
         TabIndex        =   18
         Top             =   1185
         Width           =   6375
         Begin Threed.SSOption optActivo 
            Height          =   240
            Left            =   735
            TabIndex        =   19
            Top             =   180
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Activo"
            Value           =   -1
         End
         Begin Threed.SSOption optAnulado 
            Height          =   240
            Left            =   2835
            TabIndex        =   20
            Top             =   180
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Anulado"
         End
      End
      Begin VB.CheckBox ChkDomingo 
         Alignment       =   1  'Right Justify
         Caption         =   "No considerar DOMINGOS"
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
         Left            =   4080
         TabIndex        =   16
         Top             =   915
         Value           =   1  'Checked
         Width           =   2490
      End
      Begin VB.CheckBox chkSabado 
         Caption         =   "No considerar SABADOS"
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
         Left            =   1605
         TabIndex        =   15
         Top             =   900
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.TextBox txtCuposCE 
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
         Left            =   6030
         TabIndex        =   5
         Top             =   2310
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txtCupostotales 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   2325
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txtMedico 
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
         Height          =   315
         Left            =   1605
         TabIndex        =   7
         Top             =   150
         Width           =   4980
      End
      Begin MSMask.MaskEdBox txtHoraInicio 
         Height          =   315
         Left            =   3075
         TabIndex        =   1
         Top             =   525
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraFin 
         Height          =   315
         Left            =   5820
         TabIndex        =   3
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaIni 
         Height          =   315
         Left            =   1605
         TabIndex        =   0
         Top             =   525
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   4350
         TabIndex        =   2
         Top             =   540
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo de Anulación"
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
         Left            =   60
         TabIndex        =   22
         Top             =   1800
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "al"
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
         Left            =   4140
         TabIndex        =   17
         Top             =   570
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° Cupos para Consulta Externa"
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
         Left            =   3390
         TabIndex        =   14
         Top             =   2370
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Cupos totales"
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
         TabIndex        =   13
         Top             =   2370
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label44 
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
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
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
         Left            =   150
         TabIndex        =   10
         Top             =   525
         Width           =   1005
      End
   End
End
Attribute VB_Name = "SIcitasdetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
''------------------------------------------------------------------------------------
''        Inicio de código autogenerado para la clase: POTiposFinanciamiento
''        Autor: William Castro Grijalva
''        Fecha: 30/08/2004 12:28:31 p.m.
''        Empresa: Digital Works Corporation
''        Todos los derechos reservados
''        Control De Cambios:
''------------------------------------------------------------------------------------
''        Autor                      Fecha                      Cambio
''------------------------------------------------------------------------------------
'
'Dim mo_Teclado As New sighentidades.Teclado
'Dim mo_Formulario As New sighentidades.Formulario
'Dim mo_LaboratorioProg As New DoLaboratorioProg
'Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
'Dim mo_ReglasConfiguarcionReslab As New SIGHNegocios.ReglasConfiguarcionReslab
'Dim ml_idUsuario As Long
'Dim ms_MensajeError As String
'Dim mi_Opcion As sghOpciones
'Dim mb_ExistenDatos As Boolean
'Dim ml_IdProgramacion As Long
'Dim mo_cmbTipoConceptoF As New sighentidades.ListaDespleglable
'Dim mo_cmbCajaTiposComprobante As New sighentidades.ListaDespleglable
'Dim oRsTmpF As New Recordset
'Dim oRsTmpR As New Recordset
'Dim mo_lnIdTablaLISTBARITEMS As Long
'Dim mo_lcNombrePc As String
'Dim ml_idGrupo As Long
'Dim mo_FechaInicial As Date
'Property Let FechaInicial(lValue As Date)
'   mo_FechaInicial = lValue
'
'End Property
'
'
'Property Let lcNombrePc(lValue As String)
'   mo_lcNombrePc = lValue
'End Property
'Property Let lnIdTablaLISTBARITEMS(lValue As Long)
'   mo_lnIdTablaLISTBARITEMS = lValue
'End Property
'
'Sub CargarComboBoxes()
'
'End Sub
'Property Let ExistenDatos(bValue As Boolean)
'   mb_ExistenDatos = bValue
'End Property
'Property Get ExistenDatos() As Boolean
'   ExistenDatos = mb_ExistenDatos
'End Property
'Property Let Opcion(iValue As sghOpciones)
'   mi_Opcion = iValue
'End Property
'Property Get Opcion() As sghOpciones
'   Opcion = mi_Opcion
'End Property
'Property Let MensajeError(sValue As String)
'   ms_MensajeError = sValue
'End Property
'Property Get MensajeError() As String
'   MensajeError = ms_MensajeError
'End Property
'Property Let idUsuario(lValue As Long)
'   ml_idUsuario = lValue
'End Property
'Property Get idUsuario() As Long
'   idUsuario = ml_idUsuario
'End Property
'Property Let IdProgramacion(lValue As Long)
'   ml_IdProgramacion = lValue
'End Property
'Property Let idGrupo(lValue As Long)
'   ml_idGrupo = lValue
'End Property
'Property Get idGrupo() As Long
'   idGrupo = ml_idGrupo
'End Property
'
'
'
'
''------------------------------------------------------------------------------------
''   CargarDatosAlFormulario
''   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'Sub CargaDatosDefault()
'    Dim oRsTmp1 As New Recordset
'    Set oRsTmp1 = mo_ReglasConfiguarcionReslab.LabGruposSeleccionarTodos
'    oRsTmp1.Filter = "idGrupo=" & ml_idGrupo
'    If oRsTmp1.RecordCount > 0 Then
'       txtMedico.Text = oRsTmp1!nombreGrupo
'       If mi_Opcion = sghAgregar Then
'            txtHoraInicio.Text = oRsTmp1!HoraInicio
'            txtHoraFin.Text = oRsTmp1!HoraFin
'            txtCupostotales.Text = oRsTmp1!cuposTotal
'            txtCuposCE.Text = oRsTmp1!cuposCe
'            txtFechaIni.Text = Format(mo_FechaInicial, sighentidades.DevuelveFechaSoloFormato_DMY)
'            txtFechaFin.Text = Format("31/12/" & Trim(Str(Year(Date))), sighentidades.DevuelveFechaSoloFormato_DMY)
'       End If
'    End If
'    oRsTmp1.Close
'    Set oRsTmp1 = Nothing
'End Sub
'
'Sub CargarDatosAlFormulario()
' CargaDatosDefault
' Select Case mi_Opcion
'     Case sghAgregar
'     Case sghModificar
'         CargarDatosALosControles
'     Case sghConsultar
'         CargarDatosALosControles
'     Case sghEliminar
'         CargarDatosALosControles
' End Select
'End Sub
'
''------------------------------------------------------------------------------------
''   CargarDatosAlFormulario
''   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub Form_Load()
'       Frame.Enabled = False
'       txtMotivoAnulacion.Enabled = False
'       Select Case mi_Opcion
'       Case sghAgregar
'           Me.Caption = "Agregar Programación en Laboratorio"
'       Case sghModificar
'           Me.Caption = "Modificar Programación en Laboratorio"
'           Frame.Enabled = True
'       Case sghConsultar
'           Me.Caption = "Consultar Programación en Laboratorio"
'       Case sghEliminar
'           Me.Caption = "Eliminar Programación en Laboratorio"
'           txtMotivoAnulacion.Enabled = True
'       End Select
'
'       CargarComboBoxes
'       CargarDatosAlFormulario
'       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
'End Sub
'
''------------------------------------------------------------------------------------
''   CargarDatosAlFormulario
''   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub Form_Activate()
'   If mi_Opcion <> sghAgregar Then
'       If Not mb_ExistenDatos Then
'           Me.Visible = False
'       End If
'   End If
'End Sub
'Sub AdministrarKeyPreview(KeyCode As Integer)
'   Select Case KeyCode
'       Case vbKeyEscape
'           btnCancelar_Click
'       Case vbKeyF2
'           btnAceptar_Click
'       End Select
'End Sub
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   AdministrarKeyPreview KeyCode
'End Sub
'
'Private Sub btnAceptar_Click()
'   If btnAceptar.Enabled = False Then
'      Exit Sub
'   End If
'   Select Case mi_Opcion
'   Case sghAgregar
'       If ValidarDatosObligatorios() Then
'           If ValidarReglas() Then
'               If AgregarDatos() Then
'                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
'                   'LimpiarFormulario
'                   Me.Visible = False
'               Else
'                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasLaboratorio.MensajeError, vbExclamation, Me.Caption
'               End If
'           End If
'       End If
'   Case sghModificar
'       If ValidarDatosObligatorios() Then
'           If ValidarReglas() Then
'               If ModificarDatos() Then
'                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
'                   Me.Visible = False
'               Else
'                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasLaboratorio.MensajeError, vbExclamation, Me.Caption
'               End If
'           End If
'       End If
'   Case sghEliminar
'           If ValidarReglas() Then
'               If EliminarDatos() Then
'                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
'                   Me.Visible = False
'               Else
'                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasLaboratorio.MensajeError, vbExclamation, Me.Caption
'               End If
'           End If
'   End Select
'End Sub
'
'Private Sub btnCancelar_Click()
'   Me.Visible = False
'End Sub
'
'Function ValidarDatosObligatorios() As Boolean
'   Dim sMensaje As String
'   ValidarDatosObligatorios = False
'   If Me.txtFechaIni.Text = sighentidades.FECHA_VACIA_DMY Then
'       sMensaje = sMensaje + "Ingrese la Fecha Inicial" + Chr(13)
'   End If
'   If Me.txtHoraInicio.Text = sighentidades.HORA_VACIA_HM Then
'       sMensaje = sMensaje + "Ingrese la Hora Inicial" + Chr(13)
'   End If
'   If Me.txtFechaFin.Text = sighentidades.FECHA_VACIA_DMY Then
'       sMensaje = sMensaje + "Ingrese la Fecha Final" + Chr(13)
'   End If
'   If Me.txtHoraFin.Text = sighentidades.HORA_VACIA_HM Then
'       sMensaje = sMensaje + "Ingrese la hora Final" + Chr(13)
'   End If
'   If Val(Me.txtCupostotales.Text) <= 0 Then
'       sMensaje = sMensaje + "Los CUPOS TOTALES deben ser mayores a cero" + Chr(13)
'   End If
'   If Val(Me.txtCuposCE.Text) < 0 Then
'       sMensaje = sMensaje + "Los CUPOS EN CONSULTA EXTERNA no pueden ser menores a cero" + Chr(13)
'   End If
'   If Val(Me.txtCuposCE.Text) > Val(Me.txtCupostotales.Text) Then
'       sMensaje = sMensaje + "Los CUPOS EN CONSULTA EXTERNA no pueden ser menores mayor a CUPOS TOTALES" + Chr(13)
'   End If
'   If CDate(Me.txtFechaFin.Text & " " & Me.txtHoraFin.Text) < CDate(Me.txtFechaIni.Text & " " & Me.txtHoraInicio.Text) Then
'       sMensaje = sMensaje + "La FECHA/HORA INICIAL no puede ser mayor a la FECHA/HORA FINAL" + Chr(13)
'   End If
'   If sMensaje <> "" Then
'       MsgBox sMensaje, vbInformation, Me.Caption
'       Exit Function
'   End If
'   ValidarDatosObligatorios = True
'End Function
'Function ValidarReglas() As Boolean
'   ValidarReglas = False
'   ValidarReglas = True
'End Function
''------------------------------------------------------------------------------------
''   Cargar datos al objetos de datos
''   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub CargaDatosAlObjetosDeDatos()
'
'   With mo_LaboratorioProg
'           .cupos = Val(Me.txtCupostotales.Text)
'           .cuposCe = Val(Me.txtCuposCE.Text)
'           Select Case mi_Opcion
'           Case sghEliminar
'              .estado = 0
'           Case sghAgregar
'              .estado = 1
'           Case sghModificar
'              .estado = IIf(Me.optActivo.Value = True, 1, 0)
'           End Select
'           .fecha = CDate(Me.txtFechaIni.Text)
'           .HoraFinal = CDate(Me.txtFechaFin.Text)
'           .HoraInicio = Me.txtHoraInicio.Text
'           .HoraFinal = Me.txtHoraFin.Text
'           If Opcion = sghAgregar Then
'              .idGrupo = ml_idGrupo
'           End If
'           '.idProgramacion = ml_idProgramacion
'           .IdUsuarioAuditoria = ml_idUsuario
'           .MotivoAnulacion = Me.txtMotivoAnulacion.Text
'   End With
'
'End Sub
'
''------------------------------------------------------------------------------------
''        Agregar Datos
''------------------------------------------------------------------------------------
'
'Function AgregarDatos() As Boolean
'
'   CargaDatosAlObjetosDeDatos
'   AgregarDatos = mo_ReglasLaboratorio.LaboratorioProgramacionAgregar(mo_LaboratorioProg, _
'                                          mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Me.txtFechaFin.Text, _
'                                          IIf(chkSabado.Value = 1, True, False), IIf(Me.ChkDomingo.Value = 1, True, False))
'
'End Function
'
''------------------------------------------------------------------------------------
''        Modificar Datos
''------------------------------------------------------------------------------------
'
'Function ModificarDatos() As Boolean
'
'   CargaDatosAlObjetosDeDatos
'   ModificarDatos = mo_ReglasLaboratorio.LaboratorioProgramacionModificar(mo_LaboratorioProg, mo_lnIdTablaLISTBARITEMS, _
'                                         mo_lcNombrePc, _
'                                          IIf(chkSabado.Value = 1, True, False), IIf(Me.ChkDomingo.Value = 1, True, False))
'End Function
'
''------------------------------------------------------------------------------------
''        Eliminar Datos
''------------------------------------------------------------------------------------
'
'Function EliminarDatos() As Boolean
'   If Me.txtMotivoAnulacion.Text = "" Then
'      MsgBox "Debe ingresar el MOTIVO DE ANULACION", vbInformation, Me.Caption
'      Exit Function
'   End If
'   mo_LaboratorioProg.MotivoAnulacion = Me.txtMotivoAnulacion.Text
'   EliminarDatos = mo_ReglasLaboratorio.LaboratorioProgramacionEliminar(mo_LaboratorioProg, _
'                                            mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Me.txtFechaFin.Text)
'End Function
'
''------------------------------------------------------------------------------------
''   Llenar Datos Al Formulario
''   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub CargarDatosALosControles()
'
'        Set mo_LaboratorioProg = mo_ReglasLaboratorio.LaboratorioProgramacionSeleccionarXgrupoFecha(ml_idGrupo, mo_FechaInicial)
'        If mo_ReglasLaboratorio.MensajeError <> "" Then
'             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminFacturacion.MensajeError, vbCritical, Me.Caption"
'             mb_ExistenDatos = False
'             Exit Sub
'        End If
'
'        If Not mo_LaboratorioProg Is Nothing Then
'            If mo_LaboratorioProg.HoraInicio = "" Then
'               Exit Sub
'            End If
'            With mo_LaboratorioProg
'                 Me.txtCupostotales.Text = .cupos
'                 Me.txtCuposCE.Text = .cuposCe
'                ' .estado
'                 Me.txtFechaIni.Text = Format(.fecha, sighentidades.DevuelveFechaSoloFormato_DMY)
'                 Me.txtFechaFin.Text = Format(.fecha, sighentidades.DevuelveFechaSoloFormato_DMY)
'                 Me.txtHoraFin.Text = .HoraFinal
'                 Me.txtHoraInicio.Text = .HoraInicio
'                 Me.txtMotivoAnulacion.Text = .MotivoAnulacion
'                 If .estado = 1 Then
'                    Me.optActivo.Value = True
'                 Else
'                    Me.optAnulado.Value = True
'                 End If
'                ' .idGrupo
'                '.idProgramacion
'                mb_ExistenDatos = True
'            End With
'        Else
'            mb_ExistenDatos = False
'            Exit Sub
'        End If
'
'End Sub
'
''------------------------------------------------------------------------------------
''   Llenar Datos Al Formulario
''   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub LimpiarFormulario()
'
'
'End Sub
'
'
'
'
'
'
'
'
'
'
'Private Sub optActivo_Click(Value As Integer)
'  If optActivo.Value = True Then
'     Me.txtMotivoAnulacion.Text = ""
'  End If
'End Sub
'
'Private Sub txtCuposCE_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtCuposCE
'    AdministrarKeyPreview KeyCode
'
'End Sub
'
'Private Sub txtCupostotales_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtCupostotales
'    AdministrarKeyPreview KeyCode
'
'End Sub
'
'Private Sub txtFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtFechaFin
'    AdministrarKeyPreview KeyCode
'
'End Sub
'
'Private Sub txtFechaIni_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtFechaIni
'    AdministrarKeyPreview KeyCode
'
'End Sub
'
'
'
'
'Private Sub txtHoraFin_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtHoraFin
'    AdministrarKeyPreview KeyCode
'
'End Sub
'
'Private Sub txtHoraInicio_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtHoraInicio
'    AdministrarKeyPreview KeyCode
'
'End Sub
'
'
'
