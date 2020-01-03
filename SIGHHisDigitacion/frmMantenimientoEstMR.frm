VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form frmMantenimientoEstMR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   Icon            =   "frmMantenimientoEstMR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraServiciosCentral 
      Caption         =   "Servicios de Centro Digitación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   30
      TabIndex        =   9
      Top             =   1230
      Width           =   7335
      Begin VB.ComboBox cmbDepartamento 
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
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cmbEspecialidad 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5880
         Picture         =   "frmMantenimientoEstMR.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   1305
      End
      Begin UltraGrid.SSUltraGrid ugvServicios 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3201
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
         Caption         =   "Consultorios Estandares"
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Especialidad"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame fraServiciosPorEstablecimiento 
      Caption         =   "Lista de Consultorios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   30
      TabIndex        =   6
      Top             =   4170
      Width           =   7335
      Begin Threed.SSCommand btnAgregarServicio 
         Height          =   465
         Left            =   6000
         TabIndex        =   7
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMantenimientoEstMR.frx":2C55
         Caption         =   "Agregar"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand btnQuitarServicio 
         Height          =   465
         Left            =   6000
         TabIndex        =   8
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMantenimientoEstMR.frx":5BE1
         Caption         =   "Quitar"
         PictureAlignment=   9
         ShapeSize       =   1
      End
      Begin UltraGrid.SSUltraGrid ugvServiciosPorEstablecimiento 
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4048
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
         Caption         =   "Consultorios del Establecimiento"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Establecimientos de Micro Red"
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
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   7335
      Begin VB.TextBox txtCodigoEstablecimiento 
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
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscarEstablecimiento 
         Height          =   315
         Left            =   5880
         Picture         =   "frmMantenimientoEstMR.frx":8063
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox txtNombreEstablecimiento 
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
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   6750
      Width           =   7335
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmMantenimientoEstMR.frx":ACAC
         DownPicture     =   "frmMantenimientoEstMR.frx":B10C
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
         Left            =   2258
         Picture         =   "frmMantenimientoEstMR.frx":B581
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmMantenimientoEstMR.frx":B9F6
         DownPicture     =   "frmMantenimientoEstMR.frx":BEBA
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
         Left            =   3788
         Picture         =   "frmMantenimientoEstMR.frx":C3A6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmMantenimientoEstMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de los establecimientos de Salud de la Micro Red
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ml_lnIdTablaLISTBARITEMS As Long
Dim ms_lcNombrePc As String
Dim ml_IdEstablecimiento As Long
Dim ms_mesajeError As String

Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos           'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mr_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_cmbDepartamento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbEspecialidad As New SIGHEntidades.ListaDespleglable

Dim oRcs_Servicios As New ADODB.Recordset                         'Representan los servicios de un establecimiento dado
Dim oRcs_ServiciosEstablecimiento As New ADODB.Recordset    'Representa los servicios regsitrados en la MR

Dim ms_LoginPC As String
Dim ml_IdUsuario As Long                                'Indica el ID del Usuario que esta en session activa.
Dim mi_Opcion As sghOpciones
Dim mo_Establecimiento As DOEstablecimiento         'Representa el establecimiento actual
Dim ms_NombreEstablcimiento As String
Dim ms_CodigoEstablecimiento As String

'========================== PROPIEDADES ===============================
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   ml_lnIdTablaLISTBARITEMS = lValue
End Property
Property Get lnIdTablaLISTBARITEMS() As Long
   lnIdTablaLISTBARITEMS = ml_lnIdTablaLISTBARITEMS
End Property
Property Let lcNombrePc(sValue As String)
   ms_lcNombrePc = sValue
End Property
Property Get lcNombrePc() As String
   lcNombrePc = ms_lcNombrePc
End Property

Property Let IdEstablecimiento(lValue As Long)
   ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property

Property Let NombreEstablecimiento(lValue As String)
   ms_NombreEstablcimiento = lValue
End Property
Property Get NombreEstablecimiento() As String
   NombreEstablecimiento = ms_NombreEstablcimiento
End Property

Property Let CodigoEstablecimiento(lValue As String)
   ms_CodigoEstablecimiento = lValue
End Property
Property Get CodigoEstablecimiento() As String
   CodigoEstablecimiento = ms_CodigoEstablecimiento
End Property

'========================== EVENTOS ===================================
Private Sub Form_Load()
Set mo_cmbDepartamento.MiComboBox = Me.cmbDepartamento
Set mo_cmbEspecialidad.MiComboBox = Me.cmbEspecialidad
Me.btnAgregarServicio.Enabled = False

CrearTablasTemp
CargarDatosAlFormulario
CargarComboBoxes

'Deshabilita el regsitro de un nuevo establecimiento
If mi_Opcion = sghAgregar Then
    Me.cmdBuscarEstablecimiento.Enabled = True
Else
    Me.cmdBuscarEstablecimiento.Enabled = False
End If

Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Establecimiento"
    Case sghModificar
        Me.Caption = "Modificar Establecimiento"
    Case sghConsultar
        Me.Caption = "Consultar Establecimiento"
        Me.btnAceptar.Enabled = False
    Case sghEliminar
        Me.Caption = "Anular Establecimiento"
End Select

mo_Apariencia.ConfigurarFilasBiColores Me.ugvServicios, SIGHEntidades.GrillaConFilasBicolor
mo_Apariencia.ConfigurarFilasBiColores Me.ugvServiciosPorEstablecimiento, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub cmbDepartamento_Click()
mo_cmbEspecialidad.BoundColumn = "IdEspecialidad"
mo_cmbEspecialidad.ListField = "DescripcionLarga"
On Error Resume Next
Set mo_cmbEspecialidad.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbDepartamento.BoundText))
mo_cmbEspecialidad.BoundText = 0
End Sub

Private Sub btnAgregarServicio_Click()
If oRcs_Servicios.RecordCount <> 0 Then
    If ugvServicios.ActiveRow.Selected Then
        If BuscarServicioActual(CLng(Me.ugvServicios.ActiveRow.Cells("IdServicio").Value)) Then
            Call MsgBox("El servicio existe en el establecimiento Actual.", vbExclamation Or vbSystemModal, Me.Caption)
            Exit Sub
        Else
            If oRcs_ServiciosEstablecimiento.RecordCount <> 0 Then
                oRcs_ServiciosEstablecimiento.MoveFirst
            End If
            With oRcs_ServiciosEstablecimiento
                .AddNew
                .Fields!IdServicio = CLng(Me.ugvServicios.ActiveRow.Cells("IdServicio").Value)
                .Fields!IdEstablecimiento = ml_IdEstablecimiento
                .Fields!Nombre = CStr(Me.ugvServicios.ActiveRow.Cells("Nombre").Value)
                .Fields!IdEstado = 1
                .Update
            End With
        End If
    End If
Else
    MsgBox "No se ha elegido ningún servicio.", vbCritical, "HIS SIGH"
End If
End Sub

Private Sub btnBuscar_Click()
If Me.cmbDepartamento.Text = "" Or Me.cmbEspecialidad.Text = "" Then
    MsgBox "No ha buscado ningún Servicio", vbInformation, "HIS SIGH"
Else
    Set oRcs_Servicios = mr_ReglasHIS.ConsultarRegistroServiciosPorEspec(Val(mo_cmbEspecialidad.BoundText))
    Set Me.ugvServicios.DataSource = oRcs_Servicios
    Me.btnAgregarServicio.Enabled = True
End If
End Sub

Private Sub btnQuitarServicio_Click()
If VerificarRegistrosServicio Then
    If MsgBox("Desea retirar el servicio seleccionado?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton1, Me.Caption) = vbYes Then
        Select Case CLng(Me.ugvServiciosPorEstablecimiento.ActiveRow.Cells("IdEstado").Value)
            Case 1
                'Modificacion de detalle Fisicamente
                With oRcs_ServiciosEstablecimiento
                    If Not .EOF And Not .BOF Then
                       .Delete
                       .Update
                    End If
                End With
            Case 0
                'Ocultamiento de fila activa y ingreso de valor IDESTADO = 3
                Me.ugvServiciosPorEstablecimiento.ActiveRow.Hidden = True
                Me.ugvServiciosPorEstablecimiento.ActiveRow.Cells("IdEstado").Value = 3
        End Select
    End If
End If
End Sub

Private Sub cmdBuscarEstablecimiento_Click()
Dim oForm As New SIGHNegocios.BuscaEstablecimientos
Dim oDoEstablecimiento As New DOEstablecimiento
Dim mo_RcsListaEstablecimientos  As New Recordset

oForm.DescripcionEstablecimiento = Me.txtNombreEstablecimiento.Text
'mgaray201503 Establecimeinto de centro de salud y puesto de salud
oForm.NivelMaximoEstablecimiento = sghTipoEstablecimiento.CentroSalud
oForm.MostrarFormulario

Me.btnAceptar.Enabled = True

If oForm.IdRegistroSeleccionado = 0 Then
    Call MsgBox("No ha seleccionado ningún registro de la Lista.", vbExclamation, Me.Caption)
Else
    'Ingresando los valores del Establecimiento Elegido
    If oForm.BotonPresionado = sghAceptar Then
        Set oDoEstablecimiento = mr_ReglasComunes.EstablecimientosSeleccionarPorId(oForm.IdRegistroSeleccionado)
        If Not oDoEstablecimiento Is Nothing Then
            Set mo_Establecimiento = oDoEstablecimiento
            ml_IdEstablecimiento = oDoEstablecimiento.IdEstablecimiento
            Me.txtNombreEstablecimiento.Text = mo_Establecimiento.Nombre
            Me.txtCodigoEstablecimiento.Text = mo_Establecimiento.Codigo
        End If
    End If

    'Validando si el establecimiento elgido ya esta registrado en la lista principal
    Set mo_RcsListaEstablecimientos = mr_ReglasHIS.ObtenerListaEstablecimientosMR
    
    If mo_RcsListaEstablecimientos.RecordCount <> 0 Then
        mo_RcsListaEstablecimientos.MoveFirst
        While Not mo_RcsListaEstablecimientos.EOF
            If CLng(mo_RcsListaEstablecimientos!IdEstablecimiento) = ml_IdEstablecimiento Then
                MsgBox "El establecimiento seleccionado ya fue registrado.", vbInformation, Me.Caption
                Me.btnAceptar.Enabled = False
                Exit Sub
            End If
            mo_RcsListaEstablecimientos.MoveNext
        Wend
    End If
End If
End Sub

Private Sub ugvServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
Layout.Override.AllowDelete = ssAllowDeleteNo
Layout.Override.CellClickAction = ssClickActionRowSelect

With ugvServicios.Bands(0)
    .Columns("IdServicio").Hidden = True
    .Columns("Nombre").Header.Caption = "Especialidad"
    .Columns("Nombre").Width = 2000
    .Columns("DescripcionLarga").Hidden = True
    .Columns("IdEstado").Hidden = True
End With
End Sub

Private Sub ugvServiciosporEstablecimiento_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
Layout.Override.AllowDelete = ssAllowDeleteNo
Layout.Override.CellClickAction = ssClickActionRowSelect

With ugvServiciosPorEstablecimiento.Bands(0)
'    .Columns("IdHisServEstablecimiento").Hidden = True
    .Columns("IdEstablecimiento").Hidden = True
    .Columns("IdServicio").Hidden = True
    .Columns("Nombre").Header.Caption = "Nombre Servicio"
    .Columns("Nombre").Width = 2000
    .Columns("IdEstado").Hidden = True
End With
End Sub

Private Sub btnAceptar_Click()
If btnAceptar.Enabled = False Then
   Exit Sub
End If
If ValidarDatosObligatorios() Then
    If ValidarReglas() Then
        If ActualizarDatos() Then
            Call MsgBox("Los datos fuerón actualizados satisfactoriamente.", vbInformation, Me.Caption)
            Me.Visible = False
            LimpiarVariablesDeMemoria
        Else
            Call MsgBox("No se pudo actualizar los datos.", vbCritical Or vbSystemModal, Me.Caption)
        End If
    End If
End If
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

'========================== METODOS ===================================
Private Sub CrearTablasTemp()
'Servicios de Establecimiento Actual
With oRcs_ServiciosEstablecimiento
    .Fields.Append "IdEstablecimiento", adInteger, , adFldIsNullable + adFldUpdatable
    .Fields.Append "IdServicio", adInteger, , adFldIsNullable + adFldUpdatable
    .Fields.Append "Nombre", adVarChar, 50, adFldIsNullable + adFldUpdatable
    .Fields.Append "IdEstado", adInteger, , adFldIsNullable + adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With
End Sub

Private Sub CargarDatosAlFormulario()
    'Carga el Nombre del Establecimiento
    If ml_IdEstablecimiento > 0 Then
        Me.txtNombreEstablecimiento.Text = ms_NombreEstablcimiento
        txtCodigoEstablecimiento.Text = ms_CodigoEstablecimiento
        
        'Consulta de Servicios por Establecimiento en la MR
        Set oRcs_Temp = mr_ReglasHIS.ObtenerListaServiciosPorEstablecimiento(ml_IdEstablecimiento)
        If oRcs_Temp.RecordCount <> 0 Then
            oRcs_Temp.MoveFirst
            Do While Not oRcs_Temp.EOF
                With oRcs_ServiciosEstablecimiento
                    .AddNew
                    .Fields!IdEstablecimiento = oRcs_Temp!IdEstablecimiento
                    .Fields!IdServicio = oRcs_Temp!IdServicio
                    .Fields!Nombre = oRcs_Temp!Nombre
                    .Fields!IdEstado = oRcs_Temp!IdEstado
                    .Update
                End With
                oRcs_Temp.MoveNext
            Loop
        End If
        oRcs_Temp.Close
        Set oRcs_Temp = Nothing
    End If
    Set Me.ugvServiciosPorEstablecimiento.DataSource = oRcs_ServiciosEstablecimiento
End Sub

Private Sub CargarComboBoxes()
    mo_cmbDepartamento.BoundColumn = "IdDepartamento"
    mo_cmbDepartamento.ListField = "DescripcionLarga"
    Set mo_cmbDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Apariencia = Nothing
    Set mr_ReglasHIS = Nothing
    Set mr_ReglasComunes = Nothing
    Set oRcs_Servicios = Nothing
    Set oRcs_ServiciosEstablecimiento = Nothing
    Set oRcs_Establecimientos = Nothing
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim mb_resiltado As Boolean
mb_resiltado = True
ms_mesajeError = ""

If oRcs_ServiciosEstablecimiento.RecordCount = 0 Then
    ms_mesajeError = vbCrLf & "No se encuentra ningun servicio, verifique las opciones."
    mb_resiltado = False
End If

If ml_IdEstablecimiento = 0 Then
    ms_mesajeError = ms_mesajeError & vbCrLf & "No ha elegido ningun establecimiento."
    mb_resiltado = False
End If

If Len(ms_mesajeError) <> 0 Then
    MsgBox "Se encontraron las siguientes inconsistencias:" & vbCrLf & ms_mesajeError, vbCritical, "HIS SIGH"
End If
ValidarDatosObligatorios = mb_resiltado
End Function

Function ValidarReglas()
    Dim mb_ValidacionReglas As Boolean
    mb_ValidacionReglas = True
    If mi_Opcion = sghAgregar Then
        If mr_ReglasHIS.ConsultarEstablecimiento(ml_IdEstablecimiento) = True Then
            MsgBox "El establecimiento ya fue ingresado", vbExclamation, Me.Caption
            mb_ValidacionReglas = False
        End If
    End If
    If mi_Opcion = sghEliminar Then
        Dim orsTemp As New ADODB.Recordset
        Set orsTemp = mr_ReglasHIS.HisLotesXEstablecimientos(ml_IdEstablecimiento)
        If orsTemp.RecordCount > 0 Then
            MsgBox "No puede eliminar el establecimiento porque ya tiene lotes asignados", vbExclamation, Me.Caption
        End If
    End If
    ValidarReglas = mb_ValidacionReglas
End Function

Function ActualizarDatos() As Boolean
If mi_Opcion = sghEliminar Then
    oRcs_ServiciosEstablecimiento.MoveFirst
    While Not oRcs_ServiciosEstablecimiento.EOF
        oRcs_ServiciosEstablecimiento!IdEstado = 3
        oRcs_ServiciosEstablecimiento.MoveNext
    Wend
End If

ActualizarDatos = mr_ReglasHIS.ActualizarServiciosPorEstablecimientos(ml_IdEstablecimiento, mi_Opcion, oRcs_ServiciosEstablecimiento)
End Function

'Devuelve V o F si se encuantra un Id de servicio en el establecimiento del Digitador
Private Function BuscarServicioActual(ml_IdServicio As Long) As Boolean
Dim mb_BuscarServicioLocal As Boolean
mb_BuscarServicioLocal = False

If oRcs_ServiciosEstablecimiento.RecordCount <> 0 Then
    oRcs_ServiciosEstablecimiento.MoveFirst
    While Not oRcs_ServiciosEstablecimiento.EOF
        If oRcs_ServiciosEstablecimiento!IdServicio = ml_IdServicio Then
            mb_BuscarServicioLocal = True
            oRcs_ServiciosEstablecimiento.MoveLast
        End If
        oRcs_ServiciosEstablecimiento.MoveNext
    Wend
End If
BuscarServicioActual = mb_BuscarServicioLocal
End Function

Private Function VerificarRegistrosServicio() As Boolean
Dim mo_Resultado As Boolean
mo_Resultado = False
If Me.ugvServiciosPorEstablecimiento.Selected.Rows.Count <> 0 Then
    mo_Resultado = True
End If
VerificarRegistrosServicio = mo_Resultado
End Function
