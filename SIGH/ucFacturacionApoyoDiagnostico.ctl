VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucFacturacionApoyoDiagnostico 
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   LockControls    =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   8055
   Begin VB.CommandButton btnBusquedaMedico 
      Caption         =   "..."
      Height          =   315
      Left            =   2730
      TabIndex        =   11
      Top             =   600
      Width           =   345
   End
   Begin VB.Frame fraProcedimiento 
      Caption         =   "Exámenes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7965
      Begin VB.TextBox txtNroOrden 
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
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtIdMedicoOrdena 
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
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox lblDescMedicoOrdena 
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
         Left            =   3120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   600
         Width           =   4695
      End
      Begin MSMask.MaskEdBox txtHoraOrden 
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   960
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
      Begin MSMask.MaskEdBox txtFechaOrden 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   1380
         _ExtentX        =   2434
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
      Begin VB.Label Label69 
         Caption         =   "Orden Nro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label Label66 
         Caption         =   "Médico ordena"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label Label65 
         Caption         =   "Fecha orden"
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
         TabIndex        =   8
         Top             =   1020
         Width           =   1530
      End
   End
   Begin MSComctlLib.Toolbar toolProcedimientos 
      Height          =   540
      Left            =   60
      TabIndex        =   4
      Top             =   1500
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   953
      ButtonWidth     =   1402
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "lstOpciones"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "AGREGAR"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Key             =   "MODIFICAR"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "CONSULTAR"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            Key             =   "ELIMINAR"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin UltraGrid.SSUltraGrid grdExamenes 
      Height          =   2835
      Left            =   930
      TabIndex        =   5
      Top             =   1500
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   5001
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
      Caption         =   "Lista de exámenes"
   End
   Begin MSComctlLib.ImageList lstOpciones 
      Left            =   0
      Top             =   3690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucFacturacionApoyoDiagnostico.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucFacturacionApoyoDiagnostico.ctx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucFacturacionApoyoDiagnostico.ctx":08EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucFacturacionApoyoDiagnostico.ctx":0D06
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ucFacturacionApoyoDiagnostico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mi_Opcion As sghOpciones
Dim ml_IdCuentaAtencion As Long
Dim ml_IdUsuario As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ms_MensajeError As String
Dim mrs_ExamenesApoyoDx As New ADODB.Recordset
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ml_IdTipoServicio As Long
Dim ml_IdPreFacturacionApoyoDx As Long
Dim mda_FechaIngreso As Date
Dim ml_IdDepartamentoHospital As Long
Public Event SePresionoTeclaEspecial(KeyCode As Integer)

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
    Select Case mi_Opcion
    Case sghAgregar
    Case sghModificar
    Case sghConsultar
         toolProcedimientos.Buttons(1).Enabled = False
         toolProcedimientos.Buttons(2).Enabled = False
         toolProcedimientos.Buttons(4).Enabled = False
    Case sghEliminar
         toolProcedimientos.Buttons(1).Enabled = False
         toolProcedimientos.Buttons(2).Enabled = False
         toolProcedimientos.Buttons(4).Enabled = False
    End Select
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let TipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get TipoServicio() As Long
   TipoServicio = ml_IdTipoServicio
End Property
Property Let IdPreFacturacionApoyoDx(lValue As Long)
   ml_IdPreFacturacionApoyoDx = lValue
End Property
Property Get IdPreFacturacionApoyoDx() As Long
   IdPreFacturacionApoyoDx = ml_IdPreFacturacionApoyoDx
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let FechaIngreso(daValue As Date)
   mda_FechaIngreso = daValue
End Property
Property Get FechaOrden() As Date
        On Error Resume Next
        FechaOrden = CDate(txtFechaOrden.Text + " " + txtHoraOrden.Text)
End Property


Property Let IdDepartamentoHospital(lValue As Long)
    ml_IdDepartamentoHospital = lValue
End Property
Property Get IdDepartamentoHospital() As Long
    IdDepartamentoHospital = ml_IdDepartamentoHospital
End Property
Private Sub btnBusquedaMedico_Click()
       CompletarDatosDeMedico txtIdMedicoOrdena, UserControl.lblDescMedicoOrdena
End Sub


Private Sub toolProcedimientos_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim oDialog As New FacturacionApoyoDiagnosticoItem
'
'    oDialog.IdDepartamentoHospital = ml_IdDepartamentoHospital
'    Set oDialog.CurrentRecorset = mrs_ExamenesApoyoDx
'
'    Select Case Button.Key
'    Case "AGREGAR"
'        oDialog.Opcion = sghAgregar
'    Case "MODIFICAR"
'        oDialog.Opcion = sghModificar
'    Case "CONSULTAR"
'        oDialog.Opcion = sghConsultar
'    Case "ELIMINAR"
'        oDialog.Opcion = sghEliminar
'    End Select
'
'    On Error Resume Next
'    oDialog.FechaOrden = CDate(UserControl.txtFechaOrden + " " + UserControl.txtHoraOrden)
'    oDialog.Show 1

End Sub

Private Sub txtFechaOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaOrden
End Sub

Private Sub txtNroOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroOrden
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtNroOrden_LostFocus()
    mo_Formulario.MarcarComoVacio txtNroOrden
End Sub

Private Sub txtNroOrden_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Public Function Inicializar()
    GenerarRecordsetTemporal
    mo_Formulario.HabilitarDeshabilitar UserControl.lblDescMedicoOrdena, False
End Function


Private Sub grdExamenes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdExamenes.Bands(0).Columns("IdPreFactApoyoDetalle").Hidden = True
    grdExamenes.Bands(0).Columns("IdProcedimiento").Hidden = True
    grdExamenes.Bands(0).Columns("IdFacturacionServicio").Hidden = True
    grdExamenes.Bands(0).Columns("IdServicioRealiza").Hidden = True
    
    grdExamenes.Bands(0).Columns("CodigoCPT").Header.Caption = "CPT"
    grdExamenes.Bands(0).Columns("CodigoCPT").Width = 1000
    
    grdExamenes.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdExamenes.Bands(0).Columns("Descripcion").Width = 5000
    
    grdExamenes.Bands(0).Columns("FechaResultado").Header.Caption = "Fecha"
    grdExamenes.Bands(0).Columns("FechaResultado").Width = 1500
    
    grdExamenes.Bands(0).Columns("HoraResultado").Header.Caption = "Hora"
    grdExamenes.Bands(0).Columns("HoraResultado").Width = 1000
    
    grdExamenes.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
    grdExamenes.Bands(0).Columns("NombreServicio").Width = 2500


End Sub

Public Sub CargarDatosDeDeProcedimientos()
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDOPreFacturacionApoyoDx As New DOAtencionApoyoDiagnostico

    'Carga datos de la cabecera
    Dim rsProcedimiento As New Recordset
    Set oDOPreFacturacionApoyoDx = mo_AdminFacturacion.AtencionApoyoDxSeleccionarPorId(Me.IdPreFacturacionApoyoDx)
    UserControl.txtFechaOrden = oDOPreFacturacionApoyoDx.FechaOrden
    UserControl.txtHoraOrden = oDOPreFacturacionApoyoDx.HoraOrden
    UserControl.txtNroOrden = oDOPreFacturacionApoyoDx.OrdenNro
    
    'Completa datos de medico
    If mo_AdminProgramacion.MedicosSeleccionarPorId(oDOPreFacturacionApoyoDx.IdMedicoOrdena, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
        txtIdMedicoOrdena.Text = oDOEmpleado.CodigoPlanilla
        txtIdMedicoOrdena.Tag = oDoMedico.IdMedico
        lblDescMedicoOrdena = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    
  
    Dim rsProcedimientos As New Recordset
    Set rsProcedimientos = mo_AdminFacturacion.AtencionApoyoDxDetalleSeleccionarPorIdAtencionApoyoDx(Me.IdPreFacturacionApoyoDx)
    Do While Not rsProcedimientos.EOF
        With mrs_ExamenesApoyoDx
            .AddNew
            .Fields!IdPreFactApoyoDetalle = rsProcedimientos!IdPreFactApoyoDetalle
            .Fields!IdProcedimiento = rsProcedimientos!IdProcedimiento
            .Fields!CodigoCPT = rsProcedimientos!CodigoCPT
            .Fields!descripcion = rsProcedimientos!descripcion
            .Fields!IdServicioRealiza = rsProcedimientos!IdServicioRealiza
            .Fields!NombreServicio = rsProcedimientos!NombreServicio
            .Fields!FechaResultado = Format(rsProcedimientos!FechaResultado, "dd/mm/yyyy")
            .Fields!HoraResultado = rsProcedimientos!HoraResultado
            .Fields!IdFacturacionServicio = rsProcedimientos!IdFacturacionServicio
        End With
        rsProcedimientos.MoveNext
    Loop
    rsProcedimientos.Close
    mo_Apariencia.ConfigurarFilasBiColores UserControl.grdExamenes, SIGHComun.GrillaConFilasBicolor
    
End Sub

Sub CargarProcedimientosAlObjetoDatos(oApoyoDx As DOAtencionApoyoDiagnostico, oListaApoyoDxDetalle As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ProcedimientoS
    '---------------------------------------------------------------------------------
    'Datos de la cabecera
    oApoyoDx.IdAtencionApoyoDx = Me.IdPreFacturacionApoyoDx
    oApoyoDx.IdCuentaAtencion = ml_IdCuentaAtencion
    oApoyoDx.IdMedicoOrdena = Val(txtIdMedicoOrdena.Tag)
    oApoyoDx.FechaOrden = txtFechaOrden.Text
    oApoyoDx.HoraOrden = txtHoraOrden.Text
    oApoyoDx.OrdenNro = txtNroOrden.Text
    oApoyoDx.IdUsuarioAuditoria = ml_IdUsuario
    
    'Datos del detalle
    Dim oFactApoyoDxDetalle As DOAtencionApoyoDiagDetalle
    If Not (mrs_ExamenesApoyoDx.BOF And mrs_ExamenesApoyoDx.EOF) Then
        Set oFactApoyoDxDetalle = New DOAtencionApoyoDiagDetalle
        mrs_ExamenesApoyoDx.MoveFirst
        Do While Not mrs_ExamenesApoyoDx.EOF
            Set oFactApoyoDxDetalle = New DOAtencionApoyoDiagDetalle
            
            oFactApoyoDxDetalle.IdAtencionApoyoDetalle = mrs_ExamenesApoyoDx!IdPreFactApoyoDetalle
            oFactApoyoDxDetalle.IdAtencionApoyoDx = Me.IdPreFacturacionApoyoDx
            oFactApoyoDxDetalle.FechaResultado = mrs_ExamenesApoyoDx!FechaResultado
            oFactApoyoDxDetalle.HoraResultado = mrs_ExamenesApoyoDx!HoraResultado
            oFactApoyoDxDetalle.IdProcedimiento = mrs_ExamenesApoyoDx!IdProcedimiento
            oFactApoyoDxDetalle.IdServicioRealiza = mrs_ExamenesApoyoDx!IdServicioRealiza
            oFactApoyoDxDetalle.IdUsuarioAuditoria = ml_IdUsuario
            oFactApoyoDxDetalle.IdFacturacionServicio = IIf(IsNull(mrs_ExamenesApoyoDx!IdFacturacionServicio), 0, mrs_ExamenesApoyoDx!IdFacturacionServicio)
            
            oListaApoyoDxDetalle.Add oFactApoyoDxDetalle
            mrs_ExamenesApoyoDx.MoveNext
        Loop
    End If
    
End Sub
Sub GenerarRecordsetTemporal()
    
    With mrs_ExamenesApoyoDx
        .Fields.Append "IdPreFactApoyoDetalle", adInteger
        .Fields.Append "IdProcedimiento", adInteger
        .Fields.Append "CodigoCPT", adVarChar, 10
        .Fields.Append "Descripcion", adVarChar, 255
        .Fields.Append "FechaResultado", adChar, 10
        .Fields.Append "HoraResultado", adChar, 5
        .Fields.Append "IdServicioRealiza", adInteger
        .Fields.Append "NombreServicio", adVarChar, 100
        .Fields.Append "IdFacturacionServicio", adInteger, , adFldIsNullable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set UserControl.grdExamenes.DataSource = mrs_ExamenesApoyoDx
    
End Sub

Sub LimpiarDatos()
        
    txtIdMedicoOrdena.Text = ""
    txtIdMedicoOrdena.Tag = ""
    lblDescMedicoOrdena.Text = ""

    txtNroOrden = ""
    
    txtFechaOrden.Text = SIGHComun.FECHA_VACIA_DMY
    txtHoraOrden = SIGHComun.HORA_VACIA_HM

    Do While Not mrs_ExamenesApoyoDx.EOF
        mrs_ExamenesApoyoDx.Delete
        mrs_ExamenesApoyoDx.Update
        mrs_ExamenesApoyoDx.MoveNext
    Loop

End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.lblDescMedicoOrdena.Width = UserControl.Width - 3330
    fraProcedimiento.Width = UserControl.Width - 20
    UserControl.grdExamenes.Width = UserControl.Width - 965
    UserControl.grdExamenes.Height = UserControl.Height - 1570

End Sub

Private Sub txtIdMedicoOrdena_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoOrdena
    If KeyCode = vbKeyF1 Then
        btnBusquedaMedico_Click
    End If
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtIdMedicoOrdena_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoOrdena, UserControl.lblDescMedicoOrdena
    mo_Formulario.MarcarComoVacio txtIdMedicoOrdena
End Sub

Private Sub txtIdMedicoOrdena_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.IdMedico
            lblNombreMedico.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub

Sub CompletarDatosDeMedicoEnElLostFocus(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oMedicosEspecialidad As New Collection

    txtMedico = Trim(txtMedico)
    If txtMedico <> "" Then
        Dim oDOEmpleado As New dOEmpleado
        Dim oDoMedico As New DOMedico
        If mo_AdminProgramacion.MedicosSeleccionarPorCodigo(CStr(txtMedico), oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
            txtMedico.Tag = oDoMedico.IdMedico
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtMedico.Tag = ""
            lblNombreMedico = ""
        End If
    End If
    
End Sub

Private Sub btnBuscarMedicoSolicita_Click()
    CompletarDatosDeMedico txtIdMedicoOrdena, UserControl.lblDescMedicoOrdena
End Sub

Private Sub txtFechaSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaOrden
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFechaOrden_LostFocus()

       If txtFechaOrden <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaOrden, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
                 txtFechaOrden = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
        mo_Formulario.MarcarComoVacio txtFechaOrden
End Sub

Private Sub txtFechaOrden_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraOrden_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraOrden
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtHoraOrden_LostFocus()
   mo_Formulario.MarcarComoVacio txtHoraOrden
End Sub

Private Sub txtHoraOrden_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Function ValidarDatosObligatorios() As Boolean

    ValidarDatosObligatorios = False
    
    If UserControl.txtNroOrden = "" Then
        MsgBox "Ingrese el nro de orden de procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If UserControl.txtIdMedicoOrdena = "" Then
        MsgBox "Ingrese el médico que ordena el procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If UserControl.txtFechaOrden = SIGHComun.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha de orden del procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If UserControl.txtHoraOrden = SIGHComun.HORA_VACIA_HM Then
        MsgBox "Ingrese la hora de orden del procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    ValidarDatosObligatorios = True
End Function

