VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucFacturacionProcedimiento 
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   ScaleHeight     =   4395
   ScaleWidth      =   8205
   Begin MSComctlLib.ImageList lstOpciones 
      Left            =   120
      Top             =   3780
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
            Picture         =   "ucProcedimientoDetalle.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucProcedimientoDetalle.ctx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucProcedimientoDetalle.ctx":08EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucProcedimientoDetalle.ctx":0D06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolProcedimientos 
      Height          =   2160
      Left            =   30
      TabIndex        =   4
      Top             =   1620
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   3810
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
   Begin VB.CommandButton btnBusquedaMedico 
      Caption         =   "..."
      Height          =   315
      Left            =   2715
      TabIndex        =   6
      Top             =   600
      Width           =   345
   End
   Begin VB.Frame fraProcedimiento 
      Caption         =   "Procedimientos"
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
      TabIndex        =   8
      Top             =   0
      Width           =   7965
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
         TabIndex        =   11
         Top             =   1020
         Width           =   1530
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
         TabIndex        =   10
         Top             =   630
         Width           =   1350
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
         TabIndex        =   9
         Top             =   300
         Width           =   1260
      End
   End
   Begin UltraGrid.SSUltraGrid grdProcedimientos 
      Height          =   2835
      Left            =   960
      TabIndex        =   5
      Top             =   1470
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
      Caption         =   "Lista de procedimientos"
   End
End
Attribute VB_Name = "ucFacturacionProcedimiento"
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
Dim mrs_Procedimientos As New ADODB.Recordset
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ml_IdTipoServicio As Long
Dim ml_IdPreFacturacionProcedimiento As Long
Dim mda_FechaIngreso As Date
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
Property Let IdPreFacturacionProcedimiento(lValue As Long)
   ml_IdPreFacturacionProcedimiento = lValue
End Property
Property Get IdPreFacturacionProcedimiento() As Long
   IdPreFacturacionProcedimiento = ml_IdPreFacturacionProcedimiento
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

Private Sub btnBusquedaMedico_Click()
       CompletarDatosDeMedico txtIdMedicoOrdena, UserControl.lblDescMedicoOrdena
End Sub

Private Sub toolProcedimientos_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim oDialog As New FacturacionProcedimientoItem
'
'    Set oDialog.CurrentRecorset = mrs_Procedimientos
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

Private Sub txtIdProcedimiento_Change()

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


Private Sub grdProcedimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdProcedimientos.Bands(0).Columns("IdPreFacturacionProcDetalle").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdProcedimiento").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdFacturacionServicio").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdMedicoRealiza").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdServicioRealiza").Hidden = True
    
    grdProcedimientos.Bands(0).Columns("CodigoCPT").Header.Caption = "CPT"
    grdProcedimientos.Bands(0).Columns("CodigoCPT").Width = 1000
    
    grdProcedimientos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdProcedimientos.Bands(0).Columns("Descripcion").Width = 5000
    
    grdProcedimientos.Bands(0).Columns("FechaRealizacion").Header.Caption = "Fecha"
    grdProcedimientos.Bands(0).Columns("FechaRealizacion").Width = 1500
    
    grdProcedimientos.Bands(0).Columns("HoraRealizacion").Header.Caption = "Hora"
    grdProcedimientos.Bands(0).Columns("HoraRealizacion").Width = 1000
    
    grdProcedimientos.Bands(0).Columns("NombreMedico").Header.Caption = "Médico"
    grdProcedimientos.Bands(0).Columns("NombreMedico").Width = 3000

    grdProcedimientos.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
    grdProcedimientos.Bands(0).Columns("NombreServicio").Width = 2500


End Sub

Public Sub CargarDatosDeDeProcedimientos()
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDOPreFacturacionProcedimiento As New DOAtencionProcedimiento

    'Carga datos de la cabecera
    Dim rsProcedimiento As New Recordset
    Set oDOPreFacturacionProcedimiento = mo_AdminFacturacion.AtencionProcedimientosSeleccionarPorId(Me.IdPreFacturacionProcedimiento)
    UserControl.txtFechaOrden = oDOPreFacturacionProcedimiento.FechaOrden
    UserControl.txtHoraOrden = oDOPreFacturacionProcedimiento.HoraOrden
    UserControl.txtNroOrden = oDOPreFacturacionProcedimiento.NroOrden
    
    'Completa datos de medico
    If mo_AdminProgramacion.MedicosSeleccionarPorId(oDOPreFacturacionProcedimiento.IdMedicoOrdena, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
        txtIdMedicoOrdena.Text = oDOEmpleado.CodigoPlanilla
        txtIdMedicoOrdena.Tag = oDoMedico.IdMedico
        lblDescMedicoOrdena = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    
  
    Dim rsProcedimientos As New Recordset
    Set rsProcedimientos = mo_AdminFacturacion.AtencionProcedimientoDetalleSeleccionarPorIdPreFacturacionProcedimiento(Me.IdPreFacturacionProcedimiento)
    Do While Not rsProcedimientos.EOF
        With mrs_Procedimientos
            .AddNew
            .Fields!IdPreFacturacionProcDetalle = rsProcedimientos!IdPreFacturacionProcDetalle
            .Fields!IdProcedimiento = rsProcedimientos!IdProcedimiento
            .Fields!CodigoCPT = rsProcedimientos!CodigoCPT
            .Fields!descripcion = rsProcedimientos!descripcion
            .Fields!IdMedicoRealiza = rsProcedimientos!IdMedicoRealiza
            .Fields!NombreMedico = rsProcedimientos!NombreMedico
            .Fields!IdServicioRealiza = rsProcedimientos!IdServicioRealiza
            .Fields!NombreServicio = rsProcedimientos!NombreServicio
            .Fields!FechaRealizacion = Format(rsProcedimientos!FechaRealizacion, "dd/mm/yyyy")
            .Fields!HoraRealizacion = rsProcedimientos!HoraRealizacion
            .Fields!IdFacturacionServicio = rsProcedimientos!IdFacturacionServicio
        End With
        rsProcedimientos.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores UserControl.grdProcedimientos, SIGHComun.GrillaConFilasBicolor
    
End Sub

Sub CargarProcedimientosAlObjetoDatos(oProcedimiento As DOAtencionProcedimiento, oProcedimientoDetalle As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ProcedimientoS
    '---------------------------------------------------------------------------------
    'Datos de la cabecera
    oProcedimiento.IdAtencionProcedimiento = Me.IdPreFacturacionProcedimiento
    oProcedimiento.IdCuentaAtencion = ml_IdCuentaAtencion
    oProcedimiento.IdMedicoOrdena = Val(txtIdMedicoOrdena.Tag)
    oProcedimiento.FechaOrden = txtFechaOrden.Text
    oProcedimiento.HoraOrden = txtHoraOrden.Text
    oProcedimiento.NroOrden = txtNroOrden.Text
    oProcedimiento.IdUsuarioAuditoria = ml_IdUsuario
    
    'Datos del detalle
    Dim oFacturacionProcDetalle As DOAtencionProcDetalle
    If Not (mrs_Procedimientos.BOF And mrs_Procedimientos.EOF) Then
        Set oFacturacionProcDetalle = New DOAtencionProcDetalle
        mrs_Procedimientos.MoveFirst
        Do While Not mrs_Procedimientos.EOF
            Set oFacturacionProcDetalle = New DOAtencionProcDetalle
            
            oFacturacionProcDetalle.IdAtencionProcDetalle = mrs_Procedimientos!IdPreFacturacionProcDetalle
            oFacturacionProcDetalle.IdAtencionProcedimiento = Me.IdPreFacturacionProcedimiento
            oFacturacionProcDetalle.FechaRealizacion = IIf(mrs_Procedimientos!FechaRealizacion <> "__/__/____", mrs_Procedimientos!FechaRealizacion, 0)
            oFacturacionProcDetalle.HoraRealizacion = mrs_Procedimientos!HoraRealizacion
            oFacturacionProcDetalle.IdMedicoRealiza = IIf(IsNull(mrs_Procedimientos!IdMedicoRealiza), 0, mrs_Procedimientos!IdMedicoRealiza)
            oFacturacionProcDetalle.IdProcedimiento = mrs_Procedimientos!IdProcedimiento
            oFacturacionProcDetalle.IdServicioRealiza = IIf(IsNull(mrs_Procedimientos!IdServicioRealiza), 0, mrs_Procedimientos!IdServicioRealiza)
            oFacturacionProcDetalle.IdUsuarioAuditoria = ml_IdUsuario
            oFacturacionProcDetalle.IdFacturacionServicio = IIf(IsNull(mrs_Procedimientos!IdFacturacionServicio), 0, mrs_Procedimientos!IdFacturacionServicio)
            
            oProcedimientoDetalle.Add oFacturacionProcDetalle
            mrs_Procedimientos.MoveNext
        Loop
    End If
    
End Sub
Sub GenerarRecordsetTemporal()
    
    With mrs_Procedimientos
        .Fields.Append "IdPreFacturacionProcDetalle", adInteger
        .Fields.Append "IdProcedimiento", adInteger
        .Fields.Append "CodigoCPT", adVarChar, 10
        .Fields.Append "Descripcion", adVarChar, 255
        .Fields.Append "FechaRealizacion", adChar, 10
        .Fields.Append "HoraRealizacion", adChar, 5
        .Fields.Append "IdMedicoRealiza", adInteger, , adFldIsNullable
        .Fields.Append "NombreMedico", adVarChar, 100, adFldIsNullable
        .Fields.Append "IdServicioRealiza", adInteger, , adFldIsNullable
        .Fields.Append "NombreServicio", adVarChar, 100, adFldIsNullable
        .Fields.Append "IdFacturacionServicio", adInteger, , adFldIsNullable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set UserControl.grdProcedimientos.DataSource = mrs_Procedimientos
    
End Sub

Sub LimpiarDatos()
        
    txtIdMedicoOrdena.Text = ""
    txtIdMedicoOrdena.Tag = ""
    lblDescMedicoOrdena.Text = ""

    txtNroOrden = ""
    
    txtFechaOrden.Text = SIGHComun.FECHA_VACIA_DMY
    txtHoraOrden = SIGHComun.HORA_VACIA_HM

    Do While Not mrs_Procedimientos.EOF
        mrs_Procedimientos.Delete
        mrs_Procedimientos.Update
        mrs_Procedimientos.MoveNext
    Loop

End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.lblDescMedicoOrdena.Width = UserControl.Width - 3330
    fraProcedimiento.Width = UserControl.Width - 20
    UserControl.grdProcedimientos.Width = UserControl.Width - 965
    UserControl.grdProcedimientos.Height = UserControl.Height - 1570

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
