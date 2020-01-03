VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucExamenDetalle 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   LockControls    =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   9300
   Begin VB.CommandButton btnBusquedaServicio 
      Caption         =   "..."
      Height          =   315
      Left            =   2580
      TabIndex        =   20
      Top             =   960
      Width           =   315
   End
   Begin VB.CommandButton btnBusquedaMedico 
      Caption         =   "..."
      Height          =   315
      Left            =   2580
      TabIndex        =   19
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton btnBusquedaExamen 
      Caption         =   ".."
      Height          =   315
      Left            =   2580
      TabIndex        =   18
      Top             =   240
      Width           =   315
   End
   Begin VB.Frame fraExamen 
      Height          =   2085
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9255
      Begin VB.TextBox lblNombreServicio 
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
         Left            =   2955
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   960
         Width           =   6195
      End
      Begin VB.TextBox lblNombreMedico 
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
         Left            =   2940
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   6210
      End
      Begin VB.TextBox lblDescExamen 
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
         Left            =   2940
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   6210
      End
      Begin VB.CommandButton btnAgregar 
         DisabledPicture =   "ucExamenDetalle.ctx":0000
         DownPicture     =   "ucExamenDetalle.ctx":03E9
         Height          =   315
         Left            =   7005
         Picture         =   "ucExamenDetalle.ctx":07F5
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1650
         Width           =   1005
      End
      Begin VB.CommandButton btnEliminar 
         DisabledPicture =   "ucExamenDetalle.ctx":0C01
         DownPicture     =   "ucExamenDetalle.ctx":0F8C
         Height          =   315
         Left            =   8070
         Picture         =   "ucExamenDetalle.ctx":131F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1650
         Width           =   1005
      End
      Begin VB.TextBox txtOrdenNro 
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
         Left            =   4665
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1320
         Width           =   1110
      End
      Begin VB.TextBox txtIdExamen 
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
         Left            =   1545
         TabIndex        =   0
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox txtIdMedico 
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
         Left            =   1545
         TabIndex        =   1
         Top             =   600
         Width           =   975
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
         Left            =   1545
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtHoraOrden 
         Height          =   315
         Left            =   2970
         TabIndex        =   4
         Top             =   1320
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
         Left            =   1545
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox txtHoraResultado 
         Height          =   315
         Left            =   2970
         TabIndex        =   6
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
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
      Begin MSMask.MaskEdBox txtFechaResultado 
         Height          =   315
         Left            =   1545
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label1 
         Caption         =   "Nº Orden"
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
         Left            =   3810
         TabIndex        =   17
         Top             =   1380
         Width           =   1035
      End
      Begin VB.Label Label45 
         Caption         =   "Fecha resultado"
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
         TabIndex        =   16
         Top             =   1710
         Width           =   1305
      End
      Begin VB.Label Label49 
         Caption         =   "Examen"
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
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label59 
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
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label60 
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
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label62 
         Caption         =   "Servicio ordena"
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
         TabIndex        =   12
         Top             =   960
         Width           =   1395
      End
   End
   Begin UltraGrid.SSUltraGrid grdExamenes 
      Height          =   3495
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6165
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
End
Attribute VB_Name = "ucExamenDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ml_IdCuentaAtencion As Long
Dim ml_IdUsuario As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ms_MensajeError As String
Dim mrs_Examenes As New ADODB.Recordset
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ml_IdTipoServicio As Long
Dim mda_FechaIngreso As Date
Public Event SePresionoTeclaEspecial(KeyCode As Integer)

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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let FechaIngreso(daValue As Date)
   mda_FechaIngreso = daValue
End Property

Property Let MedicoOrdena(lValue As Long)
Dim oMedicosEspecialidad As New Collection

    txtIdMedico.Tag = lValue
    
    Dim oDOEmpleado As New DOEmpleado
    Dim oDoMedico As New DOMedico
    If mo_AdminProgramacion.MedicosSeleccionarPorId(lValue, oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
        'Completa el nombre
        txtIdMedico.Tag = oDoMedico.IdMedico
        Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
        txtIdMedico.Text = oDOEmpleado.CodigoPlanilla
        lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
   
End Property
Property Get IdMedicoOrdena() As Long
   MedicoOrdena = Val(txtIdMedico.Tag)
End Property

Property Let HabilitarMedico(bValue As Boolean)
   mo_Formulario.HabilitarDeshabilitar txtIdMedico, bValue
   UserControl.btnBusquedaMedico.Enabled = bValue
End Property

Property Let ServicioOrdena(lValue As Long)
Dim oDOServicio As New DOServicio

        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(lValue)
        If Not oDOServicio Is Nothing Then
            If ml_IdTipoServicio = oDOServicio.IdTipoServicio Then
                txtIdServicio.Text = oDOServicio.Codigo
                txtIdServicio.Tag = oDOServicio.IdServicio
                lblNombreServicio = oDOServicio.Nombre
            End If
        End If
   
End Property


Private Sub btnBusquedaMedico_Click()
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New DOEmpleado
Dim oDOEspecialidades As New Collection

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
            txtIdMedico.Text = oDOEmpleado.CodigoPlanilla
            txtIdMedico.Tag = oDoMedico.IdMedico
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
End Sub

Private Sub btnBusquedaServicio_Click()
Dim oBusqueda As New ServiciosBusqueda
Dim oDOServicio As New DOServicio

    oBusqueda.IdTipoServicio = ml_IdTipoServicio
    oBusqueda.HabilitarTipoServicio = False
    
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOServicio Is Nothing Then
            If ml_IdTipoServicio = oDOServicio.IdTipoServicio Then
                txtIdServicio.Text = oDOServicio.Codigo
                txtIdServicio.Tag = oDOServicio.IdServicio
                lblNombreServicio = oDOServicio.Nombre
            Else
                MsgBox "El servicio seleccionado no pertenece a emergencia", vbInformation, "Exámenes"
                txtIdServicio.Text = ""
                txtIdServicio.Tag = ""
                lblNombreServicio = ""
            End If
        End If
    End If
End Sub

Private Sub txtFechaOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaOrden
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtFechaOrden_LostFocus()
    If txtFechaOrden <> SIGHComun.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFechaOrden, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, "Exámenes"
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

Private Sub txtFechaResultado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaResultado
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtFechaResultado_LostFocus()
       
       If txtFechaResultado <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaResultado, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Exámenes"
                 txtFechaResultado = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
   mo_Formulario.MarcarComoVacio txtFechaResultado
End Sub

Private Sub txtFechaResultado_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraResultado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraResultado
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtHoraResultado_LostFocus()
    
    If txtHoraResultado <> SIGHComun.HORA_VACIA_HM Then
        If Not SIGHComun.ValidaHora(txtHoraResultado) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, "Registro de exámenes"
             txtHoraResultado = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
    
    mo_Formulario.MarcarComoVacio txtHoraResultado
End Sub

Private Sub txtHoraResultado_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraOrden
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtHoraOrden_LostFocus()
        
    If txtHoraOrden <> SIGHComun.HORA_VACIA_HM Then
        If Not SIGHComun.ValidaHora(txtHoraOrden) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, "Registro de examenes"
             txtHoraOrden = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
    
    mo_Formulario.MarcarComoVacio txtHoraOrden
End Sub

Private Sub txtHoraOrden_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtIdMedico_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdMedico
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtIdMedico_LostFocus()
Dim oMedicosEspecialidad As New Collection

    'Busca nombre del medico
    If txtIdMedico <> "" Then
        Dim oDOEmpleado As New DOEmpleado
        Dim oDoMedico As New DOMedico
        If mo_AdminProgramacion.MedicosSeleccionarPorCodigo(Val(txtIdMedico), oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
            'Completa el nombre
            txtIdMedico.Tag = oDoMedico.IdMedico
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtIdMedico.Tag = ""
            lblNombreMedico = ""
        End If
    End If

   mo_Formulario.MarcarComoVacio txtIdMedico
End Sub

Private Sub txtIdMedico_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdServicio_LostFocus()

    txtIdServicio.Text = UCase(txtIdServicio.Text)

   If txtIdServicio.Text <> "" Then
    Dim oDOServicio As DOServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDOServicio Is Nothing Then
            If ml_IdTipoServicio = oDOServicio.IdTipoServicio Then
                txtIdServicio.Tag = oDOServicio.IdServicio
                lblNombreServicio = oDOServicio.Nombre
            Else
                MsgBox "El servicio ingresado no pertenece es de emergencia", vbInformation, "Exámenes"
                txtIdServicio.Tag = ""
                lblNombreServicio = ""
            End If
        Else
            txtIdServicio.Tag = ""
            lblNombreServicio = ""
        End If
   End If
   
   mo_Formulario.MarcarComoVacio txtIdServicio
End Sub

Private Sub txtIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicio
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtOrdenNro_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtOrdenNro
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtOrdenNro_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_Initialize()
    GenerarRecordsetTemporal
    mo_Formulario.HabilitarDeshabilitar UserControl.lblNombreMedico, False
    mo_Formulario.HabilitarDeshabilitar UserControl.lblNombreServicio, False
    mo_Formulario.HabilitarDeshabilitar UserControl.lblDescExamen, False
End Sub

Private Sub btnBusquedaExamen_Click()
Dim oBusqueda As New ProcedimientosBusqueda
Dim oDOProcedimiento As New DOProcedimiento

    
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOProcedimiento Is Nothing Then
            UserControl.txtIdExamen.Text = oDOProcedimiento.CodigoCPT2004
            UserControl.txtIdExamen.Tag = oDOProcedimiento.IdProcedimiento
            UserControl.lblDescExamen = oDOProcedimiento.Descripcion
        Else
            UserControl.txtIdExamen.Text = ""
            UserControl.txtIdExamen.Tag = ""
            UserControl.lblDescExamen = ""
        End If
    End If
    
End Sub

Private Sub grdExamenes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdExamenes.Bands(0).Columns("IdExamen").Hidden = True
    
    grdExamenes.Bands(0).Columns("CodigoCPT").Header.Caption = "CPT"
    grdExamenes.Bands(0).Columns("CodigoCPT").Width = 1000
    
    grdExamenes.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdExamenes.Bands(0).Columns("Descripcion").Width = 5000
    
    grdExamenes.Bands(0).Columns("IdMedicoOrdena").Hidden = True
    grdExamenes.Bands(0).Columns("IdServicioOrdena").Hidden = True
    grdExamenes.Bands(0).Columns("IdDetalleProducto").Hidden = True
    
    grdExamenes.Bands(0).Columns("NombreMedico").Header.Caption = "Médico"
    grdExamenes.Bands(0).Columns("NombreMedico").Width = 2000

    grdExamenes.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
    grdExamenes.Bands(0).Columns("NombreServicio").Width = 2000
    
    grdExamenes.Bands(0).Columns("OrdenNro").Header.Caption = "Orden Nro"
    grdExamenes.Bands(0).Columns("OrdenNro").Width = 1000
    
    grdExamenes.Bands(0).Columns("FechaOrden").Header.Caption = "Fecha Orden"
    grdExamenes.Bands(0).Columns("FechaOrden").Width = 2000
    
    grdExamenes.Bands(0).Columns("HoraOrden").Header.Caption = "Hora Orden"
    grdExamenes.Bands(0).Columns("HoraOrden").Width = 2000
    
    grdExamenes.Bands(0).Columns("FechaResultado").Header.Caption = "Fecha Result"
    grdExamenes.Bands(0).Columns("FechaResultado").Width = 2000
    
    grdExamenes.Bands(0).Columns("HoraResultado").Header.Caption = "Hora Result"
    grdExamenes.Bands(0).Columns("HoraResultado").Width = 2000
    


End Sub
Private Sub txtIdExamen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdExamen
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtIdExamen_LostFocus()

    UserControl.txtIdExamen.Text = UCase(UserControl.txtIdExamen.Text)

   If UserControl.txtIdExamen.Text <> "" Then
    Dim oDOExamen As DOProcedimiento
        Set oDOExamen = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorCodigoCPT(UserControl.txtIdExamen.Text)
        If Not oDOExamen Is Nothing Then
            UserControl.txtIdExamen.Tag = oDOExamen.IdProcedimiento
            UserControl.lblDescExamen = oDOExamen.Descripcion
        Else
            UserControl.txtIdExamen.Tag = ""
            UserControl.lblDescExamen = ""
        End If
   End If
   
   mo_Formulario.MarcarComoVacio txtIdExamen
End Sub

Private Sub txtIdExamen_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CargarDatosDeExamens()
Dim rsExamens As New Recordset

    Set rsExamens = mo_AdminAdmision.AtencionesExamenesSeleccionarPorCuentaAtencion(ml_IdCuentaAtencion)
    Do While Not rsExamens.EOF
        With mrs_Examenes
            .AddNew
            .Fields!IdExamen = rsExamens!IdExamen
            .Fields!CodigoCPT = rsExamens!CodigoCPT
            .Fields!Descripcion = rsExamens!Descripcion
            .Fields!IdMedicoOrdena = rsExamens!IdMedicoOrdena
            .Fields!NombreMedico = rsExamens!NombreMedico
            .Fields!IdServicioOrdena = rsExamens!IdServicioOrdena
            .Fields!NombreServicio = rsExamens!NombreServicio
            .Fields!OrdenNro = rsExamens!OrdenNro
            .Fields!FechaOrden = rsExamens!FechaOrden
            .Fields!HoraOrden = rsExamens!HoraOrden
            .Fields!FechaResultado = rsExamens!FechaResultado
            .Fields!HoraResultado = rsExamens!HoraResultado
            .Fields!IdDetalleProducto = rsExamens!IdDetalleProducto
        End With
        rsExamens.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores UserControl.grdExamenes, SIGHComun.GrillaConFilasBicolor
    
End Sub

Sub CargarExamensAlObjetoDatos(oExamens As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ExamenS
    '---------------------------------------------------------------------------------
    Dim oExamen As DOAtencionExamen
    
    If Not (mrs_Examenes.BOF And mrs_Examenes.EOF) Then
        mrs_Examenes.MoveFirst
        Do While Not mrs_Examenes.EOF
            Set oExamen = New DOAtencionExamen
            oExamen.IdCuentaAtencion = 0
            oExamen.IdCuentaAtencion = ml_IdCuentaAtencion
            oExamen.IdExamen = mrs_Examenes!IdExamen
            oExamen.IdMedicoOrdena = mrs_Examenes!IdMedicoOrdena
            oExamen.IdServicioOrdena = mrs_Examenes!IdServicioOrdena
            oExamen.OrdenNro = "" & mrs_Examenes!OrdenNro
            oExamen.FechaOrden = mrs_Examenes!FechaOrden
            oExamen.HoraOrden = mrs_Examenes!HoraOrden
            oExamen.FechaResultado = mrs_Examenes!FechaResultado
            oExamen.HoraResultado = mrs_Examenes!HoraResultado
            oExamen.IdDetalleProducto = IIf(IsNull(mrs_Examenes!IdDetalleProducto), 0, mrs_Examenes!IdDetalleProducto)
            oExamen.IdUsuarioAuditoria = ml_IdUsuario
            oExamens.Add oExamen
            mrs_Examenes.MoveNext
        Loop
    End If
End Sub
Sub GenerarRecordsetTemporal()
    
    With mrs_Examenes
          .Fields.Append "IdExamen", adInteger
          .Fields.Append "CodigoCPT", adVarChar, 10
          .Fields.Append "Descripcion", adVarChar, 255
          .Fields.Append "IdMedicoOrdena", adInteger
          .Fields.Append "NombreMedico", adVarChar, 100
          .Fields.Append "IdServicioOrdena", adInteger
          .Fields.Append "NombreServicio", adVarChar, 100
          .Fields.Append "OrdenNro", adVarChar, 10, adFldIsNullable
          .Fields.Append "FechaOrden", adChar, 10
          .Fields.Append "HoraOrden", adChar, 5
          .Fields.Append "FechaResultado", adChar, 10
          .Fields.Append "HoraResultado", adChar, 5
          .Fields.Append "IdDetalleProducto", adInteger, 4, adFldIsNullable
          
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set UserControl.grdExamenes.DataSource = mrs_Examenes
    
End Sub

Private Sub btnAgregar_Click()
        
    If txtIdExamen.Text = "" Then
        MsgBox "Por favor ingrese el código del examen", vbInformation, "Ingreso de examenes"
        Exit Sub
    End If
    
    If txtFechaOrden = SIGHComun.FECHA_VACIA_DMY Or txtHoraOrden = SIGHComun.HORA_VACIA_HM Then
        MsgBox "Por favor ingrese la fecha y hora de la orden", vbExclamation, "Ingreso de exámenes"
        Exit Sub
    End If
    
    If txtFechaResultado = SIGHComun.FECHA_VACIA_DMY Or txtHoraResultado = SIGHComun.HORA_VACIA_HM Then
        MsgBox "Por favor ingrese la fecha y hora de resultado", vbExclamation, "Ingreso de exámenes"
        Exit Sub
    End If
    
    If CDate(UserControl.txtFechaOrden + " " + txtHoraOrden) <= mda_FechaIngreso Then
        MsgBox "La fecha de orden del examen no puede ser menor que la fecha de ingreso", vbExclamation, "Ingreso de exámenes"
        Exit Sub
    End If

    If CDate(txtFechaResultado + " " + txtHoraResultado) < CDate(txtFechaOrden + " " + txtHoraOrden) Then
        MsgBox "La fecha de resultado del examen no puede ser menor que la fecha de orden", vbExclamation, "Ingreso de exámenes"
        Exit Sub
    End If

'    If CDate(UserControl.txtFechaOrden) > Date Then
'        MsgBox "La fecha de orden del examen no puede ser mayor que la fecha de hoy", vbExclamation, "Ingreso de exámenes"
'        Exit Sub
'    End If

'    If CDate(UserControl.txtFechaResultado) > Date Then
'        MsgBox "La fecha de resultado del examen no puede ser mayor que la fecha de hoy", vbExclamation, "Ingreso de exámenes"
'        Exit Sub
'    End If

    With mrs_Examenes
        .AddNew
        .Fields!IdExamen = UserControl.txtIdExamen.Tag
        .Fields!CodigoCPT = UserControl.txtIdExamen.Text
        .Fields!Descripcion = UserControl.lblDescExamen
        .Fields!IdMedicoOrdena = UserControl.txtIdMedico.Tag
        .Fields!NombreMedico = UserControl.lblNombreMedico.Text
        .Fields!IdServicioOrdena = UserControl.txtIdServicio.Tag
        .Fields!NombreServicio = UserControl.lblNombreServicio.Text
        .Fields!OrdenNro = UserControl.txtOrdenNro
        .Fields!FechaOrden = UserControl.txtFechaOrden
        .Fields!HoraOrden = UserControl.txtHoraOrden
        .Fields!FechaResultado = UserControl.txtFechaResultado
        .Fields!HoraResultado = UserControl.txtHoraResultado
        .Fields!IdDetalleProducto = 0
    End With
    LimpiarDatos

End Sub
Sub LimpiarDatos()
        
        txtIdExamen.Tag = ""
        txtIdExamen.Text = ""
        lblDescExamen = ""
        
        If Not txtIdMedico.Locked Then
            txtIdMedico.Text = ""
            txtIdMedico.Tag = ""
            lblNombreMedico.Text = ""
        End If
        
'        txtIdServicio.Text=""
'        txtIdServicio.Tag = ""
'        lblNombreServicio.Text = ""
        
        txtOrdenNro = ""
        txtFechaOrden = SIGHComun.FECHA_VACIA_DMY
        txtHoraOrden = SIGHComun.HORA_VACIA_HM
        txtFechaResultado = SIGHComun.FECHA_VACIA_DMY
        txtHoraResultado = SIGHComun.HORA_VACIA_HM

End Sub

Private Sub btnEliminar_Click()
    On Error Resume Next
    With mrs_Examenes
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
End Sub

Private Sub UserControl_Resize()
    UserControl.lblNombreMedico.Width = UserControl.Width - 2980
    UserControl.lblDescExamen.Width = UserControl.Width - 2980
    UserControl.lblNombreServicio.Width = UserControl.Width - 2980
    
    fraExamen.Width = UserControl.Width - 20
    UserControl.grdExamenes.Width = UserControl.Width - 20
    UserControl.grdExamenes.Height = UserControl.Height - 2200

End Sub
