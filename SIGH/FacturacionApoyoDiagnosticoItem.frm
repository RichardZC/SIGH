VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturacionApoyoDiagnosticoItem 
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   Icon            =   "FacturacionApoyoDiagnosticoItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnBusquedaProcedimiento 
      Caption         =   ".."
      Height          =   315
      Left            =   2820
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   345
   End
   Begin VB.CommandButton btnBusquedaServicio 
      Caption         =   "..."
      Height          =   315
      Left            =   2805
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   615
      Width           =   345
   End
   Begin VB.Frame fraProcedimiento 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   90
      TabIndex        =   9
      Top             =   -15
      Width           =   7965
      Begin VB.TextBox lblDescServicio 
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   615
         Width           =   4695
      End
      Begin VB.TextBox lblDescProcedimiento 
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
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
         Left            =   1680
         TabIndex        =   2
         Top             =   615
         Width           =   975
      End
      Begin VB.TextBox txtIdProcedimiento 
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
      Begin MSMask.MaskEdBox txtHoraRealizacion 
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Top             =   990
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
      Begin MSMask.MaskEdBox txtFechaRealizacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   990
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
      Begin VB.Label Label63 
         Caption         =   "Servicio realiza"
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
         TabIndex        =   14
         Top             =   645
         Width           =   1425
      End
      Begin VB.Label Label65 
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
         Left            =   150
         TabIndex        =   13
         Top             =   1005
         Width           =   1530
      End
      Begin VB.Label Label69 
         Caption         =   "Exámen"
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
         TabIndex        =   12
         Top             =   300
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   90
      TabIndex        =   8
      Top             =   1425
      Width           =   7965
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacturacionApoyoDiagnosticoItem.frx":0CCA
         DownPicture     =   "FacturacionApoyoDiagnosticoItem.frx":118E
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
         Left            =   4065
         Picture         =   "FacturacionApoyoDiagnosticoItem.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacturacionApoyoDiagnosticoItem.frx":1B66
         DownPicture     =   "FacturacionApoyoDiagnosticoItem.frx":1FC6
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
         Left            =   2520
         Picture         =   "FacturacionApoyoDiagnosticoItem.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FacturacionApoyoDiagnosticoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mo_Teclado As New SIGHCOmun.Teclado
Dim mo_Formulario As New SIGHCOmun.Formulario
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mi_Opcion As Integer
Dim mrs_CurrentRecordset As ADODB.Recordset
Dim ml_IdDepartamentoHospital As Long
Dim ml_FechaOrden As Date

Property Let IdDepartamentoHospital(lValue As Long)
    ml_IdDepartamentoHospital = lValue
End Property
Property Get IdDepartamentoHospital() As Long
    IdDepartamentoHospital = ml_IdDepartamentoHospital
End Property
Property Let Opcion(lValue As Long)
    mi_Opcion = lValue
End Property
Property Get Opcion() As Long
    Opcion = mi_Opcion
End Property
Property Set CurrentRecorset(oValue As Recordset)
    Set mrs_CurrentRecordset = oValue
End Property
Property Get CurrentRecorset() As Recordset
    Set CurrentRecorset = mrs_CurrentRecordset
End Property
Property Let FechaOrden(daValue As Date)
    ml_FechaOrden = daValue
End Property
Property Get FechaOrden() As Date
    FechaOrden = ml_FechaOrden
End Property
Private Sub btnAceptar_Click()
    
    If Me.txtIdProcedimiento = "" Then
        MsgBox "Ingrese el código de examen", vbExclamation, Me.Caption
        Exit Sub
    End If

    If Me.txtIdServicio = "" Then
        MsgBox "Ingrese el servicio que realiza el examen", vbExclamation, Me.Caption
        Exit Sub
    End If

    If Me.FechaOrden <> 0 Then
        If CDate(Me.txtFechaRealizacion + " " + Me.txtHoraRealizacion) < Me.FechaOrden Then
            MsgBox "La fecha de resultado no puede ser menor que la fecha de orden del examen", vbExclamation, Me.Caption
            Exit Sub
        End If
    End If
    
    Select Case mi_Opcion
    Case sghAgregar
        With mrs_CurrentRecordset
            .AddNew
            .Fields!IdProcedimiento = Val(Me.txtIdProcedimiento.Tag)
            .Fields!CodigoCPT = Me.txtIdProcedimiento
            .Fields!Descripcion = Me.lblDescProcedimiento
            .Fields!IdServicioRealiza = Val(Me.txtIdServicio.Tag)
            .Fields!NombreServicio = Me.lblDescServicio
            .Fields!FechaResultado = Me.txtFechaRealizacion
            .Fields!HoraResultado = Me.txtHoraRealizacion
            .Fields!IdFacturacionServicio = 0
        End With
    Case sghModificar
        With mrs_CurrentRecordset
            .Fields!IdProcedimiento = Val(Me.txtIdProcedimiento.Tag)
            .Fields!CodigoCPT = Me.txtIdProcedimiento
            .Fields!Descripcion = Me.lblDescProcedimiento
            .Fields!IdServicioRealiza = Val(Me.txtIdServicio.Tag)
            .Fields!NombreServicio = Me.lblDescServicio
            .Fields!FechaResultado = Me.txtFechaRealizacion
            .Fields!HoraResultado = Me.txtHoraRealizacion
            '.Fields!IdFacturacionServicio vuelve con el mismo valor que tiene
            .Update
        End With
    Case sghEliminar
            mrs_CurrentRecordset.Delete
            mrs_CurrentRecordset.Update
    End Select

    Me.Visible = False

End Sub

Private Sub btnBusquedaServicio_Click()
    CompletarDatosDeServicio txtIdServicio, lblDescServicio
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
    
    mo_Formulario.HabilitarDeshabilitar Me.lblDescProcedimiento, False
    mo_Formulario.HabilitarDeshabilitar Me.lblDescServicio, False
    
    Select Case mi_Opcion
    Case sghAgregar
    Case sghModificar, sghConsultar, sghEliminar
        
        Me.txtFechaRealizacion = mrs_CurrentRecordset!FechaResultado
        Me.txtHoraRealizacion = mrs_CurrentRecordset!HoraResultado
        
        
        Me.txtIdServicio.Tag = mrs_CurrentRecordset!IdServicioRealiza
        Dim oDOServicio As New DOServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(mrs_CurrentRecordset!IdServicioRealiza)
        If Not oDOServicio Is Nothing Then
            Me.txtIdServicio.Text = oDOServicio.Codigo
            Me.lblDescServicio = oDOServicio.Nombre
        End If
        
        Me.txtIdProcedimiento.Tag = Val(mrs_CurrentRecordset!IdProcedimiento)
        Dim oDOProcedimiento As DOProcedimiento
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorId(mrs_CurrentRecordset!IdProcedimiento)
        If Not oDOProcedimiento Is Nothing Then
            txtIdProcedimiento.Text = oDOProcedimiento.CodigoCPT2004
            lblDescProcedimiento = oDOProcedimiento.Descripcion
        End If
        
        Select Case mi_Opcion
        Case sghModificar
            If mrs_CurrentRecordset!IdFacturacionServicio <> 0 Then
                MsgBox "Este procedimiento ya ha sido facturado, solo se puede modificar algunos datos", vbInformation, Me.Caption
                mo_Formulario.HabilitarDeshabilitar Me.txtIdProcedimiento, False
            End If
        Case sghConsultar
            Me.btnAceptar.Enabled = False
        Case sghEliminar
            If mrs_CurrentRecordset!IdFacturacionServicio <> 0 Then
                MsgBox "Este procedimiento ya ha sido facturado, no se puede eliminar", vbInformation, Me.Caption
                Me.btnAceptar.Enabled = False
            End If
        End Select
    End Select

    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar examen"
    Case sghModificar
        Me.Caption = "Modificar examen"
    Case sghConsultar
        Me.Caption = "Consultar examen"
    Case sghEliminar
        Me.Caption = "Eliminar examen"
    End Select


End Sub

Private Sub txtFechaRealizacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaRealizacion
End Sub
Private Sub txtFechaRealizacion_LostFocus()
    If txtFechaRealizacion <> SIGHCOmun.FECHA_VACIA_DMY Then
         If Not EsFecha(txtFechaRealizacion, "DD/MM/AAAA") Then
             MsgBox "La fecha ingresada no es válida", vbInformation, "Procedimientos"
              txtFechaRealizacion = SIGHCOmun.FECHA_VACIA_DMY
         End If
     End If
   'mo_Formulario.MarcarComoVacio txtFechaRealizacion
End Sub

Private Sub txtFechaRealizacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraRealizacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraRealizacion
    
End Sub


Private Sub txtHoraRealizacion_LostFocus()
        
    If txtHoraRealizacion <> SIGHCOmun.HORA_VACIA_HM Then
        If Not SIGHCOmun.ValidaHora(txtHoraRealizacion) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, "Registro de procedimientos"
             txtHoraRealizacion = SIGHCOmun.FECHA_VACIA_DMY
        End If
    End If
        
    'mo_Formulario.MarcarComoVacio txtHoraRealizacion
End Sub

Private Sub txtHoraRealizacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicio
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicio_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicio, lblDescServicio
    mo_Formulario.MarcarComoVacio txtIdServicio
End Sub

Private Sub txtIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub btnBusquedaProcedimiento_Click()
Dim oBusqueda As New ProcedimientosBusqueda
Dim oDOProcedimiento As DOProcedimiento

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOProcedimiento Is Nothing Then
            txtIdProcedimiento.Text = oDOProcedimiento.CodigoCPT2004
            txtIdProcedimiento.Tag = oDOProcedimiento.IdProcedimiento
            lblDescProcedimiento = oDOProcedimiento.Descripcion
        Else
            txtIdProcedimiento.Text = ""
            txtIdProcedimiento.Tag = ""
            lblDescProcedimiento = ""
        End If
    End If
    
End Sub

Private Sub txtIdProcedimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdProcedimiento
End Sub

Private Sub txtIdProcedimiento_LostFocus()

    txtIdProcedimiento.Text = UCase(txtIdProcedimiento.Text)

   If txtIdProcedimiento.Text <> "" Then
    Dim oDOProcedimiento As DOProcedimiento
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorCodigoCPT(txtIdProcedimiento.Text)
        If Not oDOProcedimiento Is Nothing Then
            txtIdProcedimiento.Tag = oDOProcedimiento.IdProcedimiento
            lblDescProcedimiento = oDOProcedimiento.Descripcion
        Else
            txtIdProcedimiento.Tag = ""
            lblDescProcedimiento = ""
        End If
   End If
   
   'mo_Formulario.MarcarComoVacio txtIdProcedimiento
End Sub

Private Sub txtIdProcedimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New ServiciosBusqueda
Dim oDOServicio As New DOServicio

    oBusqueda.IdTipoServicio = 5
    oBusqueda.IdDepartamentoHospital = ml_IdDepartamentoHospital
    oBusqueda.EjecutarBusquedaOnLoad = True
    oBusqueda.HabilitarTipoServicio = False
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOServicio Is Nothing Then
            txtIdServicio.Text = oDOServicio.Codigo
            txtIdServicio.Tag = oDOServicio.IdServicio
            lblDescripcionServicio = oDOServicio.Nombre
        End If
    End If

End Sub
Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDOServicio As DOServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDOServicio Is Nothing Then
            txtIdServicio.Tag = oDOServicio.IdServicio
            lblDescripcionServicio = oDOServicio.Nombre
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
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
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
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

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


