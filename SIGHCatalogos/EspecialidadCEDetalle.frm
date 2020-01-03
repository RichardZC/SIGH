VERSION 5.00
Begin VB.Form EspecialidadDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnProdInterconsulta 
      Caption         =   "..."
      Height          =   315
      Left            =   3510
      TabIndex        =   6
      Top             =   1740
      Width           =   315
   End
   Begin VB.CommandButton btnProdConsulta 
      Caption         =   "..."
      Height          =   315
      Left            =   3510
      TabIndex        =   3
      Top             =   1350
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   90
      TabIndex        =   15
      Top             =   1110
      Width           =   8115
      Begin VB.ComboBox cmbTiempoPromedioAtencion 
         Height          =   315
         ItemData        =   "EspecialidadCEDetalle.frx":0000
         Left            =   2340
         List            =   "EspecialidadCEDetalle.frx":000C
         TabIndex        =   8
         Top             =   1020
         Width           =   1005
      End
      Begin VB.TextBox txtCodProductoInterconsulta 
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
         Left            =   2355
         TabIndex        =   5
         Top             =   630
         Width           =   1000
      End
      Begin VB.TextBox txtCodProductoConsulta 
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
         Left            =   2355
         TabIndex        =   2
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtProductoInterconsulta 
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
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   7
         Top             =   630
         Width           =   4170
      End
      Begin VB.TextBox txtProductoConsulta 
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
         Left            =   3780
         MaxLength       =   50
         TabIndex        =   4
         Top             =   240
         Width           =   4200
      End
      Begin VB.Label Label4 
         Caption         =   "Producto Interconsulta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   18
         Top             =   660
         Width           =   1965
      End
      Begin VB.Label Label2 
         Caption         =   "Producto Consulta"
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
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Tiempo Promedio de Atención"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   150
         TabIndex        =   16
         Top             =   1050
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   12
      Top             =   2700
      Width           =   8115
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EspecialidadCEDetalle.frx":0018
         DownPicture     =   "EspecialidadCEDetalle.frx":0478
         Height          =   700
         Left            =   2625
         Picture         =   "EspecialidadCEDetalle.frx":08ED
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EspecialidadCEDetalle.frx":0D62
         DownPicture     =   "EspecialidadCEDetalle.frx":1226
         Height          =   700
         Left            =   4140
         Picture         =   "EspecialidadCEDetalle.frx":1712
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1095
      Left            =   90
      TabIndex        =   11
      Top             =   0
      Width           =   8115
      Begin VB.TextBox txtNombre 
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
         Left            =   2355
         MaxLength       =   50
         TabIndex        =   1
         Top             =   630
         Width           =   5610
      End
      Begin VB.ComboBox cmbIdDepartamento 
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
         Left            =   2355
         TabIndex        =   0
         Top             =   240
         Width           =   5610
      End
      Begin VB.Label lblDescripcion 
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
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   1185
      End
   End
End
Attribute VB_Name = "EspecialidadDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Especialidades de un Servicio
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Especialidades As New DOEspecialidades

Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones

Dim ml_IdEspecialidad As Long

Dim mb_ExistenDatos As Boolean
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim ml_IdDepartamento As Long
Dim mo_cmbIdDepartamento As New SIGHEntidades.ListaDespleglable

'especialidadesce

Dim mo_Facturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_EspecialidadesCE As New DOEspecialidadCE
Dim ml_IdEspecialidadCE As Long
Dim ml_IdProductoConsulta As Long
Dim ml_IdProductoInterconsulta As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property


Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

    
       
       mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminServiciosHosp.MensajeError
       If sMensaje <> "" Then
           MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
       
              
       
        
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

Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property

Property Let IdEspecialidadCE(lValue As Long)
   ml_IdEspecialidadCE = lValue
End Property
Property Get IdEspecialidadCE() As Long
   IdEspecialidadCE = ml_IdEspecialidadCE
End Property

Private Sub btnProdConsulta_Click()

        Dim oFrm As New SIGHNegocios.BuscaServicio
        oFrm.MostrarFormulario
        If oFrm.idRegistroSeleccionado <> 0 Then
            Me.txtCodProductoConsulta.Tag = CStr(oFrm.idRegistroSeleccionado)
            Call ObtenerNombreServicio(oFrm.idRegistroSeleccionado, Me.txtCodProductoConsulta, Me.txtProductoConsulta)
        End If
    
End Sub

Sub ObtenerNombreServicio(idServicio As Long, txtCode As TextBox, txtName As TextBox)
    Dim dOServ As New DOCatalogoServicio
    Set dOServ = mo_Facturacion.CatalogoServiciosSeleccionarPorId(idServicio)
    If Not dOServ Is Nothing Then
        txtCode.Text = dOServ.Codigo
        txtName.Text = dOServ.Nombre
    End If
End Sub

Private Sub btnProdInterconsulta_Click()

        Dim oFrm As New SIGHNegocios.BuscaServicio
        oFrm.MostrarFormulario
        If oFrm.idRegistroSeleccionado <> 0 Then
            Me.txtCodProductoInterconsulta.Tag = CStr(oFrm.idRegistroSeleccionado)
            Call ObtenerNombreServicio(oFrm.idRegistroSeleccionado, txtCodProductoInterconsulta, Me.txtProductoInterconsulta)
        End If

End Sub

Private Sub Form_Initialize()
Set mo_cmbIdDepartamento.MiComboBox = Me.cmbIdDepartamento
End Sub

Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDepartamento_LostFocus()
   mo_Formulario.MarcarComoVacio cmbIdDepartamento
End Sub

Private Sub cmbIdDepartamento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombre_LostFocus()
   mo_Formulario.MarcarComoVacio txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Especialidades
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

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Especialidades
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Especialidades"
       Case sghModificar
           Me.Caption = "Modificar Especialidades"
       Case sghConsultar
           Me.Caption = "Consultar Especialidades"
       Case sghEliminar
           Me.Caption = "Eliminar Especialidades"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Especialidades
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
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminServiciosHosp.MensajeError, vbExclamation, Me.Caption
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
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminServiciosHosp.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminServiciosHosp.MensajeError, vbExclamation, Me.Caption
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

   If Me.txtNombre.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de Nombre" + Chr(13)
   End If
   If Val(mo_cmbIdDepartamento.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de IdDepartamento" + Chr(13)
   End If
   

   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   Dim oRsTmp As New Recordset
   If mi_Opcion = sghEliminar Then
        Set oRsTmp = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorEspecialidad(mo_Especialidades.IdEspecialidad)
        If oRsTmp.RecordCount > 0 Then
           MsgBox "La ESPECIALIDAD ya tiene PROGRAMACION MEDICA registrada", vbInformation, "Mensaje"
           Set oRsTmp = Nothing
           Exit Function
        End If
   End If
   If Round(60 / Val(Me.cmbTiempoPromedioAtencion.Text), 0) <> Round(60 / Val(Me.cmbTiempoPromedioAtencion.Text), 2) Then
      MsgBox "El TIEMPO PROMEDIO DE ATENCION deberà ser multiplo de 60", vbInformation, "Mensaje"
      Exit Function
   End If
   
   ValidarReglas = True
   Set oRsTmp = Nothing
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Especialidades
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_Especialidades
           '.IdEspecialidad = Me.IdEspecialidad
           .Nombre = Me.txtNombre.Text
           .IdDepartamento = mo_cmbIdDepartamento.BoundText
           '.TiempoPromedioConsulta = Me.txtTiempoPromedioConsulta.Text
           .IdUsuarioAuditoria = Me.idUsuario
   End With
   With mo_EspecialidadesCE
        .IdEspecialidad = mo_Especialidades.IdEspecialidad
        .IdEspecialidadCE = Me.IdEspecialidadCE
        .IdProductoConsulta = Val(Me.txtCodProductoConsulta.Tag)
        .IdProductoInterconsulta = Val(Me.txtCodProductoInterconsulta.Tag)
        .IdUsuarioAuditoria = Me.idUsuario
        .TiempoPromedioAtencion = Val(Me.cmbTiempoPromedioAtencion.Text)
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminServiciosHosp.EspecialidadesAgregar(mo_Especialidades, mo_EspecialidadesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)
   
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminServiciosHosp.EspecialidadesModificar(mo_Especialidades, mo_EspecialidadesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

    CargaDatosAlObjetosDeDatos
    EliminarDatos = mo_AdminServiciosHosp.EspecialidadesEliminar(mo_Especialidades, mo_EspecialidadesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Especialidades
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

       Set mo_Especialidades = mo_AdminServiciosHosp.EspecialidadesSeleccionarPorId(Me.IdEspecialidad)
       
If mo_AdminServiciosHosp.MensajeError <> "" Then
     MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
     mb_ExistenDatos = False
     Exit Sub
End If
       If Not mo_Especialidades Is Nothing Then
           With mo_Especialidades
           Me.IdEspecialidad = .IdEspecialidad
           Me.txtNombre.Text = .Nombre
           mo_cmbIdDepartamento.BoundText = .IdDepartamento
               mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
   Set mo_EspecialidadesCE = mo_AdminServiciosHosp.EspecialidadesCESeleccionarPorIdEspecialidad(Me.IdEspecialidad)
   
   If mo_AdminServiciosHosp.MensajeError <> "" Then
     MsgBox "No se pudo obtener los datos " + Chr(13) + mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
     mb_ExistenDatos = False
     Exit Sub
   End If
       If Not mo_EspecialidadesCE Is Nothing Then
           With mo_EspecialidadesCE
           Me.IdEspecialidad = .IdEspecialidad
           IdEspecialidadCE = .IdEspecialidadCE
           Me.txtCodProductoConsulta.Tag = .IdProductoConsulta
           Call ObtenerNombreServicio(.IdProductoConsulta, Me.txtCodProductoConsulta, Me.txtProductoConsulta)
           Me.txtCodProductoInterconsulta.Tag = .IdProductoInterconsulta
           Call ObtenerNombreServicio(.IdProductoInterconsulta, Me.txtCodProductoInterconsulta, Me.txtProductoInterconsulta)
           Me.cmbTiempoPromedioAtencion.Text = .TiempoPromedioAtencion
               mb_ExistenDatos = True
           End With
           
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If

End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Especialidades
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           IdEspecialidad = 0
           IdEspecialidadCE = 0
           Me.txtNombre.Text = ""
           mo_cmbIdDepartamento.BoundText = ""
           'Me.txtTiempoPromedioConsulta.Text = ""
           
           Me.txtCodProductoConsulta.Text = ""
           Me.txtCodProductoConsulta.Tag = ""
           Me.txtProductoConsulta.Text = ""
           
           Me.txtCodProductoInterconsulta.Text = ""
           Me.txtCodProductoInterconsulta.Tag = ""
           Me.txtProductoInterconsulta.Text = ""
           
           Me.cmbTiempoPromedioAtencion.Text = ""
   
End Sub


