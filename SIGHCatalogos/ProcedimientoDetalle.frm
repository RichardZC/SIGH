VERSION 5.00
Begin VB.Form ProcedimientoDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "ProcedimientoDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   2100
      TabIndex        =   22
      Top             =   1650
      Width           =   345
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   19
      Top             =   2820
      Width           =   9735
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ProcedimientoDetalle.frx":08CA
         DownPicture     =   "ProcedimientoDetalle.frx":0D8E
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
         Left            =   4950
         Picture         =   "ProcedimientoDetalle.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ProcedimientoDetalle.frx":1766
         DownPicture     =   "ProcedimientoDetalle.frx":1BC6
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
         Left            =   3405
         Picture         =   "ProcedimientoDetalle.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame frmRestricciones 
      Caption         =   "Restricciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   45
      TabIndex        =   15
      Top             =   2145
      Width           =   9750
      Begin VB.ComboBox cmbIdTipoSexo 
         Height          =   315
         Left            =   960
         TabIndex        =   24
         Top             =   225
         Width           =   2145
      End
      Begin VB.TextBox txtEdadMinDias 
         Height          =   315
         Left            =   5625
         MaxLength       =   5
         TabIndex        =   8
         Top             =   195
         Width           =   870
      End
      Begin VB.TextBox txtEdadMaxDias 
         Height          =   315
         Left            =   8790
         MaxLength       =   5
         TabIndex        =   7
         Top             =   225
         Width           =   870
      End
      Begin VB.Label lblIdTipoSexo 
         Caption         =   "Sexo"
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
         TabIndex        =   18
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblEdadMinDias 
         Alignment       =   1  'Right Justify
         Caption         =   "Edad mínima "
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
         Left            =   4410
         TabIndex        =   17
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lblEdadMaxDias 
         Caption         =   "Edad máxima"
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
         Left            =   7605
         TabIndex        =   16
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame frmDatos 
      Height          =   2085
      Left            =   45
      TabIndex        =   9
      Top             =   30
      Width           =   9765
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   2445
         TabIndex        =   23
         Top             =   1650
         Width           =   5190
      End
      Begin VB.TextBox txtIdProducto 
         Height          =   315
         Left            =   990
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1620
         Width           =   1000
      End
      Begin VB.TextBox txtDescripcionOPCS 
         Height          =   315
         Left            =   2055
         MaxLength       =   10
         TabIndex        =   3
         Top             =   900
         Width           =   7575
      End
      Begin VB.TextBox txtCodigoOPCS 
         Height          =   315
         Left            =   990
         MaxLength       =   10
         TabIndex        =   2
         Top             =   900
         Width           =   1000
      End
      Begin VB.CheckBox chkRestriccion 
         Alignment       =   1  'Right Justify
         Caption         =   "Tiene restricción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7890
         TabIndex        =   6
         Top             =   1665
         Width           =   1725
      End
      Begin VB.TextBox txtCodigoCPT2004 
         Height          =   315
         Left            =   990
         MaxLength       =   7
         TabIndex        =   0
         Top             =   555
         Width           =   1000
      End
      Begin VB.TextBox txtCodigoCPT99 
         Height          =   315
         Left            =   990
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1260
         Width           =   1000
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   2055
         MaxLength       =   250
         TabIndex        =   1
         Top             =   555
         Width           =   7575
      End
      Begin VB.Label lblIdProducto 
         Caption         =   "Producto"
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
         Top             =   1695
         Width           =   1335
      End
      Begin VB.Label lblCodigoOPCS 
         Caption         =   "OPCS"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label lblCodigoCPT2004 
         Caption         =   "CPT2000"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCodigoCPT99 
         Caption         =   "CPT99"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Código            Descripción"
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
         Left            =   1215
         TabIndex        =   10
         Top             =   270
         Width           =   2655
      End
   End
End
Attribute VB_Name = "ProcedimientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Procedimientos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Procedimientos As New DOProcedimiento
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_AdminServiciosComunes As New ReglasComunes
Dim mo_CmbIdTipoSexo As New ListaDespleglable
Dim ml_IdProcedimiento As Long
Property Let IdProcedimiento(lValue As Long)
   ml_IdProcedimiento = lValue
End Property
Property Get IdProcedimiento() As Long
   IdProcedimiento = ml_IdProcedimiento
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


Private Sub Form_Initialize()
    Set mo_CmbIdTipoSexo.MiComboBox = cmbIdTipoSexo
End Sub

Private Sub txtEdadMinDias_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEdadMinDias
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtEdadMinDias_LostFocus()
   mo_Formulario.MarcarComoVacio txtEdadMinDias
End Sub

Private Sub txtEdadMinDias_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdProducto_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdProducto
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdProducto_LostFocus()
   mo_Formulario.MarcarComoVacio txtIdProducto
End Sub

Private Sub txtIdProducto_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoSexo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSexo
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoSexo_LostFocus()
   If cmbIdTipoSexo.Text <> "" Then
       mo_CmbIdTipoSexo.BoundText = Val(Split(cmbIdTipoSexo.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoSexo
End Sub

Private Sub cmbIdTipoSexo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtDescripcionOPCS_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDescripcionOPCS
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDescripcionOPCS_LostFocus()
   mo_Formulario.MarcarComoVacio txtDescripcionOPCS
End Sub


Private Sub txtCodigoOPCS_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoOPCS
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigoOPCS_LostFocus()
   mo_Formulario.MarcarComoVacio txtCodigoOPCS
End Sub

Private Sub txtCodigoOPCS_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtEdadMaxDias_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEdadMaxDias
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtEdadMaxDias_LostFocus()
   mo_Formulario.MarcarComoVacio txtEdadMaxDias
End Sub

Private Sub txtEdadMaxDias_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub chkRestriccion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, chkRestriccion
AdministrarKeyPreview KeyCode
End Sub

Private Sub chkRestriccion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtCodigoCPT2004_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoCPT2004
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigoCPT2004_LostFocus()
   mo_Formulario.MarcarComoVacio txtCodigoCPT2004
End Sub

Private Sub txtCodigoCPT2004_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtCodigoCPT99_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoCPT99
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigoCPT99_LostFocus()
   mo_Formulario.MarcarComoVacio txtCodigoCPT99
End Sub

Private Sub txtCodigoCPT99_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDescripcion_LostFocus()
    mo_Formulario.MarcarComoVacio txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Procedimientos
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
'   Descripción:    Seleccionar un registro unico de la tabla Procedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Procedimientos"
       Case sghModificar
           Me.Caption = "Modificar Procedimientos"
       Case sghConsultar
           Me.Caption = "Consultar Procedimientos"
           Me.frmDatos.Enabled = False
           Me.frmRestricciones.Enabled = False
       Case sghEliminar
           Me.Caption = "Eliminar Procedimientos"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Procedimientos
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
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    LimpiarFormulario
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
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
   
   If Me.chkRestriccion.Value = 1 Then
        If Me.txtEdadMaxDias.Text = 0 Then
            sMensaje = sMensaje + "Ingrese el valor de EdadMaxDias" + Chr(13)
        End If
        If Me.txtEdadMinDias.Text = "" Then
            sMensaje = sMensaje + "Ingrese el valor de EdadMinDias" + Chr(13)
        End If
   End If
   If Me.txtCodigoCPT2004.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de CodigoCPT2004" + Chr(13)
   End If
   If Me.txtDescripcion.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de Descripcion" + Chr(13)
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
'   Descripción:    Seleccionar un registro unico de la tabla Procedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_Procedimientos
           .EdadMinDias = Val(Me.txtEdadMinDias.Text)
           .idProducto = Val(Me.txtIdProducto.Text)
           .idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
           .DescripcionOPCS = Me.txtDescripcionOPCS.Text
           .CodigoOPCS = Me.txtCodigoOPCS.Text
           .EdadMaxDias = Val(Me.txtEdadMaxDias.Text)
           .Restriccion = Me.chkRestriccion.Value
           .CodigoCPT2004 = Me.txtCodigoCPT2004.Text
           .CodigoCPT99 = Me.txtCodigoCPT99.Text
           .Descripcion = Me.txtDescripcion.Text
           .IdProcedimiento = Val(Me.IdProcedimiento)
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    
    CargaDatosAlObjetosDeDatos
    AgregarDatos = mo_AdminServiciosComunes.ProcedimientosAgregar(mo_Procedimientos)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_AdminServiciosComunes.ProcedimientosModificar(mo_Procedimientos)
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    EliminarDatos = mo_AdminServiciosComunes.ProcedimientosEliminar(mo_Procedimientos)
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Procedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

       Set mo_Procedimientos = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorId(Me.IdProcedimiento)
        If mo_AdminServiciosComunes.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption
             mb_ExistenDatos = False
             Exit Sub
        End If
        
       If Not mo_Procedimientos Is Nothing Then
           With mo_Procedimientos
                Me.txtEdadMinDias.Text = .EdadMinDias
                Me.txtIdProducto.Text = .idProducto
                mo_CmbIdTipoSexo.BoundText = .idTipoSexo
                Me.txtDescripcionOPCS.Text = .DescripcionOPCS
                Me.txtCodigoOPCS.Text = .CodigoOPCS
                Me.txtEdadMaxDias.Text = .EdadMaxDias
                Me.chkRestriccion.Value = .Restriccion
                Me.txtCodigoCPT2004.Text = .CodigoCPT2004
                Me.txtCodigoCPT99.Text = .CodigoCPT99
                Me.txtDescripcion.Text = .Descripcion
                mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Procedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.txtEdadMinDias.Text = ""
           Me.txtIdProducto.Text = ""
           mo_CmbIdTipoSexo.BoundText = ""
           Me.txtDescripcionOPCS.Text = ""
           Me.txtCodigoOPCS.Text = ""
           Me.txtEdadMaxDias.Text = ""
           Me.chkRestriccion.Value = 0
           Me.txtCodigoCPT2004.Text = ""
           Me.txtCodigoCPT99.Text = ""
           Me.txtDescripcion.Text = ""
   
End Sub


Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

    mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
    mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
    Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
       If sMensaje <> "" Then
           MsgBox mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption
       End If

End Sub
