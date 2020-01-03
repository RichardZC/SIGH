VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FuenteFinanciamientoDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "FuenteFinanciamientoDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatos 
      Height          =   6600
      Left            =   30
      TabIndex        =   13
      Top             =   30
      Width           =   7395
      Begin VB.TextBox txtNcuenteUnidosis 
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
         Left            =   6300
         MaxLength       =   11
         TabIndex        =   31
         Top             =   4200
         Width           =   990
      End
      Begin VB.CheckBox chkUsadoEnUnidosis 
         Alignment       =   1  'Right Justify
         Caption         =   "Es usado en FARMACIA UNIDOSIS"
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
         TabIndex        =   29
         Top             =   4260
         Width           =   3540
      End
      Begin VB.CheckBox chkEsEPS 
         Alignment       =   1  'Right Justify
         Caption         =   "Es EPS (aseguradora no cubre 100%)"
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
         TabIndex        =   28
         Top             =   3960
         Width           =   3540
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   3510
         MaxLength       =   11
         TabIndex        =   2
         Top             =   960
         Width           =   1845
      End
      Begin VB.ComboBox cmbTipoFinanciador 
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
         ItemData        =   "FuenteFinanciamientoDetalle.frx":0CCA
         Left            =   3510
         List            =   "FuenteFinanciamientoDetalle.frx":0CD7
         TabIndex        =   1
         Top             =   600
         Width           =   3840
      End
      Begin VB.TextBox txtCodigoHIS 
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
         Left            =   3510
         MaxLength       =   2
         TabIndex        =   9
         Top             =   3600
         Width           =   525
      End
      Begin VB.CheckBox chkUsadoEnCAJA 
         Alignment       =   1  'Right Justify
         Caption         =   "Usado en CAJA (pacientes Externos)"
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
         TabIndex        =   8
         Top             =   3270
         Width           =   3555
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tarifario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   105
         TabIndex        =   20
         Top             =   4620
         Width           =   7215
         Begin VB.ComboBox cmbIdTipoFinanciamiento 
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
            Left            =   150
            TabIndex        =   24
            Top             =   240
            Width           =   4260
         End
         Begin VB.CommandButton btnQuitarDx 
            DisabledPicture =   "FuenteFinanciamientoDetalle.frx":0D15
            DownPicture     =   "FuenteFinanciamientoDetalle.frx":10A0
            Height          =   315
            Left            =   5400
            Picture         =   "FuenteFinanciamientoDetalle.frx":1433
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   255
            Width           =   825
         End
         Begin VB.CommandButton btnAgregarDx 
            DisabledPicture =   "FuenteFinanciamientoDetalle.frx":17C4
            DownPicture     =   "FuenteFinanciamientoDetalle.frx":1BAD
            Height          =   315
            Left            =   4530
            Picture         =   "FuenteFinanciamientoDetalle.frx":1FB9
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   255
            Width           =   825
         End
         Begin UltraGrid.SSUltraGrid grdFuentesFinanciamientos 
            Height          =   1170
            Left            =   150
            TabIndex        =   23
            Top             =   630
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2064
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
            Caption         =   "Producto/Plan Asignados"
         End
      End
      Begin VB.ComboBox cmbAreaTramitaR 
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
         Left            =   3510
         TabIndex        =   6
         Top             =   2490
         Width           =   3840
      End
      Begin VB.TextBox txtCodigoSEM 
         Alignment       =   1  'Right Justify
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
         Left            =   3510
         MaxLength       =   2
         TabIndex        =   7
         Top             =   2880
         Width           =   525
      End
      Begin VB.ComboBox cmbUtilizadoEn 
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
         ItemData        =   "FuenteFinanciamientoDetalle.frx":23C5
         Left            =   3510
         List            =   "FuenteFinanciamientoDetalle.frx":23D2
         TabIndex        =   5
         Top             =   2100
         Width           =   3840
      End
      Begin VB.ComboBox cmbTipoConceptoF 
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
         Left            =   3510
         TabIndex        =   4
         Top             =   1710
         Width           =   3840
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   3510
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   3825
      End
      Begin VB.TextBox txtIdFuenteFinanciamiento 
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
         Left            =   3510
         TabIndex        =   0
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta para UNIDOSIS"
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
         Left            =   4155
         TabIndex        =   30
         Top             =   4260
         Width           =   2145
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   975
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo Financiador"
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
         TabIndex        =   26
         Top             =   645
         Width           =   1965
      End
      Begin VB.Label Label7 
         Caption         =   "Código Sistema HIS"
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
         TabIndex        =   25
         Top             =   3615
         Width           =   1965
      End
      Begin VB.Label Label6 
         Caption         =   "Código Sistema SEM"
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
         TabIndex        =   19
         Top             =   2895
         Width           =   1965
      End
      Begin VB.Label Label5 
         Caption         =   "Area Tramita Seguro (Reembolsos)"
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
         Top             =   2520
         Width           =   2985
      End
      Begin VB.Label Label4 
         Caption         =   "Utilizado En"
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
         TabIndex        =   17
         Top             =   2145
         Width           =   1965
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Concepto Farmacia (Form.ICI)"
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
         TabIndex        =   16
         Top             =   1770
         Width           =   2985
      End
      Begin VB.Label Label2 
         Caption         =   "Fuente Financiamiento/IAFA"
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
         TabIndex        =   15
         Top             =   1395
         Width           =   2505
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
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
         Top             =   300
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   15
      TabIndex        =   12
      Top             =   6705
      Width           =   7395
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FuenteFinanciamientoDetalle.frx":2410
         DownPicture     =   "FuenteFinanciamientoDetalle.frx":28D4
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
         Left            =   3825
         Picture         =   "FuenteFinanciamientoDetalle.frx":2DC0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FuenteFinanciamientoDetalle.frx":32AC
         DownPicture     =   "FuenteFinanciamientoDetalle.frx":370C
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
         Left            =   2280
         Picture         =   "FuenteFinanciamientoDetalle.frx":3B81
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FuenteFinanciamientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Fuente Financiamiento del Paciente
'        Programado por: Barrantes D
'        Fecha: Agosto 2010
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_FuentesFinanciamiento As New DOFuenteFinanciamiento
Dim mrs_TiposFinanciamientos As New Recordset
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdFuenteFinanciamiento As Long
Dim mo_cmbIdTipoFinanciamiento As New sighentidades.ListaDespleglable

Dim mo_cmbTipoConceptoF As New sighentidades.ListaDespleglable
Dim mo_cmbAreaTramitaR As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcSql As String

Dim mo_cmbIdTipoFinanciador As New sighentidades.ListaDespleglable

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
Dim oRsTmp As New Recordset

       mo_cmbIdTipoFinanciamiento.BoundColumn = "IdTipoFinanciamiento"
       mo_cmbIdTipoFinanciamiento.ListField = "Descripcion"
       Set mo_cmbIdTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloPlanes("")
       '
       Set oRsTmp = mo_ReglasFarmacia.FarmTipoConceptosDevuelveTodos
       mo_cmbTipoConceptoF.BoundColumn = "idTipoConcepto"
       mo_cmbTipoConceptoF.ListField = "Concepto"
       Set mo_cmbTipoConceptoF.RowSource = oRsTmp
       '
       mo_cmbAreaTramitaR.ListField = "Descripcion"
       mo_cmbAreaTramitaR.BoundColumn = "idAreaTramitaSeguros"
       Set mo_cmbAreaTramitaR.RowSource = mo_ReglasFacturacion.AreaTramitaSegurosDevuelveTodosSegunFiltro("")
       '
       mo_cmbIdTipoFinanciador.ListField = "nombre"
       mo_cmbIdTipoFinanciador.BoundColumn = "idTipoFinanciador"
       Set mo_cmbIdTipoFinanciador.RowSource = mo_ReglasFacturacion.TipoFinanciadorSeleccionarTodos()
       Set oRsTmp = Nothing
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
Property Let IdFuenteFinanciamiento(lValue As Long)
   ml_IdFuenteFinanciamiento = lValue
End Property
Property Get IdFuenteFinanciamiento() As Long
   IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
End Property


Private Sub btnAgregarDx_Click()
     If Not Val(mo_cmbIdTipoFinanciamiento.BoundText) > 0 Then
        MsgBox "Elija el Tarifario", vbInformation, Me.Caption
        Exit Sub
     End If
     If mrs_TiposFinanciamientos.RecordCount > 0 Then
        mrs_TiposFinanciamientos.MoveFirst
        mrs_TiposFinanciamientos.Find "idTipoFinanciamiento=" & mo_cmbIdTipoFinanciamiento.BoundText
        If Not mrs_TiposFinanciamientos.EOF Then
           MsgBox "Ese Tarifario ya está registrado", vbInformation, Me.Caption
           Exit Sub
        End If
     End If
     mrs_TiposFinanciamientos.AddNew
     mrs_TiposFinanciamientos.Fields!idTipoFinanciamiento = mo_cmbIdTipoFinanciamiento.BoundText
     mrs_TiposFinanciamientos.Fields!TipoFinanciamiento = cmbIdTipoFinanciamiento.Text
     mrs_TiposFinanciamientos.Update
End Sub

Private Sub btnQuitarDx_Click()
    On Error GoTo errQuit
    mrs_TiposFinanciamientos.Delete
    mrs_TiposFinanciamientos.Update
errQuit:
End Sub

Private Sub cmbIdTipoFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoFinanciamiento
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoFinanciamiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub cmbTipoConceptoF_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
End Sub



Private Sub cmbTipoFinanciador_Click()
    Call asignarValoresCodigo(mo_cmbIdTipoFinanciador.BoundText)
End Sub

Private Sub cmbTipoFinanciador_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbUtilizadoEn_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoFinanciamiento.MiComboBox = cmbIdTipoFinanciamiento
    Set mo_cmbTipoConceptoF.MiComboBox = cmbTipoConceptoF
    Set mo_cmbAreaTramitaR.MiComboBox = cmbAreaTramitaR
    Set mo_cmbIdTipoFinanciador.MiComboBox = cmbTipoFinanciador
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mrs_TiposFinanciamientos.Close
    Set mrs_TiposFinanciamientos = Nothing
End Sub





Private Sub grdFuentesFinanciamientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdFuentesFinanciamientos.Bands(0).Columns("IdTipoFinanciamiento").Header.Caption = "Id"
    grdFuentesFinanciamientos.Bands(0).Columns("TipoFinanciamiento").Header.Caption = "Producto/Plan"
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> "" Then
        If validarValorCodigo(txtCodigo.Text, mo_cmbIdTipoFinanciador.BoundText) = False Then
            MsgBox "Ingrese RUC valido", vbInformation, Me.Caption
            txtCodigo.SetFocus
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

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla FuentesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
 Select Case mi_Opcion
     Case sghAgregar
          Me.txtCodigoHIS.Text = "8"
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
'   Descripción:    Seleccionar un registro unico de la tabla FuentesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       GenerarRecordsetTemporal
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Fuente Financiamiento (IAFA)"
       Case sghModificar
           Me.Caption = "Modificar Fuente Financiamiento (IAFA)"
       Case sghConsultar
           Me.Caption = "Consultar Fuente Financiamiento (IAFA)"
           Me.fraDatos.Enabled = False
       Case sghEliminar
           Me.Caption = "Eliminar Fuente Financiamiento (IAFA)"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       asignarAyudas
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla FuentesFinanciamiento
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
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   'LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
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
   
   If Me.cmbTipoFinanciador.Text = "" Then
       sMensaje = sMensaje + "Seleccione tipo de fuente de financiamiento" + Chr(13)
   End If
   If Me.txtCodigo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de fuente de financiamiento" + Chr(13)
   End If
   If Val(Me.txtIdFuenteFinanciamiento.Text) = 0 Then
       sMensaje = sMensaje + "Ingrese el ID de la fuente de financiamiento" + Chr(13)
   End If
   If Me.txtDescripcion.Text = "" Then
       sMensaje = sMensaje + "Ingrese la descripción" + Chr(13)
   End If
   If mrs_TiposFinanciamientos.RecordCount = 0 Then
       sMensaje = sMensaje + "Debe haber al menos un Tarifario asociado a la FUENTE DE FINANCIAMIENTO" + Chr(13)
   End If
    
   
   If Me.cmbTipoConceptoF.Text = "" Then
       sMensaje = sMensaje + "Ingrese Tipo Concepto para Farmacia" + Chr(13)
   End If
   If Me.cmbUtilizadoEn.Text = "" Then
       sMensaje = sMensaje + "Ingrese donde  será UTILIZADO la FUENTE DE FINANCIAMIENTO" + Chr(13)
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   If existeCodigoDuplicado(txtCodigo.Text, mo_cmbIdTipoFinanciador.BoundText, _
                ml_IdFuenteFinanciamiento) = True Then
        MsgBox "Código de Fuente de Financiamiento ya ha sido ingresado", vbInformation, Me.Caption
        Exit Function
   End If
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla FuentesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_FuentesFinanciamiento
           .IdFuenteFinanciamiento = Me.txtIdFuenteFinanciamiento
           .idTipoFinanciamiento = Val(mo_cmbIdTipoFinanciamiento.BoundText)
           .Descripcion = Me.txtDescripcion.Text
           .idTipoConceptoFarmacia = Val(mo_cmbTipoConceptoF.BoundText)
           .UtilizadoEn = Me.cmbUtilizadoEn.ListIndex + 1
           .IdUsuarioAuditoria = ml_idUsuario
           .CodigoFuenteFinanciamientoSEM = txtCodigoSEM.Text
           .idAreaTramitaSeguros = Val(mo_cmbAreaTramitaR.BoundText)
           .EsUsadoEnCaja = IIf(Me.chkUsadoEnCAJA.Value = 1, True, False)
           .CodigoHIS = Me.txtCodigoHIS.Text
           .codigo = Me.txtCodigo.Text
           .idTipoFinanciador = Val(mo_cmbIdTipoFinanciador.BoundText)
           .TieneEPS = IIf(Me.chkEsEPS.Value = 1, 1, 0)
           .usadoEnFUnidosis = IIf(Me.chkUsadoEnUnidosis.Value = 1, 1, 0)
           .CuentaParaUnidosis = Val(txtNcuenteUnidosis.Text)
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminFacturacion.FuentesFinanciamientoAgregar(mo_FuentesFinanciamiento, mrs_TiposFinanciamientos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminFacturacion.FuentesFinanciamientoModificar(mo_FuentesFinanciamiento, mrs_TiposFinanciamientos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
   Dim oRsTmp As New Recordset
   Set oRsTmp = mo_ReglasAdmision.AtencionesSeleccionarPorIdFuenteFinanciamiento(mo_FuentesFinanciamiento.IdFuenteFinanciamiento)
   If oRsTmp.RecordCount > 0 Then
      MsgBox "No se podrá eliminar porque existen Atenciones registradas", vbInformation, Me.Caption
      Exit Function
   End If
   oRsTmp.Close
   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminFacturacion.FuentesFinanciamientoEliminar(mo_FuentesFinanciamiento, mrs_TiposFinanciamientos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla FuentesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
        Dim oRsTmp As New Recordset
        Set mo_FuentesFinanciamiento = mo_AdminFacturacion.FuentesFinanciamientoSeleccionarPorId(Me.IdFuenteFinanciamiento)
        If mo_AdminFacturacion.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
       
       If Not mo_FuentesFinanciamiento Is Nothing Then
           With mo_FuentesFinanciamiento
                Me.IdFuenteFinanciamiento = .IdFuenteFinanciamiento
                Me.txtIdFuenteFinanciamiento = .IdFuenteFinanciamiento
                mo_cmbIdTipoFinanciamiento.BoundText = .idTipoFinanciamiento
                Me.txtDescripcion.Text = .Descripcion
                mo_cmbTipoConceptoF.BoundText = .idTipoConceptoFarmacia
                Me.cmbUtilizadoEn.ListIndex = .UtilizadoEn - 1
                txtCodigoSEM.Text = IIf(IsNull(.CodigoFuenteFinanciamientoSEM), "", .CodigoFuenteFinanciamientoSEM)
                mo_cmbAreaTramitaR.BoundText = .idAreaTramitaSeguros
                Me.chkUsadoEnCAJA.Value = IIf(.EsUsadoEnCaja = True, 1, 0)
                Me.txtCodigoHIS.Text = .CodigoHIS
                Me.txtCodigo.Text = .codigo
                mo_cmbIdTipoFinanciador.BoundText = .idTipoFinanciador
                Me.chkEsEPS.Value = .TieneEPS
                Me.chkUsadoEnUnidosis.Value = .usadoEnFUnidosis
                txtNcuenteUnidosis = .CuentaParaUnidosis
                mb_ExistenDatos = True
           End With
           'Carga grid Tarifario
           Set oRsTmp = mo_AdminFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(mo_FuentesFinanciamiento.IdFuenteFinanciamiento)
           If oRsTmp.RecordCount > 0 Then
              Do While Not oRsTmp.EOF
                    mrs_TiposFinanciamientos.AddNew
                    mrs_TiposFinanciamientos.Fields!idTipoFinanciamiento = oRsTmp.Fields!idTipoFinanciamiento
                    mrs_TiposFinanciamientos.Fields!TipoFinanciamiento = oRsTmp.Fields!Descripcion
                    mrs_TiposFinanciamientos.Update
                    oRsTmp.MoveNext
              Loop
           End If
           oRsTmp.Close
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       Set oRsTmp = Nothing
       If Me.IdFuenteFinanciamiento <= 3 Or Me.IdFuenteFinanciamiento = 5 Then
          Me.btnAceptar.Enabled = False
       End If
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla FuentesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
    Me.IdFuenteFinanciamiento = 0
    Me.txtIdFuenteFinanciamiento = ""
    mo_cmbIdTipoFinanciamiento.BoundText = ""
    Me.txtDescripcion.Text = ""
    mo_cmbIdTipoFinanciador.BoundText = ""
    Me.txtCodigo = ""
    Me.chkEsEPS.Value = 0
    Me.chkUsadoEnUnidosis.Value = 0
    txtNcuenteUnidosis.Text = ""
End Sub

Private Sub txtIdFuenteFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdFuenteFinanciamiento
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdFuenteFinanciamiento_LostFocus()
   mo_Formulario.MarcarComoVacio txtIdFuenteFinanciamiento
End Sub

Private Sub txtIdFuenteFinanciamiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub GenerarRecordsetTemporal()
    With mrs_TiposFinanciamientos
          .Fields.Append "IdTipoFinanciamiento", adInteger, 4, adFldIsNullable
          .Fields.Append "TipoFinanciamiento", adVarChar, 100, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdFuentesFinanciamientos.DataSource = mrs_TiposFinanciamientos
    mo_Apariencia.ConfigurarFilasBiColores Me.grdFuentesFinanciamientos, sighentidades.GrillaConFilasBicolor
End Sub

Private Function HabilitarDesabilitarCodigo(ml_idTipoFinanciador As Long) As Boolean
    Select Case ml_idTipoFinanciador
        Case sghidestipofinanciador.PersonaJuridica
            mo_Formulario.HabilitarDeshabilitar txtCodigo, True
        Case Else
            mo_Formulario.HabilitarDeshabilitar txtCodigo, False
    End Select
    HabilitarDesabilitarCodigo = True
End Function

Private Function LeerCodigoPredefinidoTipoFinanciador(ml_idTipoFinanciador As Long) As String
    LeerCodigoPredefinidoTipoFinanciador = ""
    
    Dim oDoTipoFinanciador As DoTipoFinanciador
    
    Set oDoTipoFinanciador = mo_ReglasFacturacion.TipoFinanciadorSeleccionarPorId(ml_idTipoFinanciador)
    If Not (oDoTipoFinanciador Is Nothing) Then
        LeerCodigoPredefinidoTipoFinanciador = oDoTipoFinanciador.codigo
    Else
        MsgBox mo_ReglasFacturacion.MensajeError, vbInformation, Me.Caption
    End If
End Function

Private Function asignarValoresCodigo(ml_idTipoFinanciador As Long)
    Dim sCodigo As String
    Call HabilitarDesabilitarCodigo(ml_idTipoFinanciador)
    
    sCodigo = LeerCodigoPredefinidoTipoFinanciador(ml_idTipoFinanciador)
    If sCodigo <> "" Then
        txtCodigo.Text = sCodigo
    End If
End Function

Private Function asignarAyudas()
    txtCodigo.ToolTipText = "Para el Caso de Personas Jurídicas consigar como código el RUC del Financiador"
End Function

Private Function validarValorCodigo(sCodigo As String, lnIdTipoFinanciador As Long) As Boolean
    Dim bReturnValue As Boolean
    
    Select Case lnIdTipoFinanciador
        Case sghidestipofinanciador.PersonaJuridica:
           bReturnValue = sighentidades.EsRucCorrecto(sCodigo)
        Case Else
            bReturnValue = True
    End Select
    validarValorCodigo = bReturnValue
End Function

Private Function existeCodigoDuplicado(sCodigo As String, _
            lnIdTipoFinanciador As Long, lIdFuenteFinanciamiento As Long) As Boolean
            
    Dim bReturnValue As Boolean
    Dim oRs As ADODB.Recordset
    
    bReturnValue = False
    If lnIdTipoFinanciador = sghidestipofinanciador.PersonaJuridica Then
        Set oRs = mo_ReglasFacturacion.FuentesFinanciamientoSeleccionarPorCodigo(sCodigo)
        If Not (oRs Is Nothing) Then
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst
                oRs.Find "codigo='" & sCodigo & "'"
                If oRs.EOF = False Then
                    If lIdFuenteFinanciamiento <> oRs.Fields!IdFuenteFinanciamiento Then
                        bReturnValue = True
                    End If
                End If
            End If
        End If
    End If
    existeCodigoDuplicado = bReturnValue
End Function
