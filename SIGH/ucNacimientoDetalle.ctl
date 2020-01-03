VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucNacimientoDetalle 
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10965
   LockControls    =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   10965
   Begin VB.Frame fraNacimientos 
      Caption         =   "Nacimientos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   30
      TabIndex        =   16
      Top             =   15
      Width           =   10905
      Begin VB.ComboBox cmbIdDocIdentidad 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4935
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1095
         Width           =   2655
      End
      Begin VB.TextBox txtNroDocumento 
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
         Left            =   7605
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1095
         Width           =   1395
      End
      Begin VB.TextBox txtNroHijo 
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
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1110
         Width           =   585
      End
      Begin VB.TextBox TxtNroOrdenParto 
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
         Left            =   90
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1110
         Width           =   1545
      End
      Begin VB.TextBox txtApgar5 
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
         Left            =   8340
         TabIndex        =   7
         Top             =   450
         Width           =   645
      End
      Begin VB.TextBox txtApgar1 
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
         Left            =   7680
         TabIndex        =   6
         Top             =   450
         Width           =   615
      End
      Begin VB.ComboBox cmbIdTipoSexo 
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
         ItemData        =   "ucNacimientoDetalle.ctx":0000
         Left            =   1560
         List            =   "ucNacimientoDetalle.ctx":0002
         TabIndex        =   1
         Top             =   450
         Width           =   1230
      End
      Begin VB.ComboBox cmbIdCondicionRN 
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
         ItemData        =   "ucNacimientoDetalle.ctx":0004
         Left            =   105
         List            =   "ucNacimientoDetalle.ctx":0006
         TabIndex        =   0
         Top             =   450
         Width           =   1395
      End
      Begin VB.TextBox txtPeso 
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
         Left            =   2820
         MaxLength       =   4
         TabIndex        =   2
         Top             =   450
         Width           =   825
      End
      Begin VB.TextBox txtTalla 
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
         Left            =   3690
         MaxLength       =   3
         TabIndex        =   3
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox txtEdad 
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   4
         Top             =   450
         Width           =   1035
      End
      Begin MSMask.MaskEdBox txtFechaNacimiento 
         Height          =   315
         Left            =   5610
         TabIndex        =   5
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand btnAgregar 
         Height          =   435
         Left            =   9345
         TabIndex        =   13
         Top             =   135
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "ucNacimientoDetalle.ctx":0008
         Caption         =   "Agregar"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand btnQuitar 
         Height          =   435
         Left            =   9345
         TabIndex        =   14
         Top             =   990
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "ucNacimientoDetalle.ctx":2F94
         Caption         =   "Quitar"
         PictureAlignment=   9
         ShapeSize       =   1
      End
      Begin MSMask.MaskEdBox txtFclamplaje 
         Height          =   315
         Left            =   2670
         TabIndex        =   10
         Top             =   1110
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand cmdModificar 
         Height          =   435
         Left            =   9345
         TabIndex        =   26
         Top             =   555
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "ucNacimientoDetalle.ctx":5416
         Caption         =   "Modificar"
         PictureAlignment=   9
      End
      Begin VB.Label Label3 
         Caption         =   "N° Orden en Parto    N° Hijo    Fecha Clampaje                   Tipo Documento Identidad      Nª Documento"
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
         Left            =   90
         TabIndex        =   25
         Top             =   900
         Width           =   9135
      End
      Begin VB.Label Label2 
         Caption         =   "Apgar5"
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
         Left            =   8340
         TabIndex        =   24
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Apgar1"
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
         Left            =   7650
         TabIndex        =   23
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label30 
         Caption         =   "Condición"
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
         TabIndex        =   22
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label31 
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
         Left            =   1560
         TabIndex        =   21
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label32 
         Caption         =   "Peso (gr.)"
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
         Left            =   2820
         TabIndex        =   20
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label33 
         Caption         =   "Talla (cm.)"
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
         Left            =   3660
         TabIndex        =   19
         Top             =   210
         Width           =   885
      End
      Begin VB.Label Label38 
         Caption         =   "Edad (Sem)"
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
         Left            =   4590
         TabIndex        =   18
         Top             =   210
         Width           =   1020
      End
      Begin VB.Label Label39 
         Caption         =   "F.nacimiento"
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
         Left            =   5640
         TabIndex        =   17
         Top             =   210
         Width           =   1095
      End
   End
   Begin UltraGrid.SSUltraGrid grdNacimientos 
      Height          =   1425
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Pulser ENTER para MODIFICAR o QUITAR"
      Top             =   1530
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   2514
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      ScrollBars      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de nacimientos"
   End
End
Attribute VB_Name = "ucNacimientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registrar datos del RN
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idUsuario As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ms_MensajeError As String
Dim mrs_Nacimientos As New ADODB.Recordset
Dim ml_TipoDiagnostico As sghTiposDiagnostico
Dim mda_FechaIngreso As Date
Dim mo_cmbIdCondicionRN As New ListaDespleglable
Dim mo_CmbIdTipoSexo As New ListaDespleglable
Dim mo_cmbIdDocIdentidad As New SIGHEntidades.ListaDespleglable
Dim ml_idTipoSexo As Long
Dim mda_FechaNacimiento As Date
Dim lbModificar As Boolean
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Property Get FechaNacimiento() As Long
   FechaNacimiento = mda_FechaNacimiento
End Property
Property Get idTipoSexo() As Long
   idTipoSexo = ml_idTipoSexo
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let FechaIngreso(daValue As Date)
   mda_FechaIngreso = daValue
End Property

Private Sub cmbIdCondicionRN_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCondicionRN
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdCondicionRN_LostFocus()
   If cmbIdCondicionRN.Text <> "" Then
       mo_cmbIdCondicionRN.BoundText = Val(Split(cmbIdCondicionRN.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdCondicionRN
End Sub

Private Sub cmbIdCondicionRN_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub cmbIdDocIdentidad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDocIdentidad
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbIdTipoSexo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSexo
RaiseEvent SePresionoTeclaEspecial(KeyCode)
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

Public Function Inicializar()
    GenerarRecordsetTemporal
    Set mo_cmbIdCondicionRN.MiComboBox = cmbIdCondicionRN
    Set mo_CmbIdTipoSexo.MiComboBox = cmbIdTipoSexo
    Set mo_cmbIdDocIdentidad.MiComboBox = cmbIdDocIdentidad
    HabilidaBotones True
End Function












Private Sub cmdModificar_Click()
    ActualizaDatos
End Sub

Private Sub grdNacimientos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       With mrs_Nacimientos
            UserControl.txtFechaNacimiento.Text = Format(.Fields!FechaNacimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            UserControl.txtEdad = .Fields!EdadSemanas
            UserControl.txtTalla = .Fields!Talla
            UserControl.txtPeso = .Fields!Peso
            mo_cmbIdCondicionRN.BoundText = .Fields!idCondicionRN
            mo_CmbIdTipoSexo.BoundText = .Fields!idTipoSexo
            UserControl.txtApgar1.Text = .Fields!apgar_1
            UserControl.txtApgar5.Text = .Fields!apgar_5
            UserControl.txtFclamplaje.Text = Format(.Fields!clamplajeFecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            UserControl.TxtNroOrdenParto.Text = .Fields!NroOrdenHijoEnParto
            UserControl.txtNroHijo.Text = .Fields!NroOrdenHijo
            mo_cmbIdDocIdentidad.BoundText = .Fields!IdDocIdentidad
            UserControl.txtNroDocumento.Text = .Fields!DocIdentidad
       End With
       HabilidaBotones False
       lbModificar = True
       cmbIdCondicionRN.SetFocus
    End If
End Sub

Private Sub txtApgar1_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApgar1
    RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub



Private Sub txtApgar1_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub txtApgar5_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApgar5
    RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub txtApgar5_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub







Private Sub txtFclamplaje_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFclamplaje
    RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub txtFclamplaje_LostFocus()
    If Not IsDate(txtFclamplaje.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFclamplaje.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If

End Sub



Private Sub txtNroDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumento
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtNroDocumento_LostFocus()
    If Val(mo_cmbIdDocIdentidad.BoundText) = 1 And Len(txtNroDocumento.Text) <> 8 Then
       MsgBox "La longitud debe ser 8 para el DNI", vbInformation, "Nacimiento "
       txtNroDocumento.Text = ""
    End If
End Sub

Private Sub txtNroHijo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHijo
    RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub txtNroHijo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub TxtNroOrdenParto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, TxtNroOrdenParto
    RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub TxtNroOrdenParto_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub UserControl_Resize()
    
    
    fraNacimientos.Width = UserControl.Width - 20
    UserControl.grdNacimientos.Width = UserControl.Width - 20
    UserControl.grdNacimientos.Height = UserControl.Height - 1600
    
End Sub

Public Sub ConfigurarComboBoxes()
Dim sMensaje As String
       
        mo_CmbIdTipoSexo.BoundColumn = "IdTipoSexo"
        mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
        Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos
        
        mo_cmbIdCondicionRN.BoundColumn = "IdCOndicionRN"
        mo_cmbIdCondicionRN.ListField = "DescripcionLarga"
        Set mo_cmbIdCondicionRN.RowSource = mo_AdminServiciosComunes.TiposCondicionRNSeleccionarTodos
        
        mo_cmbIdDocIdentidad.BoundColumn = "IdDocIdentidad"
        mo_cmbIdDocIdentidad.ListField = "DescripcionLarga"
        Set mo_cmbIdDocIdentidad.RowSource = mo_AdminServiciosComunes.TiposDocIdentidadSeleccionarTodosIncSinTipoDoc()
        
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError

End Sub

Private Sub grdNacimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdNacimientos.Bands(0).Columns("IdCondicionRN").Hidden = True
    grdNacimientos.Bands(0).Columns("IdTipoSexo").Hidden = True
    
    grdNacimientos.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    grdNacimientos.Bands(0).Columns("FechaNacimiento").Width = 1200
    
    grdNacimientos.Bands(0).Columns("EdadSemanas").Header.Caption = "Edad (Sem)"
    grdNacimientos.Bands(0).Columns("EdadSemanas").Width = 1000
    
    grdNacimientos.Bands(0).Columns("Talla").Header.Caption = "Talla (cm)"
    grdNacimientos.Bands(0).Columns("Talla").Width = 1000
    
    grdNacimientos.Bands(0).Columns("Peso").Header.Caption = "Peso (gr)"
    grdNacimientos.Bands(0).Columns("Peso").Width = 1000
    
    grdNacimientos.Bands(0).Columns("CondicionRN").Header.Caption = "Condición"
    grdNacimientos.Bands(0).Columns("CondicionRN").Width = 2000
    
    grdNacimientos.Bands(0).Columns("Sexo").Header.Caption = "Sexo"
    grdNacimientos.Bands(0).Columns("Sexo").Width = 2000
    
    grdNacimientos.Bands(0).Columns("IdDocIdentidad").Header.Caption = "Tipo Documento Identidad"
    grdNacimientos.Bands(0).Columns("docIdentidad").Header.Caption = "Nª Documento"

End Sub
Private Sub txtTalla_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtTalla
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtTalla_LostFocus()
   mo_Formulario.MarcarComoVacio txtTalla
End Sub

Private Sub txtTalla_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtPeso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPeso
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtPeso_LostFocus()
   mo_Formulario.MarcarComoVacio txtPeso
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaNacimiento
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFechaNacimiento_LostFocus()
    If Not IsDate(txtFechaNacimiento.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFechaNacimiento.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    Else
        txtFclamplaje.Text = txtFechaNacimiento.Text
    End If
    mo_Formulario.MarcarComoVacio txtFechaNacimiento
End Sub

Private Sub txtFechaNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtEdad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtEdad
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtEdad_LostFocus()
   mo_Formulario.MarcarComoVacio txtEdad
End Sub

Private Sub txtEdad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CargarDatosDeNacimientos(oConexion As Connection)
Dim rsNacimientos As New Recordset

    Set rsNacimientos = mo_AdminAdmision.AtencionesNacimientosSeleccionarPorAtencion(ml_idAtencion, oConexion)
    Do While Not rsNacimientos.EOF
        With mrs_Nacimientos
            .AddNew
            .Fields!FechaNacimiento = Format(rsNacimientos!FechaNacimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            .Fields!EdadSemanas = rsNacimientos!EdadSemanas
            .Fields!Talla = rsNacimientos!Talla
            .Fields!Peso = rsNacimientos!Peso
            .Fields!idCondicionRN = rsNacimientos!idCondicionRN
            .Fields!CondicionRN = rsNacimientos!CondicionRN
            .Fields!idTipoSexo = rsNacimientos!idTipoSexo
            .Fields!Sexo = rsNacimientos!Sexo
            .Fields!apgar_1 = IIf(IsNull(rsNacimientos!apgar_1), 0, rsNacimientos!apgar_1)
            .Fields!apgar_5 = IIf(IsNull(rsNacimientos!apgar_5), 0, rsNacimientos!apgar_5)
            .Fields!clamplajeFecha = IIf(IsNull(rsNacimientos!clamplajeFecha), 0, Format(rsNacimientos!clamplajeFecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM))
            .Fields!NroOrdenHijoEnParto = rsNacimientos!NroOrdenHijoEnParto
            .Fields!NroOrdenHijo = rsNacimientos!NroOrdenHijo
            .Fields!DocIdentidad = IIf(IsNull(rsNacimientos!DocIdentidad), "", rsNacimientos!DocIdentidad)
            .Fields!IdDocIdentidad = IIf(IsNull(rsNacimientos!IdDocIdentidad), 0, rsNacimientos!IdDocIdentidad)
        End With
        rsNacimientos.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores UserControl.grdNacimientos, SIGHEntidades.GrillaConFilasBicolor
    
End Sub

Sub CargarNacimientosAlObjetoDatos(oNacimientos As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS NacimientoS
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ProcedimientoS
    '---------------------------------------------------------------------------------
    Dim oNacimiento As DOAtencionNacimiento
    
    If mrs_Nacimientos.RecordCount > 0 Then
        mrs_Nacimientos.MoveFirst
        Do While Not mrs_Nacimientos.EOF
            Set oNacimiento = New DOAtencionNacimiento
            oNacimiento.idNacimiento = 0
            oNacimiento.EdadSemanas = mrs_Nacimientos!EdadSemanas
            oNacimiento.FechaNacimiento = mrs_Nacimientos!FechaNacimiento
            oNacimiento.idCondicionRN = mrs_Nacimientos!idCondicionRN
            oNacimiento.idTipoSexo = mrs_Nacimientos!idTipoSexo
            oNacimiento.Peso = mrs_Nacimientos!Peso
            oNacimiento.Talla = mrs_Nacimientos!Talla
            oNacimiento.IdUsuarioAuditoria = ml_idUsuario
            oNacimiento.apgar_1 = mrs_Nacimientos!apgar_1
            oNacimiento.apgar_5 = mrs_Nacimientos!apgar_5
            oNacimiento.clamplajeFecha = mrs_Nacimientos!clamplajeFecha
            oNacimiento.NroOrdenHijoEnParto = mrs_Nacimientos!NroOrdenHijoEnParto
            oNacimiento.NroOrdenHijo = mrs_Nacimientos!NroOrdenHijo
            oNacimiento.IdDocIdentidad = mrs_Nacimientos!IdDocIdentidad
            oNacimiento.DocIdentidad = mrs_Nacimientos!DocIdentidad
            oNacimientos.Add oNacimiento
            mrs_Nacimientos.MoveNext
        Loop
    End If

End Sub
Sub GenerarRecordsetTemporal()
    
    With mrs_Nacimientos
          .Fields.Append "FechaNacimiento", adChar, 16, adFldIsNullable
          .Fields.Append "EdadSemanas", adInteger
          .Fields.Append "Talla", adDouble
          .Fields.Append "Peso", adDouble
          .Fields.Append "IdCondicionRN", adInteger
          .Fields.Append "CondicionRN", adVarChar, 50
          .Fields.Append "IdtipoSexo", adInteger
          .Fields.Append "Sexo", adVarChar, 50
          .Fields.Append "Apgar_1", adInteger
          .Fields.Append "Apgar_5", adInteger
          .Fields.Append "clamplajeFecha", adChar, 16, adFldIsNullable
          .Fields.Append "NroOrdenHijoEnParto", adInteger
          .Fields.Append "NroOrdenHijo", adInteger
          .Fields.Append "IdDocIdentidad", adInteger
          .Fields.Append "docIdentidad", adVarChar, 20
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set UserControl.grdNacimientos.DataSource = mrs_Nacimientos
    
End Sub

Sub LimpiarDatos()
    On Error GoTo errLimp
    LimpiarTextos
    With mrs_Nacimientos
       If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
             .Delete
             .Update
             .MoveNext
          Loop
       End If
    End With
errLimp:
End Sub

Private Sub btnAgregar_Click()
    ActualizaDatos
End Sub

Sub ActualizaDatos()
    If Not ValidaDatosObligatorios() Then
        Exit Sub
    End If
    
    If Not ValidaReglas() Then
        Exit Sub
    End If
    Dim ldFecha As Date
    With mrs_Nacimientos
        If lbModificar = False Then
           .AddNew
        End If
        ldFecha = CDate(Format(UserControl.txtFechaNacimiento.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM))
        .Fields!FechaNacimiento = UserControl.txtFechaNacimiento.Text
        .Fields!EdadSemanas = Val(UserControl.txtEdad)
        .Fields!Talla = Val(UserControl.txtTalla)
        .Fields!Peso = Val(UserControl.txtPeso)
        .Fields!idCondicionRN = Val(mo_cmbIdCondicionRN.BoundText)
        .Fields!CondicionRN = Val(mo_cmbIdCondicionRN.BoundText)
        .Fields!idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
        .Fields!Sexo = UserControl.cmbIdTipoSexo.Text
        .Fields!apgar_1 = Val(UserControl.txtApgar1.Text)
        .Fields!apgar_5 = Val(UserControl.txtApgar5.Text)
        .Fields!clamplajeFecha = UserControl.txtFclamplaje.Text
        .Fields!NroOrdenHijoEnParto = Val(UserControl.TxtNroOrdenParto.Text)
        .Fields!NroOrdenHijo = Val(UserControl.txtNroHijo.Text)
        .Fields!IdDocIdentidad = Val(mo_cmbIdDocIdentidad.BoundText)
        .Fields!DocIdentidad = UserControl.txtNroDocumento.Text
    End With
    SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "N°hijo:" & txtNroHijo.Text
    '
    ml_idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
    mda_FechaNacimiento = CDate(UserControl.txtFechaNacimiento.Text)
    RaiseEvent SePresionoTeclaEspecial(1000)
    '
    LimpiarTextos
End Sub
Sub LimpiarTextos()
    cmbIdCondicionRN.Text = ""
    cmbIdTipoSexo.Text = ""
    txtPeso.Text = ""
    txtTalla.Text = ""
    txtEdad.Text = ""
    txtFechaNacimiento.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
    UserControl.txtApgar1.Text = ""
    UserControl.txtApgar5.Text = ""
    UserControl.TxtNroOrdenParto.Text = mrs_Nacimientos.RecordCount + 1
    UserControl.txtNroHijo.Text = ""
    UserControl.txtFclamplaje.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
    If mrs_Nacimientos.RecordCount > 0 Then
       mrs_Nacimientos.MoveLast
       UserControl.txtNroHijo.Text = mrs_Nacimientos.Fields!NroOrdenHijo + 1
    End If
    mo_cmbIdDocIdentidad.BoundText = ""
    UserControl.txtNroDocumento.Text = ""
    HabilidaBotones True
    lbModificar = False
End Sub

Function ValidaDatosObligatorios() As Boolean

    ValidaDatosObligatorios = False
    
    'Datos obligatorios
    If cmbIdCondicionRN.Text = "" Then
        MsgBox "Por favor ingreso la condición del recien nacido", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    
    If txtEdad.Text = "" Then
        MsgBox "Ingrese la edad en semanas del recien nacido", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    
    If txtTalla.Text = "" Then
        MsgBox "Ingrese la talla en cm. del recien nacido", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    
    If txtPeso.Text = "" Then
        MsgBox "Ingrese el peso en grs. del recien nacido", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    
    If cmbIdTipoSexo.Text = "" Then
        MsgBox "Ingrese el sexo del recien nacido", vbInformation, "Registro de nacimientos"
        Exit Function
    End If

    If txtFechaNacimiento = SIGHEntidades.FECHA_VACIA_DMY_HM Then
        MsgBox "Por favor ingrese la fecha de nacimiento", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    If Not (Val(TxtNroOrdenParto.Text) > 0) Then
        MsgBox "Por favor ingrese el N° EL ORDEN EN EL PARTO del Nacido Vivo o Muerto", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    If Not (Val(txtNroHijo.Text) > 0) Then
        MsgBox "Por favor ingrese el N° HIJO (de todos los Hijos que haya tenido la madre)", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    If UserControl.txtFclamplaje.Text = SIGHEntidades.FECHA_VACIA_DMY_HM Then
        MsgBox "Por favor ingrese la fecha de CLAMPAJE", vbInformation, "Registro de nacimientos"
        Exit Function
    End If

    ValidaDatosObligatorios = True
    
End Function
Function ValidaReglas() As Boolean

    ValidaReglas = False
    
    'Validar Reglas
    If txtTalla <> "" Then
        If Val(txtTalla) > 70 Then
            If MsgBox("La talla que esta ingresando es mayor que 70 cm. ¿es correcto?", vbQuestion + vbYesNo, "Registo de nacimientos") = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If txtPeso <> "" Then
        If Val(txtPeso) > 10000 Then
            If MsgBox("El peso que esta ingresando es mayor que 10 Kg. (10000 gr), ¿es correcto?", vbQuestion + vbYesNo, "Registo de nacimientos") = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If CDate(txtFechaNacimiento) < mda_FechaIngreso Then
        MsgBox "La fecha de nacimiento no puede ser menor que la fecha de ingreso", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    
'    If CDate(txtFechaNacimiento) > Date Then
'        MsgBox "La fecha de de nacimiento no puede ser mayor que la fecha de hoy", vbInformation, "Registro de nacimientos"
'        Exit Function
'    End If
'    If Not SIGHEntidades.EsFecha(UserControl.txtFclamplaje.Text, "DD/MM/AAAA") Then
'        MsgBox "La FECHA DE CLAMPLAJE no es válida", vbInformation, "Registro de nacimientos"
'        Exit Function
'    End If
    If CDate(txtFechaNacimiento.Text) > CDate(UserControl.txtFclamplaje.Text) Then
        MsgBox "La fecha de CLAMPAJE no puede ser menor que la fecha de nacimiento", vbInformation, "Registro de nacimientos"
        Exit Function
    End If
    ValidaReglas = True
End Function
Private Sub btnQuitar_Click()
    On Error Resume Next
    With mrs_Nacimientos
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
    LimpiarTextos
End Sub

Private Sub grdNacimientos_Click()
    LimpiarTextos
End Sub

Sub HabilidaBotones(lbHabilita As Boolean)
    btnAgregar.Enabled = lbHabilita
    cmdModificar.Enabled = Not lbHabilita
    btnQuitar.Enabled = Not lbHabilita
End Sub

