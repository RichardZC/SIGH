VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HerrUsuarioSistema 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Información por Usuario del Sistema"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   9195
      Begin SIGHReportes.XP_ProgressBar XP_ProgressBar1 
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   1500
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   12937777
      End
      Begin VB.ComboBox cmbUsuario 
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
         ItemData        =   "HerrUsuarioSistema.frx":0000
         Left            =   1500
         List            =   "HerrUsuarioSistema.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   690
         Width           =   2145
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
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
         Left            =   7950
         Picture         =   "HerrUsuarioSistema.frx":0026
         TabIndex        =   4
         Top             =   660
         Width           =   1125
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1500
         TabIndex        =   6
         Top             =   240
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   6930
         TabIndex        =   7
         Top             =   210
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   240
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
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   8310
         TabIndex        =   9
         Top             =   210
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
      Begin PVCOMBOLibCtl.PVComboBox cmbIdListItem 
         Height          =   300
         Left            =   1500
         TabIndex        =   13
         Top             =   1140
         Width           =   7575
         _Version        =   524288
         _cx             =   13361
         _cy             =   529
         Appearance      =   1
         Enabled         =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         Locked          =   0   'False
         Style           =   0
         Sorted          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowPictures    =   0   'False
         ColumnHeaders   =   -1  'True
         PrimaryColumn   =   1
         VisibleItems    =   10
         ColumnHeaderHeight=   20
         ListMember      =   ""
         ColumnHeaderForeColor=   0
         ColumnHeaderBackColor=   13160660
         SelectedForeColor=   16777215
         SelectedBackColor=   6956042
         AlternateBackColor=   16777215
         ItemLabelStyle  =   1
         ItemLabelType   =   0
         ItemLabelWidth  =   40
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   1
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   "[NULL]"
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   3
         Column0.Heading =   "Id"
         Column0.Width   =   20
         Column0.Alignment=   0
         Column0.Hidden  =   0   'False
         Column0.Name    =   "IdListItem"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "SubModulo"
         Column1.Width   =   200
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "SubModulo"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         Column2.Heading =   "Modulo"
         Column2.Width   =   200
         Column2.Alignment=   0
         Column2.Hidden  =   0   'False
         Column2.Name    =   "Modulo"
         Column2.Format  =   ""
         Column2.Bound   =   -1  'True
         Column2.Locked  =   0   'False
         Column2.HeaderAlignment=   0
         SortKey1.Column =   -1
         SortKey1.Ascending=   -1  'True
         SortKey1.CaseInsensitive=   -1  'True
         SortKey2.Column =   -1
         SortKey2.Ascending=   -1  'True
         SortKey2.CaseInsensitive=   -1  'True
         SortKey3.Column =   -1
         SortKey3.Ascending=   -1  'True
         SortKey3.CaseInsensitive=   -1  'True
         BoundColumn     =   ""
         Border          =   -1  'True
         VertAlign       =   1
         Format          =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Opción Sistema"
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
         TabIndex        =   14
         Top             =   1140
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Usuario Sistema"
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
         TabIndex        =   12
         Top             =   750
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
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
         TabIndex        =   11
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
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
         Left            =   6420
         TabIndex        =   10
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   0
      Top             =   1980
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HerrUsuarioSistema.frx":0338
         DownPicture     =   "HerrUsuarioSistema.frx":0798
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
         Left            =   3158
         Picture         =   "HerrUsuarioSistema.frx":0C0D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrUsuarioSistema.frx":1082
         DownPicture     =   "HerrUsuarioSistema.frx":1546
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
         Left            =   4688
         Picture         =   "HerrUsuarioSistema.frx":1A32
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "HerrUsuarioSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Auditoría en General para todos los módulos del Sistema
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbUsuario As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim lnIdAlmacen As Long
Dim ml_idUsuario As Long
Private WithEvents oRpt As RptHerrUsuarioSistema
Attribute oRpt.VB_VarHelpID = -1


Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property



Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:    F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ")" & IIf(cmbUsuario.Text = "", "", "     Usuario: " & cmbUsuario.Text)
    ml_TextoDelFiltro = ml_TextoDelFiltro & IIf(Me.cmbIdListItem.Text = "", "", "   (Opción: " & Trim(Me.cmbIdListItem.Text) & ")")
    
    If Me.txtFdesde = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de movimiento inicial"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFdesde, "DD/MM/AAAA") Then
            sMensaje = "La fecha de movimiento inicial, no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFhasta = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de movimiento final"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFhasta, "DD/MM/AAAA") Then
            sMensaje = "La fecha de movimiento final, no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrInicio = SIGHEntidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora de movimiento inicial"
    Else
        If Not SIGHEntidades.EsHora(txtHrInicio) Then
            sMensaje = "La hora de movimiento inicial, no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrFin = SIGHEntidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora de movimiento final"
    Else
        If Not SIGHEntidades.EsHora(txtHrFin) Then
            sMensaje = "La hora de movimiento final, no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
    End If
    
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function




Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub


Private Sub cmbUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
      mo_Teclado.RealizarNavegacion KeyCode, cmbUsuario
End Sub



Private Sub Form_Initialize()
    Set mo_cmbUsuario.MiComboBox = cmbUsuario
End Sub


Private Sub Form_Load()
    '
    Set oRpt = New RptHerrUsuarioSistema
    XP_ProgressBar1.ShowText = True
    '
    txtFdesde.Text = Date
    txtFhasta.Text = Date
    txtHrInicio.Text = "00:01"
    txtHrFin.Text = "23:59"
    '
    Set cmbIdListItem.ListSource = mo_AdminSeguridad.ListItemsSeleccionarTodos()
    '
    mo_cmbUsuario.BoundColumn = "IdEmpleado"
    mo_cmbUsuario.ListField = "Usuario"
    Set mo_cmbUsuario.RowSource = mo_ReglasFarmacia.DevuelveListaDeUsuariosDelSistema
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub






Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub oRpt_ProgressActualizaValor(lnValorActual As Long, lnValorTotal As Long)
    XP_ProgressBar1.Max = lnValorTotal
    XP_ProgressBar1.Min = 0
    XP_ProgressBar1.Value = lnValorActual
    DoEvents
    Me.Refresh
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdesde = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFhasta = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbUsuario = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_reglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub



Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        'Dim oRpt As New RptHerrUsuarioSistema
        Dim oCampos() As String
        Dim lnIdListItem As Long
        If Me.cmbIdListItem.Text <> "" Then
           oCampos = Split(Me.cmbIdListItem.List(cmbIdListItem.ListIndex), "|")
           lnIdListItem = oCampos(0)
        Else
           lnIdListItem = 0
        End If
        oRpt.CreaDatosParaReporte IIf(chkExcel.Value = 1, True, False), "Auditoria en General", ml_TextoDelFiltro, Val(mo_cmbUsuario.BoundText), CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), lnIdListItem, Me.hwnd
        Me.MousePointer = 1
    End If
End Sub


Private Sub txtHrFin_LostFocus()
    If txtHrFin <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.EsHora(txtHrFin) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrFin = SIGHEntidades.HORA_VACIA_HM
        End If
    End If
End Sub

Private Sub txtHrInicio_LostFocus()
    If txtHrInicio <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.EsHora(txtHrInicio) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrInicio = SIGHEntidades.HORA_VACIA_HM
        End If
    End If
End Sub
