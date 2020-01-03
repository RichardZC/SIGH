VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucDiagnosticoDetalle 
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10485
   ScaleHeight     =   5070
   ScaleWidth      =   10485
   Begin VB.Frame fraDiagnostico 
      Caption         =   "Diagnósticos     ( F1=Todos Dx )"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton btnBusquedaDiagnostico 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2535
         Picture         =   "ucDiagnosticoDetalle.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Buscar"
         Top             =   225
         Width           =   375
      End
      Begin Threed.SSCommand btnAgregarDx 
         Height          =   465
         Left            =   7770
         TabIndex        =   4
         Top             =   600
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
         Picture         =   "ucDiagnosticoDetalle.ctx":058A
         Caption         =   "Agregar"
         PictureAlignment=   9
      End
      Begin VB.ComboBox cmbIdTipoDiagnostico 
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   630
         Width           =   2805
      End
      Begin VB.TextBox lblDescripcionDx 
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
         Left            =   2925
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   7365
      End
      Begin VB.TextBox txtIdDiagnostico 
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
         Left            =   1500
         TabIndex        =   0
         Top             =   240
         Width           =   1005
      End
      Begin Threed.SSCommand btnQuitarDx 
         Height          =   465
         Left            =   9120
         TabIndex        =   5
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "ucDiagnosticoDetalle.ctx":3516
         Caption         =   "Quitar"
         PictureAlignment=   9
         ShapeSize       =   1
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbLabHis 
         Height          =   330
         Left            =   5910
         TabIndex        =   3
         Top             =   630
         Visible         =   0   'False
         Width           =   1065
         _Version        =   524288
         _cx             =   1879
         _cy             =   582
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
         ItemLabelWidth  =   20
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   0
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
         Column0.Width   =   10
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdHisSituacio"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Valores"
         Column1.Width   =   35
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "valores"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         Column2.Heading =   "Descripción"
         Column2.Width   =   100
         Column2.Alignment=   0
         Column2.Hidden  =   0   'False
         Column2.Name    =   "descripcio"
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
      Begin VB.Label lblLabConfHIS 
         Alignment       =   1  'Right Justify
         Caption         =   "Lab (HIS)"
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
         Left            =   5010
         TabIndex        =   10
         Top             =   660
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label16 
         Caption         =   "Diagnóstico"
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
         Left            =   135
         TabIndex        =   9
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label lblDiagnostico 
         Caption         =   "Tipo diagnóstico"
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
         Left            =   105
         TabIndex        =   8
         Top             =   630
         Width           =   1470
      End
   End
   Begin UltraGrid.SSUltraGrid grdDiagnosticos 
      Height          =   3855
      Left            =   30
      TabIndex        =   6
      Top             =   1170
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6800
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
      Caption         =   "Lista de diagnósticos"
   End
End
Attribute VB_Name = "ucDiagnosticoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registrar Diagnósticos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idUsuario As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ms_MensajeError As String
Dim mrs_Diagnosticos As New ADODB.Recordset
Dim ml_TipoDiagnostico As sghTiposDiagnostico
Dim mo_cmbIdTipoDiagnostico As New sighentidades.ListaDespleglable
Dim ml_SexoPaciente As Long
Dim ml_EdadPaciente As Long
Dim ml_IdListBarItem As Long
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Public Event SeIngresoDx(lcDx As String, SeElimino As Boolean)

Property Let IdListBarItem(lValue As Long)
   ml_IdListBarItem = lValue
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
Property Let TituloFrame(sValue As String)
   fraDiagnostico.Caption = sValue
End Property
Property Let TipoDiagnostico(lValue As sghTiposDiagnostico)
    ml_TipoDiagnostico = lValue
End Property
Property Let SexoPaciente(lValue As Integer)
    ml_SexoPaciente = lValue
End Property
Property Let EdadPaciente(lValue As Long)
    ml_EdadPaciente = lValue
End Property
'mgaray201410c
Property Get rsDiagnosticos() As ADODB.Recordset
   Set rsDiagnosticos = mrs_Diagnosticos
End Property

Property Let BotonAgregarEnabled(bValue As Boolean)
    UserControl.btnAgregarDx.Enabled = False
End Property
Property Let BotonQuitarEnabled(bValue As Boolean)
    UserControl.btnQuitarDx.Enabled = False
End Property

Private Sub cmbIdTipoDiagnostico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoDiagnostico
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbIdTipoDiagnostico_LostFocus()
   If cmbIdTipoDiagnostico.Text <> "" Then
        Dim lIdTipoDiagnostico As Long
        lIdTipoDiagnostico = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarIdPorCodigoYClasificacion(UCase(Split(cmbIdTipoDiagnostico.Text, " = ")(0)), ml_TipoDiagnostico)
       ' mo_cmbIdTipoDiagnostico.BoundText = lIdTipoDiagnostico
   End If
End Sub


Sub ChequeaDxVSpacienteEdadSexo()
         Dim lbContinuar As Boolean
         lbContinuar = True
         If UserControl.txtIdDiagnostico.Text <> "" And lbContinuar = True Then
             Dim oDODiagnostico As DODiagnostico
             If ml_IdListBarItem = sghOpcionGalenHos.sghRegistroAtencionCE Then
                Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(UserControl.txtIdDiagnostico.Text, False)
             Else
                Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(UserControl.txtIdDiagnostico.Text, True)
             End If
             If Not oDODiagnostico Is Nothing Then
                 UserControl.txtIdDiagnostico.Tag = oDODiagnostico.idDiagnostico
                 UserControl.lblDescripcionDx = oDODiagnostico.descripcion
                 If oDODiagnostico.Restriccion Then
                     If oDODiagnostico.idTipoSexo <> 0 Then
                         If ml_SexoPaciente <> oDODiagnostico.idTipoSexo Then
                             MsgBox "El diagnóstico no corresponde al sexo del paciente", vbInformation, "Validación paciente"
                             UserControl.txtIdDiagnostico.Tag = ""
                             UserControl.lblDescripcionDx = ""
                             Exit Sub
                         End If
                         If ml_SexoPaciente = 1 And oDODiagnostico.Gestacion = True Then
                             MsgBox "El diagnóstico de gestación no corresponde al sexo del paciente ", vbInformation, "Validación paciente"
                             UserControl.txtIdDiagnostico.Tag = ""
                             UserControl.lblDescripcionDx = ""
                             Exit Sub
                         End If
                     End If
                     If ml_EdadPaciente < 3650 And oDODiagnostico.Gestacion = True Then
                         MsgBox "El diagnóstico de gestación no corresponde a la edad del paciente ", vbInformation, "Validación paciente"
                         UserControl.txtIdDiagnostico.Tag = ""
                         UserControl.lblDescripcionDx = ""
                         Exit Sub
                     End If
                     If (ml_EdadPaciente > oDODiagnostico.EdadMaxDias) Or (ml_EdadPaciente < oDODiagnostico.EdadMinDias) Then
                         MsgBox "El diagnóstico no corresponde a la edad del paciente (Edad mínima " & oDODiagnostico.EdadMinDias & " días - Edad máxima " & oDODiagnostico.EdadMaxDias & " días)", vbInformation, "Validación paciente"
                         UserControl.txtIdDiagnostico.Tag = ""
                         UserControl.lblDescripcionDx = ""
                         Exit Sub
                     End If
                 End If
             Else
                 UserControl.txtIdDiagnostico.Tag = ""
                 UserControl.lblDescripcionDx = ""
             End If
       End If
       mo_Formulario.MarcarComoVacio txtIdDiagnostico
End Sub

Public Function Inicializar()
    GenerarRecordsetTemporal
    mo_Formulario.HabilitarDeshabilitar UserControl.lblDescripcionDx, False
    Set mo_cmbIdTipoDiagnostico.MiComboBox = cmbIdTipoDiagnostico
    On Error Resume Next
    txtIdDiagnostico.SetFocus
End Function

Public Sub FocusEnDx()
    On Error Resume Next
    txtIdDiagnostico.SetFocus
End Sub




Private Sub cmbLabHis_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbLabHis
    RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub







Private Sub UserControl_Resize()
    
    UserControl.lblDescripcionDx.Width = UserControl.Width - 3070
    fraDiagnostico.Width = UserControl.Width - 20
    UserControl.grdDiagnosticos.Width = UserControl.Width - 20
    UserControl.grdDiagnosticos.Height = UserControl.Height - 1100
    
End Sub



Public Sub ConfigurarComboBoxes()
Dim sMensaje As String
       
       
        
        '
        mo_cmbIdTipoDiagnostico.BoundColumn = "IdSubclasificacionDx"
        mo_cmbIdTipoDiagnostico.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoDiagnostico.MiComboBox = cmbIdTipoDiagnostico
        Select Case ml_TipoDiagnostico
        Case sghAtencionConsultaExterna
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxConsultaExterna
        Case sghHospitalizacionIngreso
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospIngreso
        Case sghHospitalizacionEgreso
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospEgreso
        Case sghHospitalizacionMortalidad
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospMortalidad
        Case sghHospitalizacionNacimiento
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospMuerteFetal
        Case sghHospitalizacionComplicaciones
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospComplicaciones
        Case sghInterconsultas
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxInterconsultas
        End Select
        
        Dim rsTipoDiagnostico As New Recordset
        Set rsTipoDiagnostico = mo_cmbIdTipoDiagnostico.RowSource
        Select Case rsTipoDiagnostico.RecordCount
        Case 0
            lblDiagnostico.Visible = False
            cmbIdTipoDiagnostico.Visible = False
        Case 1
        Case 2
        End Select
        
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError

End Sub
Private Sub btnBusquedaDiagnostico_Click()
    BusquedaDx ""
End Sub

Sub BusquedaDx(lcCodigoDx As String)
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    If ml_IdListBarItem = sghOpcionGalenHos.sghRegistroAtencionCE Then
       oBusqueda.SoloMuestraDxGalenHos = False
    Else
       oBusqueda.SoloMuestraDxGalenHos = True
    End If
    oBusqueda.CodigoDx = lcCodigoDx
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            UserControl.txtIdDiagnostico.Text = oDODiagnostico.CodigoCIE2004
            UserControl.txtIdDiagnostico.Tag = oDODiagnostico.idDiagnostico
            UserControl.lblDescripcionDx = oDODiagnostico.descripcion
        Else
            UserControl.txtIdDiagnostico.Text = ""
            UserControl.txtIdDiagnostico.Tag = ""
            UserControl.lblDescripcionDx = ""
        End If
    Else
        UserControl.txtIdDiagnostico.Text = ""
        UserControl.txtIdDiagnostico.Tag = ""
        UserControl.lblDescripcionDx = ""
    End If
    Set oBusqueda = Nothing
End Sub


Private Sub grdDiagnosticos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdDiagnosticos.Bands(0).Columns("IdTipoDiagnostico").Hidden = True
    grdDiagnosticos.Bands(0).Columns("labConfHIS").Header.Caption = "Lab (HIS)"
    
    Select Case ml_TipoDiagnostico
    Case sghHospitalizacionIngreso, sghHospitalizacionComplicaciones
        grdDiagnosticos.Bands(0).Columns("DescripcionTipoDx").Hidden = True
    Case sghAtencionConsultaExterna, sghHospitalizacionEgreso, sghHospitalizacionMortalidad, sghHospitalizacionNacimiento
        grdDiagnosticos.Bands(0).Columns("DescripcionTipoDx").Header.Caption = "Tipo diagnóstico"
        grdDiagnosticos.Bands(0).Columns("DescripcionTipoDx").Width = 2000
        grdDiagnosticos.Bands(0).Columns("DescripcionTipoDx").Activation = ssActivationActivateNoEdit 'Actualizado 25092014
    End Select
    
    grdDiagnosticos.Bands(0).Columns("IdDiagnostico").Hidden = True
    
    grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Header.Caption = "CIE"
    grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Width = 1000
    grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Activation = ssActivationActivateNoEdit 'Actualizado 25092014
    
    grdDiagnosticos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdDiagnosticos.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit 'Actualizado 25092014

    If ml_TipoDiagnostico <> sghAtencionConsultaExterna Then
       grdDiagnosticos.Bands(0).Columns("labConfHIS").Hidden = True
       grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 8000
    Else
       grdDiagnosticos.Bands(0).Columns("labConfHIS").Activation = ssActivationActivateNoEdit 'Actualizado 25092014
       grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 7000
    End If
    
  mo_Apariencia.ConfigurarFilasBiColores grdDiagnosticos, sighentidades.GrillaConFilasBicolor
End Sub
Private Sub txtIdDiagnostico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdDiagnostico
    If KeyCode = vbKeyF1 Then
        btnBusquedaDiagnostico_Click
        
    End If
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtIdDiagnostico_LostFocus()
   If Len(txtIdDiagnostico.Text) > 0 And lblDescripcionDx.Text = "" Then
      BusquedaDx txtIdDiagnostico.Text
   End If
   UserControl.txtIdDiagnostico.Text = UCase(UserControl.txtIdDiagnostico.Text)
End Sub

Private Sub txtIdDiagnostico_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CargarDatosDeDiagnosticos(oConexion As Connection)
Dim rsDiagnosticos As New Recordset

    Set rsDiagnosticos = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(ml_idAtencion, ml_TipoDiagnostico, oConexion)
    Do While Not rsDiagnosticos.EOF
        With mrs_Diagnosticos
            .AddNew
            .Fields!idTipoDiagnostico = rsDiagnosticos!idTipoDiagnostico
            .Fields!DescripcionTipoDx = rsDiagnosticos!DescripcionTipoDx
            .Fields!idDiagnostico = rsDiagnosticos!idDiagnostico
            .Fields!CodigoCIE2004 = rsDiagnosticos!CodigoCIE2004
            .Fields!descripcion = rsDiagnosticos!descripcion
            .Fields!labConfHIS = rsDiagnosticos!labConfHIS
        End With
        rsDiagnosticos.MoveNext
    Loop
    On Error Resume Next
    rsDiagnosticos.MoveFirst
    
    Set rsDiagnosticos = Nothing
End Sub


Sub CargarDiagnosticosAlObjetoDatos(oDiagnosticos As Collection)
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS DIAGNOSTICOS
    '---------------------------------------------------------------------------------
    Dim oRow As SSRow
    Dim oDiagnostico As DOAtencionDiagnostico
    Set oRow = UserControl.grdDiagnosticos.GetRow(ssChildRowFirst)
    
    If Not oRow Is Nothing Then
        'Para el primero
        Set oDiagnostico = New DOAtencionDiagnostico
        oDiagnostico.IdAtencionDiagnostico = 0
        oDiagnostico.idAtencion = ml_idAtencion
        oDiagnostico.idDiagnostico = oRow.Cells("IdDiagnostico").Value
        oDiagnostico.IdClasificacionDx = ml_TipoDiagnostico
        oDiagnostico.IdSubclasificacionDx = IIf(IsNull(oRow.Cells("IdTipoDiagnostico").Value), 0, oRow.Cells("IdTipoDiagnostico").Value)
        oDiagnostico.IdUsuarioAuditoria = ml_idUsuario
        oDiagnostico.labConfHIS = IIf(IsNull(oRow.Cells("labConfHIS").Value), "", oRow.Cells("labConfHIS").Value)
        oDiagnosticos.Add oDiagnostico
        
        'Para los siguientes
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            Set oDiagnostico = New DOAtencionDiagnostico
            oDiagnostico.IdAtencionDiagnostico = 0
            oDiagnostico.idAtencion = ml_idAtencion
            oDiagnostico.idDiagnostico = oRow.Cells("IdDiagnostico").Value
            oDiagnostico.IdClasificacionDx = ml_TipoDiagnostico
            oDiagnostico.IdSubclasificacionDx = IIf(IsNull(oRow.Cells("IdTipoDiagnostico").Value), 0, oRow.Cells("IdTipoDiagnostico").Value)
            oDiagnostico.IdUsuarioAuditoria = ml_idUsuario
            oDiagnostico.labConfHIS = IIf(IsNull(oRow.Cells("labConfHIS").Value), "", oRow.Cells("labConfHIS").Value)
            oDiagnosticos.Add oDiagnostico
        Loop
    End If
   

End Sub
Sub GenerarRecordsetTemporal()
    If mrs_Diagnosticos.State = 1 Then Set mrs_Diagnosticos = Nothing
    With mrs_Diagnosticos
          .Fields.Append "IdTipoDiagnostico", adInteger, 4, adFldIsNullable
          .Fields.Append "DescripcionTipoDx", adVarChar, 100, adFldIsNullable
          .Fields.Append "IdDiagnostico", adInteger
          .Fields.Append "CodigoCIE2004", adVarChar, 10
          .Fields.Append "Descripcion", adVarChar, 255
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set UserControl.grdDiagnosticos.DataSource = mrs_Diagnosticos
    mo_Apariencia.ConfigurarFilasBiColores UserControl.grdDiagnosticos, sighentidades.GrillaConFilasBicolor
End Sub

Sub LimpiarDatos()
    On Error GoTo errLimp
    With mrs_Diagnosticos
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

Private Sub btnAgregarDx_Click()
    
    
    
    ChequeaDxVSpacienteEdadSexo
    
    If UserControl.txtIdDiagnostico.Text = "" Then
        MsgBox "Por favor ingrese el diagnóstico", vbInformation, "Diagnósticos"
        Exit Sub
    End If
    
    If UserControl.txtIdDiagnostico.Tag = "" Then
        MsgBox "Por favor ingrese un diagnóstico válido", vbInformation, "Diagnósticos"
        Exit Sub
    End If
    
    If cmbIdTipoDiagnostico.Visible = True Then
        If UserControl.cmbIdTipoDiagnostico.Text = "" Then
            MsgBox "Por favor ingrese el tipo de diagnóstico", vbInformation, "Diagnósticos"
            Exit Sub
        End If
    End If
    
    '***************daniel barrantes**************
    '***************Valida Diagnosticos REPETIDOS
    '***************
    
    If mrs_Diagnosticos.RecordCount > 0 Then
        mrs_Diagnosticos.MoveFirst
        Do While Not mrs_Diagnosticos.EOF
        
        'Yamill Palomino
        If ml_IdListBarItem = sghOpcionGalenHos.sghAdmisionEmergencia Or ml_IdListBarItem = sghOpcionGalenHos.sghAdmisionHospitalizacion Then
            If txtIdDiagnostico.Tag = mrs_Diagnosticos!idDiagnostico Then
                MsgBox "El diagnóstico ya fue agregado al listado", vbInformation, "Admision"
                Exit Sub
            End If
        Else
        'Actualizado 29092014
'            If txtIdDiagnostico.Tag = mrs_Diagnosticos!IdDiagnostico And Trim(cmbLabHis.Text) = mrs_Diagnosticos!labConfHIS And mrs_Diagnosticos!idTipoDiagnostico = Val(mo_cmbIdTipoDiagnostico.BoundText) Then
            If txtIdDiagnostico.Tag = mrs_Diagnosticos!idDiagnostico And Trim(cmbLabHis.Text) = IIf(IsNull(mrs_Diagnosticos!labConfHIS) = True, "", Trim(mrs_Diagnosticos!labConfHIS)) Then
                If Trim(cmbLabHis.Text) = "" Then
                    MsgBox "El diagnóstico ya fué agregado al listado", vbInformation, "Admisión"
                Else
                    MsgBox "El diagnóstico con el mismo codigo lab ya fué registrado", vbInformation, "Admisión"
                End If
                Exit Sub
            End If
        End If
            mrs_Diagnosticos.MoveNext
        Loop
    End If

    With mrs_Diagnosticos
        .AddNew
        .Fields!idDiagnostico = Val(UserControl.txtIdDiagnostico.Tag)
        .Fields!CodigoCIE2004 = UserControl.txtIdDiagnostico.Text
        .Fields!descripcion = UserControl.lblDescripcionDx
        .Fields!idTipoDiagnostico = Val(mo_cmbIdTipoDiagnostico.BoundText)
        .Fields!DescripcionTipoDx = UserControl.cmbIdTipoDiagnostico.Text
        .Fields!labConfHIS = Right(Trim(cmbLabHis.Text), 3)
    End With
    If Val(mo_cmbIdTipoDiagnostico.BoundText) = 301 Or Val(mo_cmbIdTipoDiagnostico.BoundText) = 303 Then
       RaiseEvent SeIngresoDx(mrs_Diagnosticos!CodigoCIE2004, False)
    End If
    txtIdDiagnostico.Tag = ""
    txtIdDiagnostico.Text = ""
    lblDescripcionDx = ""
    mo_cmbIdTipoDiagnostico.BoundText = ""
'    cmbIdTipoDiagnostico.Text = ""
    cmbIdTipoDiagnostico.ListIndex = -1
    cmbLabHis.Text = ""
    
    On Error Resume Next
    sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Dx: " & UserControl.txtIdDiagnostico.Text
    
    RaiseEvent SePresionoTeclaEspecial(vbKeyTab)
    If ml_IdListBarItem = sghOpcionGalenHos.sghRegistroAtencionCE Then  'CE-Reg.Dx
       FocusEnDx
    End If

End Sub

'Actualizado 15102014
Private Sub btnQuitarDx_Click()


    EliminarDiagnosticoSeleccionado
End Sub

Public Sub EditaLabConfHIS()
    cmbLabHis.Visible = True
    lblLabConfHIS.Visible = True
    grdDiagnosticos.Bands(0).Columns("labConfHIS").Width = 1000
    grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 1000
    'debb-9-2-211
    Set cmbLabHis.ListSource = mo_AdminServiciosComunes.DevuelveHIS_SITUACIOporDescripcion()
    'debb-9-2-211
End Sub

Public Sub TipoDxDefault(lcValorDefault As String)            'debb-06-03-2012
    If cmbIdTipoDiagnostico.Text = "" Then
       mo_cmbIdTipoDiagnostico.BoundText = lcValorDefault
    End If
End Sub

'Actualizado 25092014
Private Sub grdDiagnosticos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    EliminarDiagnosticoSeleccionado
End Sub

Public Sub EliminarDiagnosticoSeleccionado()
    If mrs_Diagnosticos.RecordCount > 0 Then
        If MsgBox("¿Desea eliminar el diagnóstico seleccionado?", vbYesNo, "Eliminar diagnósticos") = vbYes Then
            On Error Resume Next
            Dim lcDx11 As String
            lcDx11 = ""
            If mrs_Diagnosticos.RecordCount > 0 Then
                With mrs_Diagnosticos
                    If Not .EOF And Not .BOF Then
                       lcDx11 = .Fields!CodigoCIE2004
                       .Delete
                       .Update
                    End If
                End With
            End If
            Set grdDiagnosticos.DataSource = mrs_Diagnosticos
            If lcDx11 <> "" Then
               RaiseEvent SeIngresoDx(lcDx11, True)
            End If
        End If
    End If
End Sub

'mgaray20141008
Public Function DeshabilitarEdicionDatos() As Boolean
    fraDiagnostico.Enabled = False
    grdDiagnosticos.Enabled = False
    btnBusquedaDiagnostico.Enabled = False
End Function

Public Function HabilitarEdicionDatos() As Boolean
    fraDiagnostico.Enabled = True
    grdDiagnosticos.Enabled = True
    btnBusquedaDiagnostico.Enabled = True
End Function

'A.Yañez*************************************
Public Function limpiacampos() As Boolean
     txtIdDiagnostico.Text = ""
     lblDescripcionDx.Text = ""
End Function
'*********************************************


'debb-23/02/2015
Sub CargarDatosDeDiagnosticosEmergCE(oConexion As Connection, lnTipoDiagnostico As Long)
Dim rsDiagnosticos As New Recordset

    Set rsDiagnosticos = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(ml_idAtencion, lnTipoDiagnostico, oConexion)
    Do While Not rsDiagnosticos.EOF
        With mrs_Diagnosticos
            .AddNew
            '.Fields!idTipoDiagnostico = rsDiagnosticos!idTipoDiagnostico    'debb-22/07/2016
            .Fields!DescripcionTipoDx = rsDiagnosticos!DescripcionTipoDx
            .Fields!idDiagnostico = rsDiagnosticos!idDiagnostico
            .Fields!CodigoCIE2004 = rsDiagnosticos!CodigoCIE2004
            .Fields!descripcion = rsDiagnosticos!descripcion
            .Fields!labConfHIS = rsDiagnosticos!labConfHIS
        End With
        rsDiagnosticos.MoveNext
    Loop
    On Error Resume Next
    rsDiagnosticos.MoveFirst
    
    Set rsDiagnosticos = Nothing
End Sub
Function DevuelveDx() As Recordset
   Set DevuelveDx = mrs_Diagnosticos.Clone
End Function

