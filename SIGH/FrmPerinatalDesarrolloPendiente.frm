VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FrmPerinatalDesarrolloPendiente 
   Caption         =   "Sesiones de desarrollo no ejecutadas en atenciones"
   ClientHeight    =   5505
   ClientLeft      =   5130
   ClientTop       =   2760
   ClientWidth     =   11325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPerinatalDesarrolloPendiente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   11325
   Begin UltraGrid.SSUltraGrid grdPlanDesarrollo 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   7435
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sesiones Desarrollo Psicomotor No Ejecutadas"
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   11115
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cerrar (ESC)"
         DisabledPicture =   "FrmPerinatalDesarrolloPendiente.frx":000C
         DownPicture     =   "FrmPerinatalDesarrolloPendiente.frx":04D0
         Height          =   705
         Left            =   4440
         Picture         =   "FrmPerinatalDesarrolloPendiente.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frAtencionDesarrollo 
      Caption         =   "Sesión 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4335
      Left            =   4320
      TabIndex        =   8
      Top             =   0
      Width           =   6930
      Begin VB.CommandButton cmdBuscarEstablecimiento 
         Caption         =   "..."
         Height          =   300
         Left            =   6360
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FrmPerinatalDesarrolloPendiente.frx":0EA8
         DownPicture     =   "FrmPerinatalDesarrolloPendiente.frx":1308
         Height          =   705
         Left            =   2880
         Picture         =   "FrmPerinatalDesarrolloPendiente.frx":177D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3550
         Width           =   1725
      End
      Begin VB.TextBox txtEvalucionDesarrollo 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
         Top             =   1080
         Width           =   1125
      End
      Begin VB.TextBox txtFechaProgramadaDesarrollo 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
         Top             =   720
         Width           =   1125
      End
      Begin VB.TextBox txtIdAtencionDesarrollo 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
         Top             =   720
         Visible         =   0   'False
         Width           =   1125
      End
      Begin UltraGrid.SSUltraGrid grdPlanDesarrolloPendientes 
         Height          =   2070
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   3651
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
         Caption         =   "Items a Evaluar"
      End
      Begin MSMask.MaskEdBox mskFechaEjecucionDes 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   720
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbIdEstablecimiento 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   4665
         _Version        =   524288
         _cx             =   8229
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
            Size            =   8.25
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
         ColumnCount     =   2
         Column0.Heading =   "Id"
         Column0.Width   =   10
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdEstablecimiento"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Valores"
         Column1.Width   =   35
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "Nombre"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   380
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ejecución"
         Height          =   210
         Left            =   3960
         TabIndex        =   12
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Evaluacion"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Programada"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1500
      End
   End
End
Attribute VB_Name = "FrmPerinatalDesarrolloPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Mantenimiento para Desarrollo Psicomotor de sesiones vencidas
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IdPaciente As Long
Dim ml_idAtencion As Long
Dim ml_FechaAtencion As Date
Dim ml_idUsuario As Long
Dim ml_YaCargoUnaSolaVez As Boolean
Dim oEdad As Edad
Dim md_fechaNacimiento As Date
Dim md_fechaActual As Date
Dim noEjecutarAccion As Boolean
'para alamcenar la celda activa
Dim ssRowActivate As SSRow
Dim ssCellActivate As SSCell
Dim mo_RsDesarrolloPendiente As ADODB.Recordset
Dim ml_IdEstablecimiento As Long
Dim ms_MensajeError As String
Dim oDOAtenIntePlanDesPaciente As New DOAtenIntePlanDesPaciente
'mgaray201411b
Dim mb_EstaMarcadoEjecucionPsicomotor As Boolean


Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property

Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property

Property Let FechaAtencion(lValue As Date)
   ml_FechaAtencion = lValue
'   calcularEdadPaciente
End Property

Property Let FechaNacimiento(lValue As Date)
   md_fechaNacimiento = lValue
'   calcularEdadPaciente
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Public Sub Inicializar()
    If ml_YaCargoUnaSolaVez = False Then
        ml_YaCargoUnaSolaVez = True
'        CreaTemporales
        
        mo_Formulario.HabilitarDeshabilitar txtFechaProgramadaDesarrollo, False
        mo_Formulario.HabilitarDeshabilitar txtEvalucionDesarrollo, False
'        mo_Formulario.HabilitarDeshabilitar mskFechaEjecucionDes, False
        Call ConfigurarCombos
    End If
    
    Dim oDoEstablecimiento As New DOEstablecimiento
    Dim sCodigoRenaes As String
    sCodigoRenaes = lcBuscaParametro.SeleccionaFilaParametro(280)
    
    If mo_ReglasComunes.EstablecimientosSeleccionarPorCodigo(sCodigoRenaes, oDoEstablecimiento) = True Then
        ml_IdEstablecimiento = oDoEstablecimiento.IdEstablecimiento
    Else
        MsgBox "Codigo RENAES " & sCodigoRenaes & " No Encontrado en la Lista de Establecimientos, revise tabla parametros(280)", vbInformation, "Modulo Niño Sano"
        
    End If
    
    Call initializeControls
    Call cargarDatosAtencionIntegralDesarrollo
End Sub

Private Sub btnAceptar_Click()
    If ValidarDatosIngreso() = False Then
        Exit Sub
    End If
'    Private Function SetDatosEjecucionDesarrollo(cambioEjecucion As Boolean, _
'        oRsDesarrollo As ADODB.Recordset) As ADODB.Recordset
'    If cambioEjecucion = True Then
    Dim oReglasAdmision As New ReglasAdmision
    Dim oCampos() As String
    
    oCampos = Split(cmbIdEstablecimiento.List(cmbIdEstablecimiento.ListIndex), "|")
    
    oDOAtenIntePlanDesPaciente.FechaEjecucion = mskFechaEjecucionDes.Text
    oDOAtenIntePlanDesPaciente.IdEstablecimiento = Val(oCampos(0))
    oDOAtenIntePlanDesPaciente.evaluacion = Me.txtEvalucionDesarrollo.Tag
    
    If oReglasAdmision.grabarAtencionIntegralDesarrolloVencidos(ml_idAtencion, oDOAtenIntePlanDesPaciente, getRecorsetDesarrollo()) = True Then
        btnAceptar.Enabled = False
    End If
End Sub

Private Sub btnCancelar_Click()
    If validarCambiosPendientes() = False Then
        Exit Sub
    End If
    Err = 0
    Unload Me
End Sub


Private Sub cmbIdEstablecimiento_Change()
    Call HabilitarGuardar
End Sub

Private Sub cmbIdEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub grdPlanDesarrollo_AfterRowActivate()
    Dim oSSRowActive As SSRow
    Set oSSRowActive = grdPlanDesarrollo.ActiveRow
    If Not (oSSRowActive Is Nothing) Then
        Call AsignarDatosAControlesDesarrollo(oSSRowActive.Cells("IdPlanIntegralPaciente").Value, _
                            oSSRowActive.Cells("IdPlanDesarrolloPaciente").Value)
        
    End If
End Sub


Private Sub grdPlanDesarrollo_BeforeRowDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    If validarCambiosPendientes() = False Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub grdPlanDesarrollo_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPlanDesarrollo_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanDesarrollo.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanDesarrollo.UpdateMode = ssUpdateOnUpdate
    grdPlanDesarrollo.CollapseAll
    'Cabecera de grupo
    With Layout.Bands(0)
        '.ColHeadersVisible = False
        'establecer etiqueta de columnas y formato
        .Columns("IdPlanAtencion").Header.Caption = "Id Plan"
        
        .Columns("FechaProgramada").Header.Caption = "F. Programada"
        .Columns("FechaProgramada").Width = 1400
                
        .Columns("FechaEjecucion").Header.Caption = "F. Ejecución"
        .Columns("FechaEjecucion").Width = 0
        
        .Columns("NumeroSesion").Header.Caption = "N° Sesión"
        .Columns("NumeroSesion").Width = 1000
        
        .Columns("Evaluacion").Header.Caption = "Evaluación"
        .Columns("Evaluacion").Width = 0
        
        .Columns("Descripcion").Header.Caption = "Edad"
        .Columns("Descripcion").Width = grdPlanDesarrollo.Width - 300 - _
                                    .Columns("FechaProgramada").Width - _
                                    .Columns("FechaEjecucion").Width - _
                                    .Columns("NumeroSesion").Width - _
                                    .Columns("Evaluacion").Width
        
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 0, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "EdadAnio", "EdadMes", "EdadDia", _
                                        "IdEstablecimiento", "Establecimiento", _
                                        "FechaEjecucion", "Evaluacion", "EvaluacionDesc", _
                                        "IdAtencion")
        
        
        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 0, ssActivationActivateNoEdit, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "EdadAnio", "EdadMes", "EdadDia", _
                                        "FechaEjecucion", "NumeroSesion", "Evaluacion", _
                                        "IdEstablecimiento", "Establecimiento", "FechaProgramada")
                                        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 0, ssAlignCenter, _
                                "FechaEjecucion", "FechaProgramada", "NumeroSesion")
    
    End With
End Sub

Private Sub grdPlanDesarrollo_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
'    If Row.HasParent = False Then
'        If frAtenInteDesarrollo.Tag = "1" Then
'            Call activarEdicionPlanIntegral(Row)
'        Else
'            Row.Cells("FechaProgramada").Activation = ssActivationActivateNoEdit
'        End If
'    Else
'        noEjecutarAccion = True
'        Call SeleccionarRespuestaAccionDesarrollo(Row)
'        noEjecutarAccion = False
'    End If
'    Call formatoFilaPlanIntegral(Row)
End Sub

Private Sub grdPlanDesarrollo_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
'    AdministrarKeyPreview KeyCode
End Sub

Private Sub grdPlanDesarrolloPendientes_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Set ssCellActivate = Cell
    'mgaray201411b
    If Cell.Column.Key = "SiEjecutaAccion" Or Cell.Column.Key = "NoEjecutaAccion" Then
        mb_EstaMarcadoEjecucionPsicomotor = Cell.Value
    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not (ssCellActivate Is Nothing) Then
        If ssCellActivate.Column.Key = "SiEjecutaAccion" Or ssCellActivate.Column.Key = "NoEjecutaAccion" Then
            EventsSeleccinarEjecutaAccion ssCellActivate, ssCellActivate.Row.Cells(ssCellActivate.Column.Key).Value
        End If
    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
'    If Not IsNull(NewValue) And Cell.Column.Key = "FechaEjecucion" Then
'        If NewValue < ml_FechaAtencion Or NewValue > md_fechaActual Then
'            MsgBox "Fecha no puede ser menor que la fecha de atención ni mayor que la fecha actual", vbInformation, "Advertencia"
'            NewValue = Cell.Value
'        End If
'    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPlanDesarrolloPendientes_CellChange(ByVal Cell As UltraGrid.SSCell)
    If noEjecutarAccion = True Then: Exit Sub
    If Cell.Column.Key = "SiEjecutaAccion" Or Cell.Column.Key = "NoEjecutaAccion" Then
        'mgaray201411b
        mb_EstaMarcadoEjecucionPsicomotor = Not mb_EstaMarcadoEjecucionPsicomotor
        If mb_EstaMarcadoEjecucionPsicomotor = True Then
'        If Cell.Value = True Then
            noEjecutarAccion = True
            If Cell.Column.Key = "SiEjecutaAccion" Then
                Cell.Row.Cells("NoEjecutaAccion").Value = False
            Else
                Cell.Row.Cells("SiEjecutaAccion").Value = False
            End If
            noEjecutarAccion = False
        End If
        Call HabilitarGuardar
    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanDesarrolloPendientes.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanDesarrolloPendientes.UpdateMode = ssUpdateOnUpdate
    
    'detalle del grupo
    With Layout.Bands(0)
        .Columns.Add "SiEjecutaAccion", "SI"
        .Columns.Add "NoEjecutaAccion", "NO"
        
        .Columns("SiEjecutaAccion").DataType = ssDataTypeBoolean
        .Columns("SiEjecutaAccion").Style = ssStyleCheckBox
        
        .Columns("NoEjecutaAccion").DataType = ssDataTypeBoolean
        .Columns("NoEjecutaAccion").Style = ssStyleCheckBox
        
        'establecer etiqueta de columnas Y formato
        
        .Columns("SiEjecutaAccion").Width = 1000
        .Columns("NoEjecutaAccion").Width = 1000
                
        .Columns("ItemDesarrollo").Header.Caption = "Descripción de Item a Evaluar"
        .Columns("ItemDesarrollo").Width = grdPlanDesarrolloPendientes.Width - 500 - .Columns("SiEjecutaAccion").Width - _
                                                    .Columns("NoEjecutaAccion").Width
        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 0, ssAlignCenter, "SiEjecutaAccion", "NoEjecutaAccion")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 0, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "IdItemDesarrollo", "OrdenItem", "EjecutaAccion", _
                                        "EdadAnio", "EdadMes", "EdadDia", "FechaEjecucion")
        
        '.Columns("EsEjecutada").Activation = ssActivationAllowEdit
        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 0, ssActivationActivateNoEdit, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "IdItemDesarrollo", "ItemDesarrollo", "OrdenItem", "EjecutaAccion", "EdadAnio", "EdadMes", "EdadDia", "FechaEjecucion")
    End With
End Sub

Private Sub grdPlanDesarrolloPendientes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If noEjecutarAccion = True Then Exit Sub
    noEjecutarAccion = True
    Call SeleccionarRespuestaAccionDesarrollo(Row)
    noEjecutarAccion = False
'    Call formatoFilaPlanIntegral(Row)
    
    Row.Cells("SiEjecutaAccion").Activation = ssActivationAllowEdit
    Row.Cells("NoEjecutaAccion").Activation = ssActivationAllowEdit
End Sub

Private Sub grdPlanDesarrolloPendientes_LostFocus()
    Dim ssReturnValue As SSReturnBoolean
    
    grdPlanDesarrolloPendientes_BeforeCellDeactivate ssReturnValue
    Me.txtEvalucionDesarrollo.Tag = ObtenerEvaluacion()
    Me.txtEvalucionDesarrollo.Text = ObtenerEvaluacionDescripcion(Me.txtEvalucionDesarrollo.Tag)
End Sub

Private Sub mskFechaEjecucionDes_Change()
    Call HabilitarGuardar
End Sub

Private Sub mskFechaEjecucionDes_GotFocus()
    mskFechaEjecucionDes.Tag = mskFechaEjecucionDes.Text
End Sub

Private Sub mskFechaEjecucionDes_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub mskFechaEjecucionDes_LostFocus()
    If mskFechaEjecucionDes.Text <> sighEntidades.FECHA_VACIA_DMY Then
        On Error Resume Next
        If Not EsFecha(mskFechaEjecucionDes.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de Desarrollo"
            mskFechaEjecucionDes.Text = mskFechaEjecucionDes.Tag
            mskFechaEjecucionDes.SetFocus
        ElseIf CDate(mskFechaEjecucionDes.Text) < CDate(txtFechaProgramadaDesarrollo.Text) _
                        Or CDate(mskFechaEjecucionDes.Text) > md_fechaActual Then
            MsgBox "Fecha no puede ser menor que la fecha programada ni mayor que la fecha actual", vbInformation, "Datos de Desarrollo"
            mskFechaEjecucionDes.Text = mskFechaEjecucionDes.Tag
            mskFechaEjecucionDes.SetFocus
        End If
        
    End If
   'mo_Formulario.MarcarComoVacio txtFechaNacimiento
End Sub

Private Function getFechaActual() As Date
'    If md_fechaActual = 0 Then
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        md_fechaActual = lcBuscaParametro.RetornaFechaServidorSQL
'    End If
    getFechaActual = md_fechaActual
End Function

Public Function cargarDatosAtencionIntegralDesarrollo() As Boolean
    Call getFechaActual
    Call cargarListaDesarrolloVencidos
'    Call cargarListaDesarrolloPendientes
    cargarDatosAtencionIntegralDesarrollo = True
End Function

Public Sub cargarListaDesarrolloVencidos()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
    oDOAtenIntePlanIntePaciente.idAtencion = ml_idAtencion
    
    Set grdPlanDesarrollo.DataSource = oReglasAtencionIntegral.ListarPlanDesarrolloPacienteVencidos(oDOAtenIntePlanIntePaciente)
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanDesarrollo, sighEntidades.GrillaConFilasBicolor
End Sub

Public Sub cargarListaPlanDesarrolloPacienteDetalle(lIdPlanIntegralPaciente As Long, _
            lIdPlanDesarrolloPaciente As Long)
            
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanDesPacienteDet As New DOAtenIntePlanDesPacienteDet
    'mgaray20141022
    oDOAtenIntePlanDesPacienteDet.IdPlanDesarrolloPaciente = lIdPlanDesarrolloPaciente
    oDOAtenIntePlanDesPacienteDet.IdPlanIntegralPaciente = lIdPlanIntegralPaciente
        
    Set grdPlanDesarrolloPendientes.DataSource = oReglasAtencionIntegral.ListarPlanDesarrolloPacienteDetallePorId(oDOAtenIntePlanDesPacienteDet)
        
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
'    Else
'        Call AsignarDatosAControlesDesarrollo(oDOAtenIntePlanIntePaciente)
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanDesarrolloPendientes, sighEntidades.GrillaConFilasBicolor
End Sub


Private Function LimpiarDatosAControlesDesarrollo() As Boolean
    Dim oFechaHOra As New FechaHora
    frAtencionDesarrollo.Caption = ""
    txtFechaProgramadaDesarrollo.Text = ""
    txtIdAtencionDesarrollo.Text = ""
    mskFechaEjecucionDes.Text = oFechaHOra.FECHA_VACIA_DMY
    txtEvalucionDesarrollo.Text = ""
    txtEvalucionDesarrollo.Tag = ""
    
    cmbIdEstablecimiento.Text = ""
    
    
    btnAceptar.Enabled = False
End Function

Private Function AsignarDatosAControlesDesarrollo(lIdPlanIntegralPaciente As Long, _
            lIdPlanDesarrolloPaciente As Long) As Boolean
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oFechaHOra As New FechaHora
    
    frAtencionDesarrollo.Enabled = True
    
    Set oDOAtenIntePlanDesPaciente = oReglasAtencionIntegral.PlanDesarrolloPacienteSeleccionarPorId(lIdPlanIntegralPaciente, _
                            lIdPlanDesarrolloPaciente)

    If Not (oDOAtenIntePlanDesPaciente Is Nothing) Then
        frAtencionDesarrollo.Caption = "Sesión " & oDOAtenIntePlanDesPaciente.NumeroSesion
        txtFechaProgramadaDesarrollo.Text = IIf(oDOAtenIntePlanDesPaciente.FechaProgramada = 0, oFechaHOra.FECHA_VACIA_DMY, oDOAtenIntePlanDesPaciente.FechaProgramada)
        txtIdAtencionDesarrollo.Text = IIf(oDOAtenIntePlanDesPaciente.idAtencion = 0, "", oDOAtenIntePlanDesPaciente.idAtencion)
        mskFechaEjecucionDes.Text = IIf(oDOAtenIntePlanDesPaciente.FechaEjecucion = 0, md_fechaActual, oDOAtenIntePlanDesPaciente.FechaEjecucion)
        txtEvalucionDesarrollo.Text = oDOAtenIntePlanDesPaciente.EvaluacionDesc
        
        txtEvalucionDesarrollo.Tag = IIf(oDOAtenIntePlanDesPaciente.evaluacion = 0, "", oDOAtenIntePlanDesPaciente.evaluacion)
        
        
        If oDOAtenIntePlanDesPaciente.IdEstablecimiento > 0 Then
            cmbEspecialidad_UbicaPosicion (oDOAtenIntePlanDesPaciente.IdEstablecimiento)
        Else
            cmbIdEstablecimiento.Text = ""
        End If
    Else
        Call LimpiarDatosAControlesDesarrollo
    End If
    Call cargarListaPlanDesarrolloPacienteDetalle(lIdPlanIntegralPaciente, lIdPlanDesarrolloPaciente)
    btnAceptar.Enabled = False
End Function

Private Function ObtenerEvaluacionDescripcion(idEvaluacion As Integer) As String
    Select Case idEvaluacion
        Case 1:
            ObtenerEvaluacionDescripcion = "NORMAL"
        Case 2:
            ObtenerEvaluacionDescripcion = "DEFICIT"
        Case Else:
            ObtenerEvaluacionDescripcion = ""
    End Select
End Function

Public Function ObtenerEvaluacion() As Integer
    Dim oRs As ADODB.Recordset
    Dim totalItems As Integer
    Dim itemNoEjecutados As Integer, itemEjecutados As Integer
    Dim evaluacion As Integer
    
    evaluacion = 0
    itemEjecutados = 0
    itemNoEjecutados = 0
    
    If validarEvaluacionDesarrollo() = True Then
        Set oRs = grdPlanDesarrolloPendientes.DataSource
        If Not (oRs Is Nothing) Then
            totalItems = oRs.RecordCount
            If totalItems > 0 Then
                Dim siguienteFila As Boolean
                Dim oSRow As SSRow
                
                'leer las filas debido a que se necesita acceder a dos columnas agregadas que no estan en el recorset
                Set oSRow = grdPlanDesarrolloPendientes.GetRow(ssChildRowFirst)
                siguienteFila = True
                If Not (oSRow Is Nothing) Then
                    While siguienteFila = True
                        If oSRow.Cells("SiEjecutaAccion").Value = True Or oSRow.Cells("NoEjecutaAccion").Value = True Then
                            If oSRow.Cells("SiEjecutaAccion").Value = True Then
                                itemEjecutados = itemEjecutados + 1
                            Else
                                itemNoEjecutados = itemNoEjecutados + 1
                            End If
                        End If
                        siguienteFila = oSRow.HasNextSibling
                        If siguienteFila = True Then
                            Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                        End If
                    Wend
                End If
            End If
        End If
        If itemNoEjecutados = 0 Then
            evaluacion = 1
        Else
            evaluacion = 2
        End If
    End If
    ObtenerEvaluacion = evaluacion
End Function

Public Function validarEvaluacionDesarrollo() As Boolean
    validarEvaluacionDesarrollo = False
    Dim oRs As ADODB.Recordset
    Dim totalItems As Integer
    Dim itemNoEjecutados As Integer, itemEjecutados As Integer
    Err = 0
    itemEjecutados = 0
    itemNoEjecutados = 0
    totalItems = 0
    
    Set oRs = grdPlanDesarrolloPendientes.DataSource
    
    If Not (oRs Is Nothing) Then
        totalItems = oRs.RecordCount
        
        If totalItems > 0 Then
            Dim siguienteFila As Boolean
            Dim oSRow As SSRow
            
            'leer las filas debido a que se necesita acceder a dos columnas agregadas que no estan en el recorset
            Set oSRow = grdPlanDesarrolloPendientes.GetRow(ssChildRowFirst)
            siguienteFila = True
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.Cells("SiEjecutaAccion").Value = True Or oSRow.Cells("NoEjecutaAccion").Value = True Then
                        If oSRow.Cells("SiEjecutaAccion").Value = True Then
                            itemEjecutados = itemEjecutados + 1
                        Else
                            itemNoEjecutados = itemNoEjecutados + 1
                        End If
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    If totalItems = itemEjecutados + itemNoEjecutados Then
        validarEvaluacionDesarrollo = True
    End If
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Advertencia"
    End If
End Function

Private Function initializeControls() As Boolean
    frAtencionDesarrollo.Enabled = False
    btnAceptar.Enabled = False
    initializeControls = True
End Function

'hacer refactory
Public Function SeleccionarRespuestaAccionDesarrollo(ByVal Row As UltraGrid.SSRow)
    If IsNull(Row.Cells("EjecutaAccion").Value) Then
        Row.Cells("SiEjecutaAccion").Value = False
        Row.Cells("NoEjecutaAccion").Value = False
        
    Else
        If Row.Cells("EjecutaAccion").Value = True Then
            Row.Cells("SiEjecutaAccion").Value = True
            Row.Cells("NoEjecutaAccion").Value = False
        Else
            Row.Cells("SiEjecutaAccion").Value = False
        Row.Cells("NoEjecutaAccion").Value = True
        End If
    End If
End Function

Private Function EventsSeleccinarEjecutaAccion(Cell As SSCell, EjecutaAccion As Boolean)
    If EjecutaAccion = True Then
        noEjecutarAccion = True
        If Cell.Column.Key = "SiEjecutaAccion" Then
            Cell.Row.Cells("NoEjecutaAccion").Value = False
        Else
            Cell.Row.Cells("SiEjecutaAccion").Value = False
        End If
        noEjecutarAccion = False
    End If
    
End Function

Public Sub ConfigurarCombos()
       
    Set cmbIdEstablecimiento.ListSource = mo_ReglasComunes.EstablecimientosSeleccionarTodos()
End Sub

Private Function validarCambiosPendientes() As Boolean
    validarCambiosPendientes = False
    If btnAceptar.Enabled = True Then
        Dim i As Integer
        If MsgBox("No ha guardado los cambios realizados, ¿Desea descartalos?", vbQuestion + vbYesNo, "Advertencia") = vbNo Then
            Exit Function
        End If
    End If
    validarCambiosPendientes = True
End Function

Private Sub txtEvalucionDesarrollo_Change()
    Call HabilitarGuardar
End Sub

Private Function HabilitarGuardar()
    btnAceptar.Enabled = True
End Function

Sub cmbEspecialidad_UbicaPosicion(lnIdEstablecimiento As Long)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbIdEstablecimiento.ListCount - 1)
        cmbIdEstablecimiento.ListIndex = lnFor
        If cmbIdEstablecimiento.SubItem(cmbIdEstablecimiento.ListIndex, 0) = Val(lnIdEstablecimiento) Then
           Exit For
        End If
    Next
End Sub

Private Function ValidarDatosIngreso() As Boolean
    Dim ms_MensajeError As String
    ValidarDatosIngreso = False
    
    ms_MensajeError = ""
    
    If mskFechaEjecucionDes.Text = sighEntidades.FECHA_VACIA_DMY Then
        ms_MensajeError = ms_MensajeError & "Ingrese fecha de Ejecución" & Chr(13)
        mskFechaEjecucionDes.SetFocus
    End If
    If cmbIdEstablecimiento.ListIndex < 0 Then
       cmbIdEstablecimiento.Text = ""
    End If
    If cmbIdEstablecimiento.Text = "" Then
        ms_MensajeError = "Elija Establecimieniento" & Chr(13)
        cmbIdEstablecimiento.SetFocus
    End If
        
    If validarEvaluacionDesarrollo() = False Then
        ms_MensajeError = ms_MensajeError & "Debe de Evaluar Todos los Item de Desarrollo"
    End If
    
    If ms_MensajeError = "" Then
        ValidarDatosIngreso = True
    Else
        MsgBox ms_MensajeError, vbInformation, "Faltan Datos"
    End If
End Function

Private Function getRecorsetDesarrollo()
    Dim oRsDesarrollo As New Recordset
    Dim oRsDesarrolloGrida As ADODB.Recordset
    Dim cambioEjecucion As Boolean
    
    Set oRsDesarrolloGrida = grdPlanDesarrolloPendientes.DataSource
    'Datos de la session de desarrollo
    cambioEjecucion = False
    
    
    If Not (oRsDesarrolloGrida Is Nothing) Then
        'crear un recorset temporal basado en el recorset de la grida
        Set oRsDesarrollo = clonarRecorset(oRsDesarrolloGrida)
'        oRsDesarrollo.Fields.Append "SiEjecutaAccion", adBoolean, 1, adFldIsNullable
'        oRsDesarrollo.Fields.Append "NoEjecutaAccion", adBoolean, 1, adFldIsNullable
        oRsDesarrollo.Open
        
        If oRsDesarrolloGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim oSRow As SSRow
            
            Set oSRow = grdPlanDesarrolloPendientes.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.DataChanged = True Then
                        'trasladar los valores de la grida al recorset temporral
                        agregarFilaAtencionIntegral oRsDesarrollo, oSRow
                        
                        If oSRow.Cells("SiEjecutaAccion").Value = True Or _
                                        oSRow.Cells("NoEjecutaAccion").Value = True Then
                            If oSRow.Cells("SiEjecutaAccion").Value = True Then
                                oRsDesarrollo!EjecutaAccion = 1
                            Else
                                oRsDesarrollo!EjecutaAccion = 0
                            End If
                        Else
                            oRsDesarrollo!EjecutaAccion = Null
                        End If
                        oRsDesarrollo.Update
                        cambioEjecucion = True
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    
'    Call SetDatosEjecucionDesarrollo(cambioEjecucion, oRsDesarrollo)
    
    Set getRecorsetDesarrollo = oRsDesarrollo
End Function

Private Function clonarRecorset(oRsOriginal As ADODB.Recordset, Optional fieldExclude As String = "") As ADODB.Recordset
    Dim oRsLocal As New Recordset
    If Not (oRsOriginal Is Nothing) Then
        'If oRsOriginal.RecordCount > 0 Then
            Dim i As Integer
            With oRsLocal
                For i = 0 To oRsOriginal.Fields.Count - 1
                    If fieldExclude = "" Or oRsOriginal.Fields(i).Name <> fieldExclude Then
                        .Fields.Append oRsOriginal.Fields(i).Name, _
                                        oRsOriginal.Fields(i).Type, _
                                        oRsOriginal.Fields(i).DefinedSize, _
                                        adFldIsNullable
                    End If
                Next i
                '.Fields.Append "EsEjecutada", adBoolean, 1, adFldIsNullable
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
            End With
        'End If
    End If
    Set clonarRecorset = oRsLocal
End Function

Private Function agregarFilaAtencionIntegral(ByRef oRs As ADODB.Recordset, oSRow As SSRow) As Boolean
On Error GoTo miError
    Dim i As Integer
    oRs.AddNew
    For i = 0 To oRs.Fields.Count - 1
        oRs.Fields(oSRow.Cells(i).Column.Key).Value = oSRow.Cells(i).Value
    Next i
    agregarFilaAtencionIntegral = True
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Perinatal - Inmunizaciones"
    End If
End Function

Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
    End Select
       
End Sub

Private Sub txtEvalucionDesarrollo_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaProgramadaDesarrollo_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdBuscarEstablecimiento_Click()
    Dim oForm As New SIGHNegocios.BuscaEstablecimientos
    Dim oDoEstablecimiento As New DOEstablecimiento
    Dim mo_RcsListaEstablecimientos  As New Recordset
    
    oForm.DescripcionEstablecimiento = "" 'Me.txtNombreEstablecimiento.Text
    oForm.NivelMaximoEstablecimiento = 0
    oForm.MostrarFormulario
    
    Me.btnAceptar.Enabled = True
    
    If oForm.idRegistroSeleccionado = 0 Then
        Call MsgBox("No ha seleccionado ningún registro de la Lista.", vbExclamation, Me.Caption)
    Else
        cmbEspecialidad_UbicaPosicion (oForm.idRegistroSeleccionado)
'        'Ingresando los valores del Establecimiento Elegido
'        If oForm.BotonPresionado = sghAceptar Then
'            Set oDoEstablecimiento = mr_ReglasComunes.EstablecimientosSeleccionarPorId(oForm.idRegistroSeleccionado)
'            If Not oDoEstablecimiento Is Nothing Then
'                Set mo_Establecimiento = oDoEstablecimiento
'                ml_IdEstablecimiento = oDoEstablecimiento.IdEstablecimiento
'                Me.txtNombreEstablecimiento.Text = mo_Establecimiento.Nombre
'                Me.txtCodigoEstablecimiento.Text = mo_Establecimiento.Codigo
'            End If
'        End If
    End If
End Sub
