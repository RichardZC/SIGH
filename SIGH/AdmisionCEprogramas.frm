VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form AdmisionCEprogramas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultorios que registran información al mismo tiempo"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAceptar 
      Height          =   1035
      Left            =   45
      TabIndex        =   3
      Top             =   6960
      Width           =   11775
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AdmisionCEprogramas.frx":0000
         DownPicture     =   "AdmisionCEprogramas.frx":0460
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
         Left            =   4470
         Picture         =   "AdmisionCEprogramas.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AdmisionCEprogramas.frx":0D4A
         DownPicture     =   "AdmisionCEprogramas.frx":120E
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
         Left            =   6037
         Picture         =   "AdmisionCEprogramas.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdBusqueda 
      Height          =   6975
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   12303
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
      Caption         =   "Lista de UPS"
   End
End
Attribute VB_Name = "AdmisionCEprogramas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Lista atenciones de un Paciente
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idCuentaAtencion  As Long
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_ReglasAdmision  As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_Facturacion As New SIGHNegocios.ReglasFacturacion
'Dim oRsFuas As New Recordset
Dim ml_EstadoCuenta As String
Dim ml_idCorrelativo As Long
Dim ml_formLlamante As String       'boton agregar cpt desde atencion Medico, boton FUA desde atencion del MEdico
Dim ml_NroFua As Long
Dim ml_FuaIdCuentaAtencion As Long
Dim mrs_oRsFua As Recordset
Dim oRsItemsMasivosElegidos As New Recordset
Dim oRsTipoDx As New Recordset
Dim oRsActividades As New Recordset
Dim lbPulsoBotonAceptar As Boolean
Dim ml_ups As String
Dim lcSql As String

Property Let UPS(lValue As String)
    ml_ups = lValue
End Property

Property Set oRsItemsElegidos(oValue As Recordset)
    Set oRsTipoDx = oValue
End Property

Property Set oRsFua(oValue As Recordset)
    Set mrs_oRsFua = oValue
End Property
Property Get NroFua() As Long
   NroFua = ml_NroFua
End Property
Property Get FuaIdCuentaAtencion() As Long
   FuaIdCuentaAtencion = ml_FuaIdCuentaAtencion
End Property
Property Let FormLlamante(lValue As String)
    ml_formLlamante = lValue
End Property
Property Let Correlativo(lValue As Long)
    ml_idCorrelativo = lValue
End Property
Property Let EstadoCuenta(lValue As String)
    ml_EstadoCuenta = lValue
End Property
Property Get EstadoCuenta() As String
    EstadoCuenta = ml_EstadoCuenta
End Property

Property Set Atenciones(oValue As Recordset)
    Set grdBusqueda.DataSource = oValue
End Property

Property Get Atenciones() As Recordset
End Property

Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property

Property Get idCuentaAtencion() As Long
    idCuentaAtencion = ml_idCuentaAtencion
End Property

Private Sub btnAceptar_Click()
    If ValidarReglas Then
        'grdBusqueda_DblClick
        ml_idCuentaAtencion = 1
        lbPulsoBotonAceptar = True
        Me.Visible = False
    End If
End Sub

Private Sub btnCancelar_Click()
    ml_idCuentaAtencion = 0
    Select Case ml_formLlamante
    Case "ACTIVIDADES"
         lbPulsoBotonAceptar = False
    End Select
    Me.Visible = False
End Sub














Sub LimpiaDatos()
End Sub









Private Sub Form_Load()
    ml_idCuentaAtencion = 0
    ml_NroFua = 0
    Me.Width = 11955
    Select Case ml_formLlamante
    Case "ACTIVIDADES"
         If oRsItemsMasivosElegidos.State = 1 Then Set oRsItemsMasivosElegidos = Nothing
         With oRsItemsMasivosElegidos
              .Fields.Append "Grupo", adInteger
              .Fields.Append "SubGrupo", adInteger
              .Fields.Append "lab", adVarChar, 3, adFldIsNullable
              .Fields.Append "Tipo", adVarChar, 20, adFldIsNullable
              .Fields.Append "id", adVarChar, 20, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable
              .Fields.Append "Elija", adBoolean
              .Fields.Append "ElijaTipo", adInteger
              .Fields.Append "ElijaUPS", adInteger
              .Fields.Append "ElijaLab", adVarChar, 3, adFldIsNullable
              .Fields.Append "IdCuentaAtencion", adInteger, 4
              .Fields.Append "IdOrden", adInteger, 4
              .Fields.Append "Fua", adInteger
              .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
              .Fields.Append "IdServicio", adInteger
              .Fields.Append "FuaCodigoPrestacion", adVarChar, 3, adFldIsNullable
              .Fields.Append "idTipo", adInteger
              .Fields.Append "idServicioPaciente", adInteger
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open
         End With
    
         
         mrs_oRsFua.MoveFirst
         Set grdBusqueda.DataSource = mrs_oRsFua
         Dim lnFor As Integer
         With grdBusqueda.ValueLists.Add("ElijaTipoList").ValueListItems
             oRsTipoDx.MoveFirst
             Do While Not oRsTipoDx.EOF
                lnFor = oRsTipoDx!IdSubclasificacionDx - 100
                .Add lnFor, oRsTipoDx!DescripcionLarga  'oRsTipoDx!IdSubclasificacionDx, oRsTipoDx!DescripcionLarga
                oRsTipoDx.MoveNext
             Loop
         End With
         grdBusqueda.Bands(0).Columns("ElijaTipo").ValueList = "ElijaTipoList"
         grdBusqueda.Bands(0).Columns("ElijaTipo").ButtonDisplayStyle = ssButtonDisplayStyleAlways
         '
         
         With grdBusqueda.ValueLists.Add("TipoDxCptLab").ValueListItems
             .Add 1, LxCPT
             .Add 2, Lx_Lab
             .Add 3, LxDx
         End With
         grdBusqueda.Bands(0).Columns("idTipo").ValueList = "TipoDxCptLab"
         grdBusqueda.Bands(0).Columns("idTipo").ButtonDisplayStyle = ssButtonDisplayStyleAlways
         '
        Dim oRsLab As New Recordset
        Set oRsLab = mo_AdminServiciosComunes.DevuelveHIS_SITUACIOporDescripcion()
        With grdBusqueda.ValueLists.Add("LabLista").ValueListItems
           oRsLab.MoveFirst
           Do While Not oRsLab.EOF
              .Add Trim(oRsLab!valores), oRsLab!valores
              oRsLab.MoveNext
           Loop
        End With
        oRsLab.Close
        Set oRsLab = Nothing
        grdBusqueda.Bands(0).Columns("ElijaLab").ValueList = "LabLista"
        grdBusqueda.Bands(0).Columns("ElijaLab").ButtonDisplayStyle = ssButtonDisplayStyleAlways
         
         '
         grdBusqueda.Bands(0).Columns("ElijaUPS").Hidden = True
         grdBusqueda.Bands(0).Columns("IdCuentaAtencion").Hidden = True
         grdBusqueda.Bands(0).Columns("IdOrden").Hidden = True
         grdBusqueda.Bands(0).Columns("Fua").Hidden = True
         grdBusqueda.Bands(0).Columns("Consultorio").Hidden = True
         grdBusqueda.Bands(0).Columns("IdServicio").Hidden = True
         grdBusqueda.Bands(0).Columns("FuaCodigoPrestacion").Hidden = True
         grdBusqueda.Bands(0).Columns("Tipo").Hidden = True
         grdBusqueda.Bands(0).Columns("lab").Hidden = True
         grdBusqueda.Bands(0).Columns("idServicioPaciente").Hidden = True
         grdBusqueda.Bands(0).Columns("Elija").Header.Appearance.BackColor = vbRed
         grdBusqueda.Bands(0).Columns("Elija").Header.Appearance.Font.Bold = True
         grdBusqueda.Bands(0).Columns("ElijaTipo").Header.Appearance.BackColor = vbRed
         grdBusqueda.Bands(0).Columns("ElijaTipo").Header.Appearance.Font.Bold = True
         grdBusqueda.Bands(0).Columns("ElijaLab").Header.Appearance.BackColor = vbRed
         grdBusqueda.Bands(0).Columns("ElijaLab").Header.Appearance.Font.Bold = True
         grdBusqueda.Bands(0).Columns("Grupo").Activation = ssActivationActivateNoEdit
         grdBusqueda.Bands(0).Columns("SubGrupo").Activation = ssActivationActivateNoEdit
         grdBusqueda.Bands(0).Columns("idTipo").Header.Appearance.BackColor = vbRed
         grdBusqueda.Bands(0).Columns("id").Header.Appearance.BackColor = vbRed
         grdBusqueda.Bands(0).Columns("nombre").Header.Appearance.BackColor = vbRed
         '
         grdBusqueda.Caption = "Lista de ACTIVIDADES a llenar"
         Me.Caption = "ACTIVIDADES"

    End Select
End Sub




Private Sub grdBusqueda_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    Select Case ml_formLlamante
    Case "ACTIVIDADES"
        
        Select Case Cell.Column.Key
        Case "Elija"
           If Cell.Row.Cells("Elija").Value = True Then
              If Trim(Cell.Row.Cells("Lab").Value) <> "" Then
                 Cell.Row.Cells("ElijaLab").Value = Cell.Row.Cells("Lab").Value
              End If
           Else
              Cell.Row.Cells("ElijaLab").Value = ""
           End If
        Case "idTipo"
           Select Case Cell.Row.Cells("idTipo").Value
           Case "2"
              Cell.Row.Cells("id").Value = sighEntidades.Lx_LabVacio
              Cell.Row.Cells("nombre").Value = sighEntidades.Lx_LabVacio
              Cell.Row.Cells("elija").Value = True
           Case "1", "3"
              Cell.Row.Cells("elija").Value = False
           End Select
        Case "id"
           If Cell.Row.Cells("id").Value <> "" Then
              Cell.Row.Cells("id").Value = Trim(Cell.Row.Cells("id").Value)
              Dim oRsTmp2 As New Recordset
              Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
              Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
              Dim lnIdTipoAct As Integer
              lnIdTipoAct = Val(Cell.Row.Cells("idTipo").Value)
              Select Case lnIdTipoAct
              Case sghActividadesTipo.TipoCPT
                Set oRsTmp2 = mo_AdminCaja.FactCatalogoServiciosSeleccionarPorCodigoOnombre(Cell.Row.Cells("id").Value, "")
                If oRsTmp2.RecordCount > 0 Then
                   Cell.Row.Cells("Nombre").Value = Left(oRsTmp2!nombre, 255)
                   Cell.Row.Cells("elija").Value = True
                End If
                oRsTmp2.Close
              Case sghActividadesTipo.TipoDX
                Set oRsTmp2 = mo_AdminServiciosComunes.DiagnosticosSeleccionarXCodigo(Cell.Row.Cells("id").Value)
                If oRsTmp2.RecordCount > 0 Then
                   Cell.Row.Cells("Nombre").Value = Left(oRsTmp2!descripcion, 255)
                   Cell.Row.Cells("elija").Value = True
                End If
                 oRsTmp2.Close
              End Select
              Set oRsTmp2 = Nothing
              Set mo_AdminCaja = Nothing
           End If
        End Select
    End Select
End Sub

Private Sub grdBusqueda_ClickCellButton(ByVal Cell As UltraGrid.SSCell)
            Dim lnIdTipoAct As Integer
            lnIdTipoAct = Val(grdBusqueda.ActiveRow.Cells("idTipo").Value)
            Select Case lnIdTipoAct
            Case sghActividadesTipo.TipoCPT
                Dim oFrm As New SIGHNegocios.BuscaServicio
                Dim dOServ As New DOCatalogoServicio
                oFrm.MostrarFormulario
                If oFrm.idRegistroSeleccionado <> 0 Then
                    Set dOServ = mo_Facturacion.CatalogoServiciosSeleccionarPorId(oFrm.idRegistroSeleccionado)
                    If Not dOServ Is Nothing Then
                        grdBusqueda.ActiveRow.Cells("id").Value = dOServ.Codigo
                        grdBusqueda.ActiveRow.Cells("nombre").Value = dOServ.nombre
                        grdBusqueda.ActiveRow.Cells("elija").Value = True
                    End If
                End If
                Set oFrm = Nothing
                Set dOServ = Nothing
            Case sghActividadesTipo.TipoLAB
                grdBusqueda.ActiveRow.Cells("id").Value = sighEntidades.Lx_LabVacio
                grdBusqueda.ActiveRow.Cells("nombre").Value = sighEntidades.Lx_LabVacio
                grdBusqueda.ActiveRow.Cells("elija").Value = True
            Case sghActividadesTipo.TipoDX
                Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
                Dim oDODiagnostico As DODiagnostico
                oBusqueda.SoloMuestraDxGalenHos = False
                oBusqueda.MostrarFormulario
                If oBusqueda.BotonPresionado = sghAceptar Then
                    Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
                    If Not oDODiagnostico Is Nothing Then
                        grdBusqueda.ActiveRow.Cells("id").Value = oDODiagnostico.CodigoCIE2004
                        grdBusqueda.ActiveRow.Cells("nombre").Value = oDODiagnostico.descripcion
                        grdBusqueda.ActiveRow.Cells("elija").Value = True
                    End If
                End If
                Set oBusqueda = Nothing
            End Select
End Sub

Private Sub grdBusqueda_DblClick()
'    Dim rsRecordset As Recordset
'    Set rsRecordset = grdBusqueda.DataSource
'    Select Case ml_formLlamante
'    Case "ACTIVIDADES"
'         ml_idCuentaAtencion = 1
'    End Select
    'Me.Visible = False
End Sub

Private Sub grdBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Select Case ml_formLlamante
    Case "ACTIVIDADES"
        grdBusqueda.Bands(0).Columns("Grupo").Hidden = True
        grdBusqueda.Bands(0).Columns("subGrupo").Width = 700
        grdBusqueda.Bands(0).Columns("grupoTIT").Width = 500
        grdBusqueda.Bands(0).Columns("grupoTIT").Header.Caption = "Grupo"
        grdBusqueda.Bands(0).Columns("lab").Width = 500
        grdBusqueda.Bands(0).Columns("tipo").Width = 500
        grdBusqueda.Bands(0).Columns("id").Width = 800
        grdBusqueda.Bands(0).Columns("idTipo").Header.Caption = "Tipo"
        grdBusqueda.Bands(0).Columns("id").Header.Caption = "Código"
        grdBusqueda.Bands(0).Columns("nombre").Width = 5000
        grdBusqueda.Bands(0).Columns("nombre").Style = ssStyleButton
        grdBusqueda.Bands(0).Columns("idTipo").Width = 1000
        grdBusqueda.Bands(0).Columns("ElijaLab").Width = 1000
        
       
        
    End Select
    mo_Apariencia.ConfigurarFilasBiColores grdBusqueda, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub grdBusqueda_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
   If KeyAscii = 13 Then
      grdBusqueda_DblClick
   End If
End Sub

Property Get ItemsMasivosElegidos() As Recordset
    If lbPulsoBotonAceptar = True Then
        Dim oRow As SSRow
        grdBusqueda.Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
        grdBusqueda.Bands(0).Columns("Elija").SortIndicator = ssSortIndicatorAscending
        Set oRow = Me.grdBusqueda.GetRow(ssChildRowFirst)
        If Not oRow Is Nothing Then
                Do While oRow.HasNextSibling
                    Set oRow = oRow.GetSibling(ssSiblingRowNext)
                    If oRow.Cells("Elija").Value = True Then
                        oRsItemsMasivosElegidos.AddNew
                        oRsItemsMasivosElegidos!Grupo = oRow.Cells("Grupo").Value
                        oRsItemsMasivosElegidos!SubGrupo = oRow.Cells("SubGrupo").Value
                        oRsItemsMasivosElegidos!lab = oRow.Cells("lab").Value
                        oRsItemsMasivosElegidos!ID = oRow.Cells("id").Value
                        oRsItemsMasivosElegidos!tipo = oRow.Cells("tipo").Value
                        oRsItemsMasivosElegidos!nombre = oRow.Cells("nombre").Value
                        oRsItemsMasivosElegidos!elija = oRow.Cells("elija").Value
                        oRsItemsMasivosElegidos!elijaTipo = oRow.Cells("elijaTipo").Value + 100
                        oRsItemsMasivosElegidos!ElijaUPS = oRow.Cells("ElijaUPS").Value
                        oRsItemsMasivosElegidos!ElijaLab = oRow.Cells("ElijaLab").Value
                        
                        oRsItemsMasivosElegidos!IdOrden = oRow.Cells("idOrden").Value
                        oRsItemsMasivosElegidos!FUA = oRow.Cells("Fua").Value
                        oRsItemsMasivosElegidos!Consultorio = oRow.Cells("Consultorio").Value
                        oRsItemsMasivosElegidos!IdServicio = oRow.Cells("idServicio").Value
                        oRsItemsMasivosElegidos!FuaCodigoPrestacion = oRow.Cells("FuaCodigoPrestacion").Value
                        oRsItemsMasivosElegidos!idTipo = oRow.Cells("idTipo").Value
                        oRsItemsMasivosElegidos!idCuentaAtencion = oRow.Cells("idCuentaAtencion").Value
                        oRsItemsMasivosElegidos!idServicioPaciente = oRow.Cells("idServicioPaciente").Value
                    End If
                Loop
    
        End If
    
        Set oRow = Nothing
       
        Set ItemsMasivosElegidos = oRsItemsMasivosElegidos.Clone
    End If
End Sub



Function ValidarReglas() As Boolean
    ValidarReglas = False
    Dim oRow As SSRow
    Dim lcMensaje As String
'    grdBusqueda.Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
'    grdBusqueda.Bands(0).Columns("Elija").SortIndicator = ssSortIndicatorAscending
    lcMensaje = ""
    Set oRow = Me.grdBusqueda.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        Do While oRow.HasNextSibling
            
            If oRow.Cells("Elija").Value = True Then
               If oRow.Cells("idTipo").Value = sghActividadesTipo.TipoLAB And oRow.Cells("elijaLAB").Value = "" Then
                  lcMensaje = lcMensaje & "para el Grupo: " & oRow.Cells("grupo").Value & " SubGrupo:" & _
                                          oRow.Cells("subGrupo").Value & " debe ingresar LAB" & Chr(13)
               End If
               If oRow.Cells("idTipo").Value <> sghActividadesTipo.TipoLAB And oRow.Cells("nombre").Value = "" Then
                  lcMensaje = lcMensaje & "para el Grupo: " & oRow.Cells("grupo").Value & " SubGrupo:" & _
                                          oRow.Cells("subGrupo").Value & " debe ingresar " & _
                                          IIf(oRow.Cells("idTipo").Value = sghActividadesTipo.TipoCPT, "CPT", "DX") & Chr(13)
               End If
            End If
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
        Loop
    Else
        lcMensaje = lcMensaje & "No hay registros" & Chr(13)
    End If
    Set oRow = Nothing
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, ""
       Exit Function
    End If
    ValidarReglas = True
End Function















