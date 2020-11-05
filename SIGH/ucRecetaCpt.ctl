VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucRecetaCpt 
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   ScaleHeight     =   1665
   ScaleWidth      =   11250
   Begin VB.Frame FraOtrosCpt 
      Caption         =   "Otros CPT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      Begin VB.CommandButton btnOtrosCpt 
         DisabledPicture =   "ucRecetaCpt.ctx":0000
         DownPicture     =   "ucRecetaCpt.ctx":03E9
         Height          =   315
         Left            =   10920
         Picture         =   "ucRecetaCpt.ctx":07F5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdOtrosCpt 
         Height          =   1395
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   2461
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdOtrosCpt"
      End
      Begin VB.Label lblOtrosCpt 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1170
         TabIndex        =   3
         Top             =   15
         Width           =   90
      End
   End
End
Attribute VB_Name = "ucRecetaCpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim ml_IdTipoFinanciamiento As Long
Dim lcDx As String
Dim oRsOtrosCpt As New Recordset
Dim lnMaximoItems As Long
Property Let MaximoItems(lValue As Long)
    lnMaximoItems = lValue
End Property
Property Let idTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Let Dx(lValue As String)
   lcDx = lValue
End Property


Private Sub btnOtrosCpt_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsServ As New Recordset
    Dim lnIdProducto As Long, lcProducto As String, lbContinuar As Boolean, lnTotalReg As Long, lnPrecio As Double
    Dim lbVariosFormatosFua As Boolean, lbSeEligioPaquete As Boolean
    oPaquetesBuscar.idPuntoCarga = 2600   'sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion   '2600
    oPaquetesBuscar.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    'oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
           Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
           Set oRsDevuelveTodosLosItemsServ = oPaquetesBuscar.DevuelveTodosLosItemsServ
           oRsItemsElegidos.MoveFirst
           Do While Not oRsItemsElegidos.EOF
                lnIdProducto = oRsItemsElegidos.Fields!idProducto
                lcProducto = oRsItemsElegidos.Fields!Producto
                lbContinuar = True
                lnTotalReg = oRsOtrosCpt.RecordCount
                If lnTotalReg >= lnMaximoItems Then
                   MsgBox "Solo puede registrar hasta " & Trim(Str(lnMaximoItems)) & " items", vbInformation, "Receta"
                   lbContinuar = False
                End If
                '
                If lnTotalReg > 0 And lbContinuar = True Then        'debb-24/06/2015
                   oRsOtrosCpt.MoveFirst
                   oRsOtrosCpt.Find "id=" & lnIdProducto
                   If Not oRsOtrosCpt.EOF Then
                      lbContinuar = False
                   End If
                End If
                If lbContinuar = True Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, ml_IdTipoFinanciamiento)
                    oRsOtrosCpt.AddNew
                    If lbVariosFormatosFua = True Then
                       oRsOtrosCpt.Fields!FUA = 1
                    End If
                    oRsOtrosCpt.Fields!ID = lnIdProducto
                    oRsOtrosCpt.Fields!procedimiento = lcProducto
                    If lbSeEligioPaquete = True Then
                       oRsOtrosCpt.Fields!Cantidad = oRsItemsElegidos!Cantidad
                    Else
                       oRsOtrosCpt.Fields!Cantidad = 1
                    End If
                    oRsOtrosCpt.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsOtrosCpt.Fields!hayCpt = True
                    End If
                    oRsOtrosCpt.Fields!saldoActual = 0
                    If lcDx <> "" Then
                       oRsOtrosCpt.Fields!Dx = lcDx
                    End If
                    oRsOtrosCpt.Fields!idPuntoCarga = oRsItemsElegidos!idPuntoCarga
                    oRsOtrosCpt.Update
                    sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)
                End If
                oRsItemsElegidos.MoveNext
           Loop
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsServ = Nothing
End Sub

Public Sub Inicializar()
    CreaTemporales True, True
    InicializarLaGrilla grdOtrosCpt
End Sub

Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("SaldoActual").Hidden = True
         oGrilla.Bands(0).Columns("Precio").Hidden = True
         oGrilla.Bands(0).Columns("idDosisRecetada").Hidden = True
         oGrilla.Bands(0).Columns("idEstadoDetalle").Hidden = True
         oGrilla.Bands(0).Columns("MotivoAnulacionMedico").Hidden = True
         oGrilla.Bands(0).Columns("Procedimiento").Header.Caption = "Procedimiento"
         oGrilla.Bands(0).Columns("Procedimiento").Width = 6000
         oGrilla.Bands(0).Columns("Procedimiento").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("Cantidad").Width = 400
         oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("hayCpt").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("hayCpt").Width = 400
         oGrilla.Bands(0).Columns("observaciones").Width = 3200
         oGrilla.Bands(0).Columns("Dx").Width = 600
End Sub

Sub CreaTemporales(lbHabilitaFrame As Boolean, _
                   lbSoloLimpiaOtrosCpt As Boolean)
    On Error Resume Next
    If lbSoloLimpiaOtrosCpt = True Then
        If oRsOtrosCpt.State = 1 Then Set oRsOtrosCpt = Nothing
        With oRsOtrosCpt
              .Fields.Append "Receta", adInteger
              .Fields.Append "Id", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "idDosisRecetada", adInteger
              .Fields.Append "HayCpt", adBoolean
              .Fields.Append "Precio", adDouble
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .Fields.Append "idPuntoCarga", adInteger
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdOtrosCpt.DataSource = oRsOtrosCpt
        mo_Apariencia.ConfigurarFilasBiColores grdOtrosCpt, sighEntidades.GrillaConFilasBicolor
        grdOtrosCpt.Caption = ""
        If lbHabilitaFrame = True Then
           FraOtrosCpt.Enabled = True
        End If
    End If
End Sub

Function DevuelvePrecioItem(lnIdProducto As Long, ml_IdTipoFinanciamiento As Long, Optional oConexion1 As Connection) As Double
      Dim oRsTmp As New Recordset
      Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
      Set oRsTmp = mo_ReglasComunes.FactCatalogoServiciosHospXfiltro("idProducto=" & lnIdProducto & " and idTipoFinanciamiento=" & ml_IdTipoFinanciamiento)
      DevuelvePrecioItem = 0
      If oRsTmp.RecordCount > 0 Then
         DevuelvePrecioItem = oRsTmp.Fields!PrecioUnitario
      End If
      Set mo_ReglasComunes = Nothing
End Function

Private Sub grdOtrosCpt_AfterRowsDeleted()
   Set grdOtrosCpt.DataSource = oRsOtrosCpt
End Sub

Public Function DevuelveOtrosCpt() As Recordset
    Set DevuelveOtrosCpt = oRsOtrosCpt
End Function

Public Sub InhabilitaControles(lclblOtrosCpt As String)
    FraOtrosCpt.Enabled = False
    lblOtrosCpt.Caption = lclblOtrosCpt
End Sub

Public Sub CargaDatosAcontroles(oRsDetalleReceta As Recordset)
    If oRsDetalleReceta.RecordCount > 0 Then
       Dim oRsTmp198 As New Recordset
       oRsDetalleReceta.MoveFirst
       Do While Not oRsDetalleReceta.EOF
            Set oRsTmp198 = mo_ReglasComunes.SeleccionarPuntosDeCargaSegunFiltro("idPuntoCarga=" & Trim(Str(oRsDetalleReceta!idPuntoCarga)))
            If oRsTmp198.RecordCount > 0 Then
                oRsOtrosCpt.AddNew
                oRsOtrosCpt.Fields!Receta = oRsDetalleReceta!idReceta
                oRsOtrosCpt.Fields!ID = oRsDetalleReceta.Fields!idItem
                oRsOtrosCpt.Fields!procedimiento = Left(Trim(oRsDetalleReceta.Fields!Producto) & " <<" & Trim(oRsTmp198!descripcion) & ">>", 255)
                oRsOtrosCpt.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                oRsOtrosCpt.Fields!precio = oRsDetalleReceta.Fields!precio
                If oRsDetalleReceta.Fields!precio > 0 Then
                   oRsOtrosCpt.Fields!hayCpt = True
                End If
                oRsOtrosCpt.Fields!saldoActual = 0
                oRsOtrosCpt.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                oRsOtrosCpt.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                oRsOtrosCpt.Fields!idPuntoCarga = oRsDetalleReceta!idPuntoCarga
                oRsOtrosCpt.Update
            End If
            oRsDetalleReceta.MoveNext
       Loop
    End If
End Sub
