VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mUbigeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubigeos GalenHos vs LolCli"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   585
      Left            =   30
      TabIndex        =   10
      Top             =   8460
      Width           =   12855
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Top             =   210
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmbCargaUbigeos 
         Caption         =   "Carga últimos UBIGEOS desde LolCli"
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   150
         Width           =   3045
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ubigeo en LolCli"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   12885
      Begin MSDataGridLib.DataGrid grdLolCli 
         Height          =   3105
         Left            =   60
         TabIndex        =   3
         Top             =   270
         Width           =   12645
         _ExtentX        =   22304
         _ExtentY        =   5477
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   18
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "DptoLolCli"
            Caption         =   "Dpto (LolCli)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "ProvLolCli"
            Caption         =   "Prov (lolCli)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "DistritoLolCli"
            Caption         =   "Distrito (lolcli)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "UbigeoLolCli"
            Caption         =   "Ubigeo (lolCli)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Dpto"
            Caption         =   "Dpto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Prov"
            Caption         =   "Provincia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "distrito"
            Caption         =   "Distrito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.Label txtFilas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Filas"
         Height          =   255
         Left            =   12030
         TabIndex        =   9
         Top             =   3480
         Width           =   675
      End
      Begin VB.Label txtFilasPorActualizar 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Filas"
         Height          =   255
         Left            =   5070
         TabIndex        =   8
         Top             =   3480
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENTER= elimina DATOS en las columnas: Dpto, Provincia, Distrito"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   3480
         Width           =   4770
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ubigeo en GalenHos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   30
      TabIndex        =   0
      Top             =   3930
      Width           =   12855
      Begin MSDataGridLib.DataGrid grdGalenHos 
         Height          =   3435
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   6059
         _Version        =   393216
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "dpto"
            Caption         =   "Departamento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "prov"
            Caption         =   "Provincia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "dist"
            Caption         =   "Distrito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Doble Clic= Solo añade datos a columnas: Dpto, Provincia  (hacia el Ubigeo de LOLCLI marcado)"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   4110
         Width           =   6960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENTER= añade datos a columnas: Dpto, Provincia, Distrito  (hacia el Ubigeo de LOLCLI marcado)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3780
         Width           =   7005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BARRA= Solo añade datos a columnas: Dpto (hacia el Ubigeo de LOLCLI marcado)"
         Height          =   255
         Left            =   7200
         TabIndex        =   5
         Top             =   3780
         Width           =   6000
      End
   End
End
Attribute VB_Name = "mUbigeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Ubigeos Lolcli
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Dim oRsUbigeoLol As New ADODB.Recordset
Dim oRsUbigeoGal As New ADODB.Recordset
Dim lnNumUbigeoLol As Long
Dim oRsTmpD As New ADODB.Recordset
Dim oRsBus As New ADODB.Recordset

Private Sub cmbCargaUbigeos_Click()
  If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    Me.MousePointer = 11
    'carga por Direccion
    oRsTmpD.Open "select ubicod from Pacientes where not (ubicod is null) order by ubicod", wxConexion, adOpenKeyset, adLockOptimistic
    If oRsTmpD.RecordCount > 0 Then
       CargaUbigeo 1
    End If
    oRsTmpD.Close
    'carga por Nacimiento
    oRsTmpD.Open "select pacLun from Pacientes where not (pacLun is null) order by pacLun", wxConexion, adOpenKeyset, adLockOptimistic
    If oRsTmpD.RecordCount > 0 Then
       CargaUbigeo 2
    End If
    oRsTmpD.Close
    Me.MousePointer = 1
    Unload Me
  End If
End Sub

Sub CargaUbigeo(lnCampo As Integer)
       ProgressBar1.Min = 0
       ProgressBar1.Max = oRsTmpD.RecordCount
       lnReg = 1
       oRsTmpD.MoveFirst
       Do While Not oRsTmpD.EOF
          ProgressBar1.Value = lnReg: lnReg = lnReg + 1
          Select Case lnCampo
          Case 1
             lcUbigeo = oRsTmpD!ubicod
          Case 2
             lcUbigeo = oRsTmpD!pacLun
          End Select
          lcNombre = ""
          oRsBus.Open "select ubides from ubigeo where ubicod='" & lcUbigeo & "'", wxConexion, adOpenKeyset, adLockOptimistic
          If oRsBus.RecordCount > 0 Then
             lcNombre = oRsBus.Fields!ubides
          End If
          oRsBus.Close
          lbEsNuevo = True
          If lnNumUbigeoLol > 0 Then
             oRsUbigeoLol.MoveFirst
             oRsUbigeoLol.Find "ubigeoLolcli='" & lcUbigeo & "'"
             If Not oRsUbigeoLol.EOF Then
                lbEsNuevo = False
             End If
          End If
          If lbEsNuevo Then
             If Right(lcUbigeo, 6) = "000000" Then
                'No se eligio Ni Depart/Prov/Dist
                lcDpto = lcNombre
                lcProv = ""
                lcDistrito = ""
             ElseIf Right(lcUbigeo, 4) = "0000" Then
                'Se eligio solo Departamento
                lcDpto = lcNombre
                lcProv = ""
                lcDistrito = ""
             ElseIf Right(lcUbigeo, 2) = "00" Then
                'Se eligio solo Provincia
                lcDpto = ""
                oRsBus.Open "select ubides from ubigeo where ubicod='" & Left(lcUbigeo, 2) & "0000'", wxConexion, adOpenKeyset, adLockOptimistic
                If oRsBus.RecordCount > 0 Then
                   lcDpto = oRsBus.Fields!ubides
                End If
                oRsBus.Close
                lcProv = lcNombre
                lcDistrito = ""
             Else
                'Se eligio Distrito
                lcDpto = ""
                oRsBus.Open "select ubides from ubigeo where ubicod='" & Left(lcUbigeo, 2) & "0000'", wxConexion, adOpenKeyset, adLockOptimistic
                If oRsBus.RecordCount > 0 Then
                   lcDpto = oRsBus.Fields!ubides
                End If
                oRsBus.Close
                lcProv = ""
                oRsBus.Open "select ubides from ubigeo where ubicod='" & Left(lcUbigeo, 4) & "00'", wxConexion, adOpenKeyset, adLockOptimistic
                If oRsBus.RecordCount > 0 Then
                   lcProv = oRsBus.Fields!ubides
                End If
                oRsBus.Close
                lcDistrito = lcNombre
             End If
             
             oRsUbigeoLol.AddNew
             oRsUbigeoLol.Fields!ubigeoLolcli = lcUbigeo
             oRsUbigeoLol.Fields!dptoLolCli = lcDpto
             oRsUbigeoLol.Fields!provLolCli = lcProv
             oRsUbigeoLol.Fields!distritoLolCli = lcDistrito
             oRsUbigeoLol.Fields!IdDepartamento = 0
             oRsUbigeoLol.Fields!IdProvincia = 0
             oRsUbigeoLol.Fields!IdDistrito = 0
             oRsUbigeoLol.Fields!dpto = ""
             oRsUbigeoLol.Fields!prov = ""
             oRsUbigeoLol.Fields!Distrito = ""
             oRsUbigeoLol.Update
          End If
          Select Case lnCampo
          Case 1
                Do While Not oRsTmpD.EOF And lcUbigeo = oRsTmpD.Fields!ubicod
                   oRsTmpD.MoveNext
                   If oRsTmpD.EOF Then
                      Exit Do
                   End If
                Loop
          Case 2
                Do While Not oRsTmpD.EOF And lcUbigeo = oRsTmpD!pacLun
                   oRsTmpD.MoveNext
                   If oRsTmpD.EOF Then
                      Exit Do
                   End If
                Loop
          End Select
       Loop
     
End Sub



Private Sub Form_Load()
   oRsUbigeoLol.Open "select * from lolcliUbigeo order by DptoLolCli,ProvLolcli,distritoLolCli", wxConexionRed, adOpenKeyset, adLockOptimistic
   Set grdLolCli.DataSource = oRsUbigeoLol
   lnNumUbigeoLol = oRsUbigeoLol.RecordCount
   txtFilas.Caption = "Nº total de Filas: " & Trim(Str(lnNumUbigeoLol))
   oRsTmpD.Open "select * from lolcliUbigeo where idDepartamento=0 and idProvincia=0 and idDistrito=0", wxConexionRed, adOpenKeyset, adLockOptimistic
   txtFilasPorActualizar.Caption = "Nº filas que faltan UBIGEO: " & Trim(Str(oRsTmpD.RecordCount))
   oRsTmpD.Close
   CargaUbigeosGalenHos
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo ErrUn
   oRsUbigeoLol.Close
   oRsUbigeoGal.Close
   oRsTmpD.Close
   oRsBus.Close
   Exit Sub
ErrUn:
   Resume Next
End Sub


Sub CargaUbigeosGalenHos()
    lcSql = "SELECT   dbo.Departamentos.Nombre AS dpto, dbo.Provincias.Nombre AS prov, dbo.Distritos.Nombre AS dist,  dbo.Distritos.IdDistrito, " & _
            "          dbo.Distritos.IdProvincia , dbo.Provincias.IdDepartamento" & _
            " FROM         dbo.Distritos LEFT OUTER JOIN dbo.Provincias ON dbo.Distritos.IdProvincia = dbo.Provincias.IdProvincia LEFT OUTER JOIN" & _
                    "  dbo.Departamentos ON dbo.Provincias.IdDepartamento = dbo.Departamentos.IdDepartamento" & _
            " ORDER BY dbo.Departamentos.Nombre, dbo.Provincias.Nombre, dbo.Distritos.Nombre"
    oRsUbigeoGal.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
    Set grdGalenHos.DataSource = oRsUbigeoGal
End Sub


Private Sub grdGalenHos_DblClick()
       oRsUbigeoLol.Fields!dpto = oRsUbigeoGal.Fields!dpto
       oRsUbigeoLol.Fields!prov = oRsUbigeoGal.Fields!prov
       oRsUbigeoLol.Fields!Distrito = ""
       oRsUbigeoLol.Fields!IdDepartamento = oRsUbigeoGal.Fields!IdDepartamento
       oRsUbigeoLol.Fields!IdProvincia = oRsUbigeoGal.Fields!IdProvincia
       oRsUbigeoLol.Fields!IdDistrito = 0
       oRsUbigeoLol.Update
End Sub

Private Sub grdGalenHos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       oRsUbigeoLol.Fields!dpto = oRsUbigeoGal.Fields!dpto
       oRsUbigeoLol.Fields!prov = oRsUbigeoGal.Fields!prov
       oRsUbigeoLol.Fields!Distrito = oRsUbigeoGal.Fields!dist
       oRsUbigeoLol.Fields!IdDepartamento = oRsUbigeoGal.Fields!IdDepartamento
       oRsUbigeoLol.Fields!IdProvincia = oRsUbigeoGal.Fields!IdProvincia
       oRsUbigeoLol.Fields!IdDistrito = oRsUbigeoGal.Fields!IdDistrito
       oRsUbigeoLol.Update
'       oRsUbigeoLol.Requery
'       oRsUbigeoLol.MoveFirst
'       oRsUbigeoLol.Find "ubigeoLolcli='" & lcUbigeo & "'"
    End If
    If KeyAscii = 32 Then
       oRsUbigeoLol.Fields!dpto = oRsUbigeoGal.Fields!dpto
       oRsUbigeoLol.Fields!prov = ""
       oRsUbigeoLol.Fields!Distrito = ""
       oRsUbigeoLol.Fields!IdDepartamento = oRsUbigeoGal.Fields!IdDepartamento
       oRsUbigeoLol.Fields!IdProvincia = 0
       oRsUbigeoLol.Fields!IdDistrito = 0
       oRsUbigeoLol.Update
    End If
End Sub


Private Sub grdLolCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       oRsUbigeoLol.Fields!dpto = ""
       oRsUbigeoLol.Fields!prov = ""
       oRsUbigeoLol.Fields!Distrito = ""
       oRsUbigeoLol.Fields!IdDepartamento = 0
       oRsUbigeoLol.Fields!IdProvincia = 0
       oRsUbigeoLol.Fields!IdDistrito = 0
       oRsUbigeoLol.Update
       'oRsUbigeoLol.Requery
    End If
End Sub
