VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FLETE01_VB6 
   Caption         =   "TABLA CODIFICADA DE PAISES"
   ClientHeight    =   5610
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   5850
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5850
      TabIndex        =   7
      Top             =   5010
      Width           =   5850
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1213
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         Height          =   300
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4675
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   300
         Left            =   3521
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   2367
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edici�n"
         Height          =   300
         Left            =   1213
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ag&regar"
         Height          =   300
         Left            =   59
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5850
      TabIndex        =   1
      Top             =   5310
      Width           =   5850
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "FLETE01_VB6.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "FLETE01_VB6.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "FLETE01_VB6.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "FLETE01_VB6.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   3360
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FLETE01_VB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'*  Forma: TABLA CODIFICADA DE PAISES.
'*  Creado x el Wizard de Formularios Visual Basic v 6.0.
'*  Formulario Tipo Matriz con Codigo ADO conecctado via ODBC a la B.D.
'*  el 18-10-2006.
'*  Actualizado x Henry J. Pulgar B.
'*  Ultima fecha actualizado:  02-11-2006.
'*     Trabajando con MATRICES >>>>     MODULO TIPO MODELO.
'*********************************************************************************
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim StringConeccionDesfle00 As String  'Agregado x Henry Pulgar

'* Autor: HPB.  23-10-2006.
Private Sub ACTUALIZAR_ATTBR_GRID(TipoBoolean As Boolean)
     grdDataGrid.AllowAddNew = TipoBoolean
     grdDataGrid.AllowDelete = TipoBoolean
     grdDataGrid.AllowUpdate = TipoBoolean
End Sub

'************************************************************************
'* Autor: Henry Pulgar B.
'* Fecha: 26-10-2006.
'* General Porpuse: Coneccion ODBC utillizando CODIGO ADO
'* Caracteristica : De esta menera. ( Puede ser hecho de otra  forma )
'***********************************************************************
Private Sub INICIAR_VARIABLES_GLOBALES()
    StringConeccionDesfle00 = "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$desfle00;pwd=ops$desfle00;"
End Sub

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  INICIAR_VARIABLES_GLOBALES
  db.CursorLocation = adUseClient
  db.Open StringConeccionDesfle00 '"PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$desfle00;pwd=ops$desfle00;"
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select C1_COD_PAIS Cod_Pais, " & _
                            "C1_NOMBRE_PAIS Nombre_Pais " & _
                            "from FLETE01_DAT " & _
                            "Order by C1_COD_PAIS", db, adOpenStatic, adLockOptimistic

  Set grdDataGrid.DataSource = adoPrimaryRS
  
  mbDataChanged = False
  
  ACTUALIZAR_ATTBR_GRID (False)   'Agregado.
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'Esto cambiar� el tama�o de la cuadr�cula al cambiar el tama�o del formulario
  grdDataGrid.Height = Me.ScaleHeight - 30 - picButtons.Height - picStatBox.Height
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrar� la posici�n de registro actual para este Recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aqu� se coloca el c�digo de validaci�n
  'Se llama a este evento cuando ocurre la siguiente acci�n
  Dim bCancel As Boolean
  '
  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select
  '
  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  grdDataGrid.AllowAddNew = True  'Agregado by Myself
  grdDataGrid.AllowUpdate = True  'Agregado x myself
  On Error GoTo AddErr
  adoPrimaryRS.MoveLast
  adoPrimaryRS.AddNew
  grdDataGrid.SetFocus
  '
  cmdEdit_Click                    'Agrgado by myself.
  Exit Sub
AddErr:
  'MsgBox Err.Description, vbAbortRetryIgnore, "ATENCION"
   MsgBox Err.Description
End Sub
'************************************************************************
'* Autor: Henry Pulgar B.
'* Fecha: 26-20-2006.
'* General Porpuse: Coneccion ODBC utillizando CODIGO ADO
'* Caracteristica : De esta menera. ( Puede ser hecho de otra  forma )
'***********************************************************************
Private Function VALIDAR_DESTINO(CodPais As String) As Boolean
   Dim CadenaSql As String
   Dim SqlCriterio As String
   '
   Dim FLETE03_DAT As Connection    'Coneccion.Ubicar la tabla.
   Dim FLETE03_REC As New Recordset 'Coneccion crear el Record set
   '
   Set FLETE03_DAT = New Connection
   Set FLETE03_REC = New Recordset
   '
   FLETE03_DAT.CursorLocation = adUseClient
   FLETE03_DAT.Open StringConeccionDesfle00 '"PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$desfle00;pwd=ops$desfle00;"
   '
   '*** Cuerpo del query: Invocar todos los registros para luego ejecutar un find,...
        CadenaSql = "select * " & _
                    "from    FLETE03_DAT " & _
                    "where   C3_COD_PAIS is not null "
        '
        FLETE03_REC.Open CadenaSql, FLETE03_DAT, adOpenStatic, adLockOptimistic
        If (FLETE03_REC.RecordCount > 0) Then  'La Tabla posee registros ?
            SqlCriterio = " C3_COD_PAIS =" + "'" + CodPais + "'"
            'MsgBox SqlCriterio
            FLETE03_REC.Find (SqlCriterio)   'Buscar Utilizando comando Find
            If FLETE03_REC.EOF Then
               'MsgBox "Registro no encontrado puede ser eliminado"
               VALIDAR_DESTINO = False
            Else
               'MsgBox "Registro encontrado en la tabla FLETE03_DAT"
               VALIDAR_DESTINO = True
            End If
            Exit Function
        Else
           '** MsgBox "La tabla esta Vacia", vbCritical, "ATENCION"
        End If
   '***Cerrar coneccion ****
   FLETE03_REC.Close
   FLETE03_DAT.Close
End Function 'VALIDAR_DESTINO

'***********************************************************
'* Modificado x Henry J. Pulgar B.
'* el 02-11-2006.
'***********************************************************
Private Sub cmdDelete_Click()
  Dim Botones, Respuesta
  On Error GoTo DeleteErr
  '
  Botones = vbYesNo + vbQuestion + vbDefaultButton1
  Respuesta = MsgBox("Deseas eliminar este Registro ?", Botones, "ATENCION")
  If (Respuesta = vbYes) Then
    If Not VALIDAR_DESTINO(adoPrimaryRS("COD_PAIS")) Then
       With adoPrimaryRS
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
       End With
    Else
       Beep
       MsgBox "El Pais no puede ser eliminado porque posee registro en la tabla Destino", vbCritical, "ATENCION"
    End If
  End If   'If Respuesta,...
    Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto s�lo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS

  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  grdDataGrid.AllowUpdate = True          ' Modificado by myself
  '
  On Error GoTo EditErr
  '
  lblStatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  '
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
     adoPrimaryRS.Bookmark = mvBookMark
  Else
     adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  '
  ACTUALIZAR_ATTBR_GRID (False)    'Agrgado by myself
End Sub


' Autor Henry J. Pulgar B.
' Fecha : 19-10-2006.
Private Sub ACTUALIZAR_CAMPOS_GRID()
    Dim Cadena As String
    Dim ValorOriginal As String
    '
    If (grdDataGrid.AllowAddNew <> True) And (grdDataGrid.AllowUpdate = True) Then
        ValorOriginal = adoPrimaryRS("COD_PAIS").OriginalValue  'Porque este campo no deberia ser modificado; es un campo clave.
        adoPrimaryRS("COD_PAIS") = ValorOriginal   'Solo en modo Update
    End If
    Cadena = UCase(adoPrimaryRS("COD_PAIS").Value)
    adoPrimaryRS("COD_PAIS") = Cadena
    Cadena = UCase(adoPrimaryRS("NOMBRE_PAIS").Value)
    adoPrimaryRS("NOMBRE_PAIS").Value = UCase(Cadena)
End Sub

'************************************************************************
'* Autor: Henry Pulgar B.
'* Fecha: 27-20-2006.
'* General Porpuse:
'* Caracteristica :
'***********************************************************************
Function VALIDAR_REGISTRADO(CodPais As String) As Boolean
   Dim CadenaSql As String
   Dim SqlCriterio As String
   '
   Dim FLETE01_DAT As Connection    'Coneccion.Ubicar la tabla.
   Dim FLETE01_REC As New Recordset 'Coneccion crear el Record set
   '
   Set FLETE01_DAT = New Connection
   Set FLETE01_REC = New Recordset
   '
   FLETE01_DAT.CursorLocation = adUseClient
   FLETE01_DAT.Open StringConeccionDesfle00
   '
   '*** Cuerpo del query: Invocar todos los registros para luego ejecutar un find,...
        CadenaSql = "select * " & _
                    "from    FLETE01_DAT " & _
                    "where   C1_COD_PAIS is not null "
        '
        FLETE01_REC.Open CadenaSql, FLETE01_DAT, adOpenStatic, adLockOptimistic
        'MsgBox "Record count=" + Str(FLETE01_REC.RecordCount)
        If (FLETE01_REC.RecordCount > 0) Then
            SqlCriterio = " C1_COD_PAIS =" + "'" + CodPais + "'"
            'MsgBox SqlCriterio
            FLETE01_REC.Find (SqlCriterio)   'Buscar Utilizando comando Find
            If FLETE01_REC.EOF Then
               'MsgBox "Registro no encontrado; CONTINUAR
               VALIDAR_REGISTRADO = False
               Exit Function
            Else
               'MsgBox "Registro encontrado en la tabla FLETE01_DAT; se esta violando el codigo de unicidad"
                VALIDAR_REGISTRADO = True
                Exit Function
            End If
        Else
           Beep
           'MsgBox "La tabla esta Vacia", vbCritical, "ATENCION"
        End If
   '***Cerrar coneccion ****
   FLETE01_REC.Close
   FLETE01_DAT.Close
End Function 'VALIDAR_REGISTRADO(CodPais As String)

' Modificado by H.P.B el 19-10-2006.
Private Sub cmdUpdate_Click()
  Dim Exitoso As Boolean
  On Error GoTo UpdateErr
   
  ' H.P.B.
  ' Validar_Campos
  '
  ACTUALIZAR_CAMPOS_GRID               'Agregado x Henry Pulgar.
  Exitoso = True
  If grdDataGrid.AllowAddNew = True Then 'Estoy en modo de insercion
     If Not VALIDAR_REGISTRADO(adoPrimaryRS("COD_PAIS")) Then
          adoPrimaryRS.UpdateBatch adAffectAll
          If mbAddNewFlag Then
             adoPrimaryRS.MoveLast              'va al nuevo registro
          End If
     Else
          Beep
          MsgBox "Codigo del Pais ya esta registrado en la Base de Datos", vbCritical, "ATENCION"
          Exitoso = False
     End If
  ElseIf (grdDataGrid.AllowUpdate = True) Then   'Estoy en modo "update"
          adoPrimaryRS.UpdateBatch adAffectAll
          If mbAddNewFlag Then
             adoPrimaryRS.MoveLast              'va al nuevo registro
          End If
  End If  'grdDataGrid = Trueb ?,...
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  ACTUALIZAR_ATTBR_GRID (False)        ' Modificado by myself
  If Not Exitoso Then
     Form_Load   'Refrescar la matriz
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  '  adoPrimaryRS.Close
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atr�s
    adoPrimaryRS.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atr�s
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

'-----------------------------------eof(FLETE01_VB6)------------------------------
