VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_DataEntryClient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tbl_Client"
   ClientHeight    =   7200
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   19440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   19440
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   19440
      TabIndex        =   24
      Top             =   5865
      Width           =   19440
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4680
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3525
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2370
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1215
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   60
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Search_Active"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   7440
      TabIndex        =   23
      Top             =   1425
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Phone_Key"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   1455
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Contact_Key"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   12960
      TabIndex        =   19
      Top             =   1140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Client_Type"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   7440
      TabIndex        =   17
      Top             =   1065
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Client_Status"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   1095
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Client_Number"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   12960
      TabIndex        =   13
      Top             =   780
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Client_Name"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   11
      Top             =   705
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Client_Key"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   735
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Client_Code"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   12960
      TabIndex        =   7
      Top             =   300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Client_Category"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   7440
      TabIndex        =   5
      Top             =   345
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Client_Alias"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Billing_Active"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6870
      Width           =   19440
      _ExtentX        =   34290
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=C:\source\camel\CAMELSDATA.mdb;"
      OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=C:\source\camel\CAMELSDATA.mdb;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frm_DataEntryClient.frx":0000
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      Caption         =   "Search_Active:"
      Height          =   255
      Index           =   11
      Left            =   5520
      TabIndex        =   22
      Top             =   1425
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Phone_Key:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Contact_Key:"
      Height          =   255
      Index           =   9
      Left            =   11040
      TabIndex        =   18
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Client_Type:"
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   16
      Top             =   1065
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Client_Status:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Client_Number:"
      Height          =   255
      Index           =   6
      Left            =   11040
      TabIndex        =   12
      Top             =   780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Client_Name:"
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   10
      Top             =   705
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Client_Key:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   735
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Client_Code:"
      Height          =   255
      Index           =   3
      Left            =   11040
      TabIndex        =   6
      Top             =   300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Client_Category:"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   345
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Client_Alias:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Billing_Active:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "Frm_DataEntryClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

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

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

