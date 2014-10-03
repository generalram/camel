VERSION 5.00
Begin VB.Form Frm_Main 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   3120
   ClientTop       =   420
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmd_Help 
      Caption         =   "Help"
      Height          =   360
      Left            =   8040
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmd_About 
      Caption         =   "About"
      Height          =   375
      Left            =   8040
      TabIndex        =   37
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmd_Search 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox cbo_MatterList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   4800
      Width           =   5775
   End
   Begin VB.CommandButton cmd_ClearMatterList 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmd_ClearAll 
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   32
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox chk_MatterBillingActive 
      Caption         =   "Matter Billing Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   30
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CheckBox chk_MatterSearchActive 
      Caption         =   "Matter Search Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   29
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CheckBox chk_ClientBillingActive 
      Caption         =   "Client Billing Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   28
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CheckBox chk_ClientSearchActive 
      Caption         =   "Client Search Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   27
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmd_ClearAttorney 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   26
      Top             =   4080
      Width           =   735
   End
   Begin VB.ComboBox cbo_AttorneyList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   4200
      Width           =   5775
   End
   Begin VB.CommandButton cmd_ClearMatterNumber 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmd_ClearMatterName 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmd_ClearClientNumber 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmd_ClearClientName 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   21
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txt_MatterName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txt_MatterNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.ComboBox cbo_MatterNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   8
      Text            =   "Matter Number List"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.ComboBox cbo_MatterName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      Text            =   "Matter Name List"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.ComboBox cbo_ClientName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      Text            =   "Client Name List"
      Top             =   720
      Width           =   2655
   End
   Begin VB.ComboBox cbo_ClientNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      Text            =   "Client Number List"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txt_ClientNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txt_ClientName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lbl_Title2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Client And Matter Electronic Legal System)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   0
      TabIndex        =   36
      Top             =   240
      Width           =   8115
   End
   Begin VB.Label lbl_SearchStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   35
      Top             =   3600
      Width           =   5655
   End
   Begin VB.Label lbl_MatterList 
      Caption         =   "Matter List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   34
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lbl_Copyright 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Copyright XORMAD 2007, 2008, All rights reserved. Duplication of this product without authorization from XORMAD is prohibited."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   0
      TabIndex        =   31
      Top             =   5760
      Width           =   8775
   End
   Begin VB.Label lbl_Attorney 
      Caption         =   "Attorney List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   25
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lbl_MatterName 
      Caption         =   "Matter Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lbl_MatterNumber 
      Caption         =   "Matter Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lbl_Or1 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lbl_Title 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAMEL Search"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8115
   End
   Begin VB.Label lbl_ClientNumber 
      Caption         =   "Client Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lbl_ClientName 
      Caption         =   "Client Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_About_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub cmd_ClearClientName_Click()
    Frm_Main.txt_ClientName.Text = ""
    Frm_Main.cbo_ClientName.Text = "Client Name List"
End Sub

Private Sub cmd_ClearClientNumber_Click()
    Frm_Main.txt_ClientNumber.Text = ""
    Frm_Main.cbo_ClientNumber.Text = "Client Number List"
End Sub

Private Sub cmd_ClearMatterName_Click()
    Frm_Main.txt_MatterName.Text = ""
    Frm_Main.cbo_MatterName.Text = "Matter Name List"
End Sub

Private Sub cmd_ClearMatterNumber_Click()
    Frm_Main.txt_MatterNumber.Text = ""
    Frm_Main.cbo_MatterNumber.Text = "Matter Number List"
End Sub

Private Sub cmd_Exit_Click()
    ExitProgram
End Sub

Private Sub cmd_Help_Click()
    ' Disabled for now, not visible until help is added to the program.
End Sub

Private Sub cmd_Search_Click()
    Dim iresponse As Integer
    Dim s_Query As String
    Dim Conn1 As New ADODB.Connection
    Dim RS1 As New ADODB.Recordset
    Dim FileDatabase As String
    Dim s_ClientName As String
    Dim s_ClientNumber As String
    Dim b_cboClientName As Boolean
    Dim b_txtClientName As Boolean
    
    
    FileDatabase = App.Path & "\CAMELSDATA.mdb"

    

    If XorTrim(Frm_Main.txt_ClientName.Text) = "" And XorTrim(Frm_Main.txt_ClientNumber.Text) = "" And XorTrim(Frm_Main.cbo_ClientName.Text) = "" Then
     
        iresponse = MsgBox("Please enter client information for a search.", vbCritical, "Missing Client Info")
     
    Else
        'iresponse = MsgBox("Entered Client Information", vbOKOnly, "Client Information entered")
        'Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileDatabase
        
        Conn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileDatabase
        'Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:nwind.mdb;"
        
        If XorTrim("" & Frm_Main.txt_ClientName.Text) = "" Then
            s_ClientName = XorTrim("" & Frm_Main.cbo_ClientName.Text)
            b_txtClientName = False
        Else
            s_ClientName = XorTrim("" & Frm_Main.txt_ClientName.Text)
            b_txtClientName = True
        End If
        
         If XorTrim("" & Frm_Main.cbo_ClientName.Text) = "Client Name List" Then
            s_ClientName = XorTrim("" & Frm_Main.txt_ClientName.Text)
            b_cboClientName = False
        Else
            s_ClientName = XorTrim("" & Frm_Main.cbo_ClientName.Text)
            b_cboClientName = True
        End If
       
        s_Query = "SELECT CLIENT_NAME, CLIENT_NUMBER "
        s_Query = s_Query & "FROM TBL_CLIENT "
        If b_cboClientName = True Then
            s_Query = s_Query & "WHERE CLIENT_NAME = '" & s_ClientName & "'"
            If b_txtClientName = True Then
                s_Query = s_Query & " OR CLIENT_NAME LIKE '" & XorTrim("" & Frm_Main.txt_ClientName.Text) & "%'"
            End If
        Else
            s_Query = s_Query & "WHERE CLIENT_NAME LIKE '" & s_ClientName & "%'"
        End If
        
        s_Query = s_Query & " ORDER BY CLIENT_NAME"
       
    
        RS1.Open s_Query, Conn1, adOpenForwardOnly, adLockReadOnly
        If RS1.EOF Then
            MsgBox "No Records found"
        End If
        'RS1.Open s_Query, Conn1
        'Set RS1 = Conn1.Execute(s_Query)
        While Not RS1.EOF
            Debug.Print "CLIENT NAME = " & CStr(RS1("CLIENT_NAME").Value) & " CLIENT_NUMBER = " & CStr(RS1("CLIENT_NUMBER").Value) & Chr(13) & Chr(10)
            MsgBox "" & RS1("CLIENT_NAME")
            RS1.MoveNext
        Wend
        RS1.Close
        Conn1.Close
        Set RS1 = Nothing
        Set Conn1 = Nothing

    End If
        
End Sub

Private Sub ExitProgram()
    Unload Frm_Main 'Unload Main form from memory
    End ' End the program
End Sub

Private Sub Form_Load()
Dim iresponse As Integer
    Dim s_Query As String
    Dim Conn1 As New ADODB.Connection
    Dim RS1 As New ADODB.Recordset
    Dim FileDatabase As String
    
    FileDatabase = App.Path & "\CAMELSDATA.mdb"

    
    'Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileDatabase
        
    Conn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileDatabase
    'Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:nwind.mdb;"
    
    ' Load Client Names
    s_Query = "SELECT CLIENT_NAME, CLIENT_NUMBER FROM TBL_CLIENT "
    s_Query = s_Query & "ORDER BY CLIENT_NAME"
    
    RS1.Open s_Query, Conn1, adOpenForwardOnly, adLockReadOnly
    If RS1.EOF Then
        MsgBox "No Records found"
    End If

    While Not RS1.EOF
        Frm_Main.cbo_ClientName.AddItem (RS1("CLIENT_NAME"))
        RS1.MoveNext
    Wend
    
    ' Load Client Numbers
    s_Query = "SELECT CLIENT_NAME, CLIENT_NUMBER FROM TBL_CLIENT "
    s_Query = s_Query & "ORDER BY CLIENT_NUMBER"
    
    RS1.Close
    
    RS1.Open s_Query, Conn1, adOpenForwardOnly, adLockReadOnly
    If RS1.EOF Then
        MsgBox "No Records found"
    End If

    While Not RS1.EOF
        Frm_Main.cbo_ClientNumber.AddItem (CStr(RS1("CLIENT_NUMBER")))
        RS1.MoveNext
    Wend
    
    ' Load Matter Names
    s_Query = "SELECT MATTER_NAME, MATTER_NUMBER FROM TBL_MATTER "
    s_Query = s_Query & "ORDER BY MATTER_NAME"
    
    RS1.Close
    
    RS1.Open s_Query, Conn1, adOpenForwardOnly, adLockReadOnly
    If RS1.EOF Then
        MsgBox "No Records found"
    End If

    While Not RS1.EOF
        Frm_Main.cbo_MatterName.AddItem (RS1("MATTER_NAME"))
        RS1.MoveNext
    Wend
    
    'Load Matter Numbers
    s_Query = "SELECT MATTER_NAME, MATTER_NUMBER FROM TBL_MATTER "
    s_Query = s_Query & "ORDER BY MATTER_NUMBER"
    
    RS1.Close
    
    RS1.Open s_Query, Conn1, adOpenForwardOnly, adLockReadOnly
    If RS1.EOF Then
        MsgBox "No Records found"
    End If

    While Not RS1.EOF
        Frm_Main.cbo_MatterNumber.AddItem (CStr(RS1("MATTER_NUMBER")))
        RS1.MoveNext
    Wend

    
    RS1.Close
    Conn1.Close
    Set RS1 = Nothing
    Set Conn1 = Nothing
End Sub

