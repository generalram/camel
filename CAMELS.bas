Attribute VB_Name = "CAMELS"
'Shared CAMELS Module between CAMEL systems.
Dim FileDatabase As String ' Global variable stores where the MS-Access file is located


Sub CMSearch()
Dim iresponse As Integer
    Dim s_Query As String
    Dim Conn1 As New ADODB.Connection
    Dim RS1 As New ADODB.Recordset
    
    Dim s_ClientName As String
    Dim s_ClientNumber As String
    Dim b_cboClientName As Boolean
    Dim b_txtClientName As Boolean
    
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

Sub InitCMSearch()
Dim iresponse As Integer
    Dim s_Query As String
    Dim Conn1 As New ADODB.Connection
    Dim RS1 As New ADODB.Recordset

    
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
