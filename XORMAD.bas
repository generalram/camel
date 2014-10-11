Attribute VB_Name = "XORMAD"
'XORMAD General functions and subroutines
Function jnow() As Variant
     'Documented at http://www.vbcode.com/Asp/showsn.asp?theID=293
     'By Norman King, 07/31/96
     'This function takes the current date and returns it as a Julian date.
     jnow = Right$(Format(Date, "yy"), 1) & Format$(Date, "y")
End Function

Function jdate(cdate1 As Variant) As Variant
     'Documented at http://www.vbcode.com/Asp/showsn.asp?theID=294
     ' By Norman King, 07/31/96
     'This function takes a date and converts it to the Julian format. 'Hence the jdate name.
     If Trim(cdate1 & "") = "" Then
        jdate = Right$(Format(Date, "yy"), 1) & Format$(Date, "y")
        Else: jdate = Right$(Format(cdate1, "yy"), 1) & Format$(cdate1, "y")
     End If
End Function

Function XorTrim(vdata As Variant) As Variant
    ' By Norman King 12/27/2007
    ' This function takes data and trims off the spaces,
    ' and returns a blank string "" if NULL or EMPTY
    ' to avoid runtime errors like invalid use of null
    XorTrim = Trim("" & vdata)
End Function

Function SQLFilter(sqlstring As String) As String
    ' By Norman King 10/10/2014
    ' This function takes a string and tries to filter it to prevent
    ' SQL injections
    Dim s_filtered As String
    
    
    s_filtered = sqlstring
    
    s_filtered = Replace(s_filtered, "'", "''")
    s_filtered = Replace(s_filtered, ";", "';'")
    s_filtered = Replace(s_filtered, ":", "':'")
    s_filtered = Replace(s_filtered, Chr(34), Chr34 & Chr(34))

    SQLFilter = s_filtered
End Function

Function HTMLFilter(sqlstring As String) As String
    ' By Norman King 10/10/2014
    ' This function takes a string and tries to filter it to prevent
    ' HTML attacks
    Dim s_filtered As String
    
    
    s_filtered = sqlstring
    
    s_filtered = Replace(s_filtered, "&", "&amp;")
    s_filtered = Replace(s_filtered, "<", "&lt;")
    s_filtered = Replace(s_filtered, ">", "&gt;")
    s_filtered = Replace(s_filtered, "'", "&quot;")
    
    HTMLFilter = s_filtered
End Function
