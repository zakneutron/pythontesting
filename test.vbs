'test.vbs
' VBScript program to document the organization specified by the
' manager and directReports attributes in Active Directory.
'
' ----------------------------------------------------------------------
' Copyright (c) 2008 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - April 30, 2008
' Version 1.1 - July 12, 2008 - Include contacts.
' Version 1.2 - May 29, 2010 - Include all object classes.
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

Option Explicit

Dim objRootDSE, strDNSDomain, adoCommand, adoConnection
Dim strBase, strAttributes

' Determine DNS domain name.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use ADO to search Active Directory.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
Set adoCommand.ActiveConnection = adoConnection
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

' Search entire domain.
strBase = "<LDAP://" & strDNSDomain & ">"

' Comma delimited list of attribute values to retrieve.
strAttributes = "distinguishedName,directReports"

' Document organization.
Call Reports("Top", "")

' Clean up.
adoConnection.Close

Sub Reports(ByVal strManager, ByVal strOffset)
    ' Recursive subroutine to document organization.

    Dim strDN, adoRecordset, arrReports, strReport
    Dim strFilter, strQuery

    If (strManager = "Top") Then
        ' Search for all managers at top of organizational tree.
        ' These are objects with direct reports but no manager.
        strFilter = "(&(!manager=*)(directReports=*))"
    Else
        ' Output object that reports to previous manager.
        Wscript.Echo strOffset & strManager
        strOffset = strOffset & "    "
        ' Search for all objects that report to this manager.
        strFilter = "(manager=" & strManager & ")"
    End If

    ' Construct the LDAP query.
    strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

    ' Run the query.
    adoCommand.CommandText = strQuery
    Set adoRecordset = adoCommand.Execute

    ' Enumerate the resulting recordset.
    Do Until adoRecordset.EOF
        ' Retrieve values.
        strDN = adoRecordset.Fields("distinguishedName").Value
        Wscript.Echo strOffset & strDN
        arrReports = adoRecordset.Fields("directReports").Value
        If Not IsNull(arrReports) Then
            For Each strReport In arrReports
                Call Reports(strReport, strOffset & "    ")
            Next
        End If
        adoRecordset.MoveNext
    Loop
    adoRecordset.Close

End Sub