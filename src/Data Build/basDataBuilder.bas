Attribute VB_Name = "basDataBuilder"
Option Explicit
Private Const DATAFILE = "\Data\Dataset"

Private FilePath As String

Private conn As New ADODB.Connection

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal strSectionNameA As String, ByVal strKeyNameA As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal strSectionNameA As String, ByVal strKeyNameA As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Const OF_EXIST = &H4000
Public Type OFSTRUCT '136 bytes --- Data Structure for OpenFile call
    cBytes As String * 1
    fFixedDisk As String * 1
    nErrCode As Integer
    reserved As String * 4
    szPathName As String * 128
End Type

Public Sub WriteZipTo(strFile As String)
Dim NewZip As New cZip
Dim strFilePath As String

    strFilePath = App.Path & "\data\zipped\" & strFile & ".zip"
    
    If IsFile(strFilePath) Then
        Kill strFilePath
    End If
    
    NewZip.ZipFile = strFilePath
    
    NewZip.ClearFileSpecs
    
    NewZip.AddFileSpec App.Path & "\data\" & strFile & ".dat"
        
    NewZip.Zip
    
    Set NewZip = Nothing
    
    ' Rename the files
'    If IsFile(strFilePath & ".zip") Then
'        Name strFilePath & ".zip" As strFilePath
'    End If
    
'    ' Delete the zips
'    If IsFile(strFilePath & ".zip") Then
'    End If
    
End Sub

Public Sub WriteFile()
    Screen.MousePointer = vbHourglass
    WriteDataVersion
    WriteAttributes
    WriteRealms
    WriteRaceDetails
    WriteClassDetails
    WriteZipTo "DataSet"

    Progress "Done"
    frmDataBuilder.Caption = " Data Done"
    Beep
    Screen.MousePointer = vbDefault
End Sub

Private Sub Progress(msg As String)

    frmDataBuilder.lbl = msg
    frmDataBuilder.Refresh

End Sub

Public Sub Main()

    ' Set the file path
    FilePath = App.Path & DATAFILE
    
    ' Show the form
    frmDataBuilder.Show
    
End Sub

Public Sub OpenData()
Dim strConn As String

    strConn = "Driver={Microsoft Access Driver (*.mdb)};" & _
        "Dbq=" & App.Path & "\Data Build\DaocData.mdb;" & _
        "Uid=Admin;Pwd=;"

    'Open connection
    With conn
        .CursorLocation = adUseClient
        .ConnectionString = strConn
        .Open
    End With
    
End Sub

Public Sub CleanUp()

    conn.Close
    Set conn = Nothing
    
End Sub

Public Function IsFile(NameOfFile As String) As Boolean
'True  = File Exists
'False = File does not exist
 Dim result As Integer
 Dim response As OFSTRUCT
 'Dim OF_EXIST As Integer
 

result = OpenFile(NameOfFile, response, OF_EXIST)
 'Response.nErrCode can be used to determine the exact error if needed if the future.
 If result < 0 Then IsFile = False Else IsFile = True
 
End Function

Private Sub WriteDataVersion()
Dim intVersion As Integer

Progress "Writing data version"

    intVersion = Val(ProfileGetItem("General", "Data Version", "0", FilePath))
    
    If IsFile(FilePath & ".dat") Then
        Kill FilePath & ".dat"
    End If
    
    If intVersion = 0 Then
        SaveItem "General", "Data Version", 1032
    Else
        SaveItem "General", "Data Version", intVersion + 1
    End If
    
End Sub

Public Sub WriteRealms()
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

Progress "Writing realms"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryRealms"
        .CommandType = adCmdStoredProc
    End With
         
    Set rs = cmd.Execute
    With rs
'        .MoveFirst
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            SaveItem "Realms", "Realm " & i, StripReturns(.Fields(0))
            SaveItem "Realms", "Realm " & i & " Description", StripReturns(.Fields(1))
            WriteRaces StripReturns(.Fields(0))
            .MoveNext
        Loop
        On Error GoTo 0
        
        frmDataBuilder.lbl = .RecordCount
        
        .Close
    End With
    
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Public Sub WriteRaceDetails()
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

Progress "Writing race details"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryRaceDetails"
        .CommandType = adCmdStoredProc
    End With
         
    Set rs = cmd.Execute
    With rs
'        .MoveFirst
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            SaveItem StripReturns(.Fields(0)), "Strength", StripReturns(.Fields(1))
            SaveItem StripReturns(.Fields(0)), "Constitution", StripReturns(.Fields(2))
            SaveItem StripReturns(.Fields(0)), "Dexterity", StripReturns(.Fields(3))
            SaveItem StripReturns(.Fields(0)), "Quickness", StripReturns(.Fields(4))
            SaveItem StripReturns(.Fields(0)), "Intelligence", StripReturns(.Fields(5))
            SaveItem StripReturns(.Fields(0)), "Piety", StripReturns(.Fields(6))
            SaveItem StripReturns(.Fields(0)), "Empathy", StripReturns(.Fields(7))
            SaveItem StripReturns(.Fields(0)), "Charisma", StripReturns(.Fields(8))
            WriteAvailableClasses (.Fields(0))
            .MoveNext
        Loop
        On Error GoTo 0
        
        .Close
    End With
    
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub
Public Sub WriteClassDetails()
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strFilePath As String

Progress "Writing class details"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryClassDetails"
        .CommandType = adCmdStoredProc
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            strFilePath = App.Path & "/Data/" & .Fields(0)
            If IsFile(strFilePath & ".dat") Then
                Kill strFilePath & ".dat"
            End If
            
            ' If there is no base class, do not bother to write
            If Not IsNull((.Fields(1))) Then
                ProfileSaveItem StripReturns(.Fields(0)), "Base Class", StripReturns(.Fields(1)), strFilePath
                ProfileSaveItem StripReturns(.Fields(0)), "Primary Attribute", StripReturns(.Fields(2)), strFilePath
                ProfileSaveItem StripReturns(.Fields(0)), "Secondary Attribute", StripReturns(.Fields(3)), strFilePath
                ProfileSaveItem StripReturns(.Fields(0)), "Tertiary Attribute", StripReturns(.Fields(4)), strFilePath
                ProfileSaveItem StripReturns(.Fields(0)), "Multiplier", StripReturns(.Fields(5)), strFilePath
                ProfileSaveItem StripReturns(.Fields(0)), "Description", StripReturns(.Fields(6)), strFilePath
                WriteSpecBased .Fields(0), .Fields(1), strFilePath
                WriteLevelBased .Fields(0), .Fields(1), strFilePath
                WriteArmorAvailable .Fields(0), strFilePath
                WriteOtherAvailable .Fields(0), strFilePath
                WriteZipTo StripReturns(.Fields(0))
            End If
            .MoveNext
        Loop
        On Error GoTo 0
        
        .Close
    End With
    
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Public Sub WriteAttributes()
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

Progress "Writing attributes"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryAttributes"
        .CommandType = adCmdStoredProc
    End With
         
    Set rs = cmd.Execute
    With rs
'        .MoveFirst
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            SaveItem "Attributes", "Attribute " & i, StripReturns(.Fields(0))
            SaveItem "Attributes", "Attribute " & i & " Description", StripReturns(.Fields(1))
            .MoveNext
        Loop
        On Error GoTo 0
        
        .Close
    End With
    
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub
Public Sub WriteLineDetails(strLineA As String, strFilePathA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter

    Progress "Writing line details"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryLineDetails"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("LineA", adVarChar, adParamInput, Len(strLineA), strLineA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
'        .MoveFirst
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem StripReturns(.Fields(0)), "Description", StripReturns(.Fields(1)), strFilePathA
            WriteStyles StripReturns(.Fields(0)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
        
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Private Sub WriteAvailableClasses(strRaceA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter

Progress "Writing " & strRaceA & " classes"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryAvailableClasses"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("RaceA", adVarChar, adParamInput, Len(strRaceA), strRaceA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            SaveItem strRaceA, "Class " & i, StripReturns(.Fields(0))
            .MoveNext
        Loop
        On Error GoTo 0
                
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Private Sub WriteStyles(strLineA As String, strFilePathA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter
Dim col As New Collection
Dim count As Integer

Progress "Writing " & strLineA & " styles"


    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryStyles"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("LineA", adVarChar, adParamInput, Len(strLineA), strLineA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        .MoveFirst
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            ' Problem with duplicate styles so as a temp fix
            ' Add to col to ensure unique, skip if lvl and name are the same
'            On Error GoTo SkipToHere
            count = col.count
            col.Add StripReturns(.Fields(0)), StripReturns(.Fields(0)) & StripReturns(.Fields(1))
'            On Error Resume Next
            If col.count > count Then
                i = i + 1
                ProfileSaveItem strLineA, "Style " & i, StripReturns(.Fields(0)), strFilePathA
                ProfileSaveItem strLineA, "Style " & i & " Level", StripReturns(.Fields(1)), strFilePathA
                ProfileSaveItem strLineA, "Style " & i & " Description", StripReturns(.Fields(2)), strFilePathA
            End If
'SkipToHere:
            .MoveNext
        Loop
        On Error GoTo 0
                
        .Close
    End With
    
    Set Param = Nothing
    Set col = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Private Sub WriteSpecBased(strClassA As String, strBaseClassA As String, strFilePathA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter

Progress "Writing " & strClassA & " spec lines"

    ' Write base class spec lines
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qrySpecLinesAvailable"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("ClassA", adVarChar, adParamInput, Len(strBaseClassA), strBaseClassA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem strClassA, "Specialisation " & i, StripReturns(.Fields(0)), strFilePathA
            ProfileSaveItem strClassA, "Specialisation " & i & " Level", StripReturns(.Fields(1)), strFilePathA
            ProfileSaveItem strClassA, "Specialisation " & i & " Cap", StripReturns(.Fields(2)), strFilePathA
            WriteLineDetails StripReturns(.Fields(0)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
                
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
    ' Write spec lines
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qrySpecLinesAvailable"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("ClassA", adVarChar, adParamInput, Len(strClassA), strClassA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem strClassA, "Specialisation " & i, StripReturns(.Fields(0)), strFilePathA
            ProfileSaveItem strClassA, "Specialisation " & i & " Level", StripReturns(.Fields(1)), strFilePathA
            ProfileSaveItem strClassA, "Specialisation " & i & " Cap", StripReturns(.Fields(2)), strFilePathA
            WriteLineDetails StripReturns(.Fields(0)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
                
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Private Sub WriteLevelBased(strClassA As String, strBaseClassA As String, strFilePathA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter

Progress "Writing " & strClassA & " level based lines"

    ' Write the base class lines first
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryLevelLinesAvailable"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("ClassA", adVarChar, adParamInput, Len(strBaseClassA), strBaseClassA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem strClassA, "Line " & i, StripReturns(.Fields(0)), strFilePathA
            ProfileSaveItem strClassA, "Line " & i & " Level", StripReturns(.Fields(1)), strFilePathA
            WriteLineDetails StripReturns(.Fields(0)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
                
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
    ' Now write the advanced class lines (i continues)
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryLevelLinesAvailable"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("ClassA", adVarChar, adParamInput, Len(strClassA), strClassA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem strClassA, "Line " & i, StripReturns(.Fields(0)), strFilePathA
            ProfileSaveItem strClassA, "Line " & i & " Level", StripReturns(.Fields(1)), strFilePathA
            WriteLineDetails StripReturns(.Fields(0)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
                
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub
Private Sub WriteArmorAvailable(strClassA As String, strFilePathA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter

Progress "Writing " & strClassA & " armor"
    
    ' Write the Armor
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryArmorAvailable"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("ClassA", adVarChar, adParamInput, Len(strClassA), strClassA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem strClassA, "Armor " & i, StripReturns(.Fields(0)), strFilePathA
            ProfileSaveItem strClassA, "Armor " & i & " Level", StripReturns(.Fields(1)), strFilePathA
            WriteArmorDetails StripReturns(.Fields(0)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
                
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Private Sub WriteOtherAvailable(strClassA As String, strFilePathA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter
 
Progress "Writing " & strClassA & " other lines"
   
    ' Write the Armor
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryOtherLinesAvailable"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("ClassA", adVarChar, adParamInput, Len(strClassA), strClassA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem strClassA, "Other " & i, StripReturns(.Fields(0)), strFilePathA
            ProfileSaveItem strClassA, "Other " & i & " Level", StripReturns(.Fields(1)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
                
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Public Sub WriteRaces(strRealmA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter

Progress "Writing " & strRealmA & " races"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryRaces"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("RealmA", adVarChar, adParamInput, Len(strRealmA), strRealmA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            SaveItem strRealmA & " Races", "Race " & i, StripReturns(.Fields(0))
            SaveItem strRealmA & " Races", "Race " & i & " Description", StripReturns(.Fields(1))
            .MoveNext
        Loop
        On Error GoTo 0
        
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Public Sub WriteArmorDetails(strArmorA As String, strFilePathA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter

Progress "Writing armor descriptions"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryArmor"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("ArmorA", adVarChar, adParamInput, Len(strArmorA), strArmorA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem "Armor", StripReturns(.Fields(0)) & " Description", StripReturns(.Fields(1)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
        
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub
Public Sub WriteOtherDetails(strOtherA As String, strFilePathA As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Param As ADODB.Parameter
 
Progress "Writing other lines details"

    With cmd
        Set .ActiveConnection = conn
        .CommandText = "qryOther"
        .CommandType = adCmdStoredProc
        Set Param = .CreateParameter("OtherA", adVarChar, adParamInput, Len(strOtherA), strOtherA)
        .Parameters.Append Param
    End With
         
    Set rs = cmd.Execute
    With rs
        On Error Resume Next ' ignore blanks
        Do While Not (.EOF)
            i = i + 1
            ProfileSaveItem "Other", StripReturns(.Fields(0)) & " Description", StripReturns(.Fields(1)), strFilePathA
            .MoveNext
        Loop
        On Error GoTo 0
        
        .Close
    End With
    
    Set Param = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
End Sub

Public Function StripReturns(strTextA As String) As String
    StripReturns = Replace(strTextA, vbCrLf, " ")
End Function

Public Function ProfileGetItem(strSectionNameA As String, strKeyNameA As String, strDefaultValueA As String, strInifileA As String) As String
        
   Dim success As Long
   Dim nSize As Long
   Dim ret As String
  
  ' Pad a string large enough to hold the data.
   ret = Space$(2048)
   nSize = Len(ret)
   success = GetPrivateProfileString(strSectionNameA, strKeyNameA, strDefaultValueA, ret, nSize, strInifileA & ".dat")
   
   If success Then
      ProfileGetItem = Left$(ret, success)
   End If
   
End Function

Public Sub ProfileDeleteItem(strSectionNameA As String, strKeyNameA As String, strInifileA As String)
   
   Call WritePrivateProfileString(strSectionNameA, strKeyNameA, vbNullString, strInifileA & ".dat")

End Sub

Public Sub ProfileDeleteSection(strSectionNameA As String, strInifileA As String)
   
   Call WritePrivateProfileString(strSectionNameA, vbNullString, vbNullString, strInifileA & ".dat")

End Sub

Public Sub SaveItem(strSectionNameA As String, strKeyNameA As String, strValueA As String)
   
   Call WritePrivateProfileString(strSectionNameA, strKeyNameA, strValueA, FilePath & ".dat")

End Sub

Public Sub ProfileSaveItem(strSectionNameA As String, strKeyNameA As String, strValueA As String, strInifileA As String)
   
   Call WritePrivateProfileString(strSectionNameA, strKeyNameA, strValueA, strInifileA & ".dat")

End Sub
