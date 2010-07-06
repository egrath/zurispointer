VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCharName As String

Private mvarDirtyBit As Boolean

Private mvarName As String

Private mvarRealm As String
Private mvarRace As String
Private mvarClass As String

Private mvarAttributePrimary As String
Private mvarAttributeSecondary As String
Private mvarAttributeTertiary As String

Private mvarMultiplier As Single

Private mvarSpecLines As Collection
Private mvarNonSpecLines As Collection
Private mvarArmorLines As clsLine
Private mvarOtherLines As clsLine
Private mvarSpecByPercent As Boolean

Private mvarLevel As Double

Private mvarBaseAttributes As Collection
Private mvarBuffedAttributes As Collection
Private mvarAdditionalAttributes As Collection

' Spend of spec points
Private mvarLinePercent As Collection
Private mvarLineLevel As Collection
Private mvarLinePoints As Collection

' Used to convert points to level and vis a versa(???)
Private mvarPointsToLevel As Collection
Private mvarLevelToPoints As Collection
Private mvarLevelToPointsEarnt1 As Collection 'different to allow for the half lvls for 1 mult
Private mvarLevelToPointsEarnt15 As Collection 'different to allow for the half lvls for 1.5 mult
Private mvarLevelToPointsEarnt2 As Collection 'different to allow for the half lvls for 2 mult
Private mvarLevelToPointsEarnt22 As Collection 'different to allow for the half lvls for 2.2 mult
Private mvarLevelToPointsEarnt25 As Collection 'different to allow for the half lvls for 2.5 mult

Public Property Let LinePercent(index As Integer, intValue As Integer)
Dim strLine As String
Dim intPoints As Integer
Dim intLevel As Integer

    ' Get the name of the line we are adjusting
    strLine = mvarSpecLines(index).LineName
    
    ' Check the level cap for the skill
    intLevel = Int(mvarLevel / 100 * intValue)
    If intLevel > -Int(-mvarLevel / LineCap(index)) Then
        MsgBox "Sorry, you have reached your level cap"
        LineLevel(index) = -Int(-mvarLevel / LineCap(index))
        ' Recalc percent
        mvarLinePercent.Remove strLine
        mvarLinePercent.Add Int(mvarLineLevel(strLine) / mvarLevel * 100), strLine
        Exit Property
    End If
    
    ' Check if we have enough points
    intPoints = cLevelToPoints(intLevel)
    If PointsAvailable + mvarLinePoints(strLine) - intPoints < 0 Then
        MsgBox "Sorry, not enough points available"
        LinePoints(index) = PointsAvailable + mvarLinePoints(strLine)
        Exit Property
    End If
    
    ' Remove that line from the three collections
    On Error Resume Next
    mvarLinePercent.Remove (strLine)
    If mvarSpecByPercent Then mvarLineLevel.Remove (strLine)
    If mvarSpecByPercent Then mvarLinePoints.Remove (strLine)
    On Error GoTo 0
    
    ' Set the spend on the line as a percent of current level
    mvarLinePercent.Add intValue, strLine
    ' Set the level based on the percent
    If mvarSpecByPercent Then mvarLineLevel.Add intLevel, strLine
    ' Set the point spend based on the level
    If mvarSpecByPercent Then mvarLinePoints.Add intPoints, strLine
    
End Property

Public Sub RecalculateSpecLines()
Dim i As Integer

    If mvarSpecByPercent Then
        For i = 1 To mvarSpecLines.Count
            LineLevel(i) = Int(mvarLevel / 100 * LinePercent(i))
'            mvarLineLevel(i) = Int(mvarLevel / 100 * mvarLinePercent(i))
'            mvarLinePoints(i) = cLevelToPoints(mvarLineLevel(i))
        Next i
    Else
        For i = 1 To mvarSpecLines.Count
            LinePercent(i) = Int(100 / mvarLevel * LineLevel(i))
'            mvarLinePercent(i) = Int(100 / mvarLevel * mvarLineLevel(i))
        Next i
    End If
    
End Sub

Public Property Get LinePercent(index As Integer) As Integer
Dim strLine As String

    ' Get the name of the line we are adjusting
    strLine = mvarSpecLines(index).LineName
    LinePercent = mvarLinePercent(strLine)

End Property

Public Property Let LineLevel(index As Integer, intValue As Integer)
Dim strLine As String
Dim intPoints As Integer
Dim intLevel As Integer

    ' Get the name of the line we are adjusting
    strLine = mvarSpecLines(index).LineName
    
    ' Check if we have enough points
    intPoints = cLevelToPoints(intValue)
    If PointsAvailable + mvarLinePoints(strLine) - intPoints < 0 Then
        MsgBox "Sorry, not enough points available"
        LinePoints(index) = PointsAvailable + mvarLinePoints(strLine)
        Exit Property
    End If
    
    ' Check the level cap for the skill
    intLevel = intValue
    If intLevel > -Int(-mvarLevel / LineCap(index)) Then
        MsgBox "Sorry, you have reached your level cap"
        LineLevel(index) = -Int(-mvarLevel / LineCap(index))
        Exit Property
    End If
        
    ' Remove that line from the three collections
    On Error Resume Next
    If Not mvarSpecByPercent Then mvarLinePercent.Remove (strLine)
    mvarLineLevel.Remove (strLine)
    mvarLinePoints.Remove (strLine)
    On Error GoTo 0
    
    ' Set the spend on the line as level
    mvarLineLevel.Add intLevel, strLine
    ' Set the point spend based on the level
    mvarLinePoints.Add intPoints, strLine
    ' Set the percent based on the level to current level
    If Not mvarSpecByPercent Then mvarLinePercent.Add Int(intLevel / mvarLevel * 100), strLine
    
End Property

Public Property Get LineLevel(index As Integer) As Integer
Dim strLine As String

    ' Get the name of the line we are adjusting
    strLine = mvarSpecLines(index).LineName
    LineLevel = mvarLineLevel(strLine)

End Property

Public Function LineCap(index) As Single
    LineCap = mvarSpecLines(index).Cap
End Function

Public Property Let LinePoints(index As Integer, intValue As Integer)
Dim strLine As String
Dim intPoints As Integer
Dim intLevel As Integer

    SetDirty
    ' Get the name of the line we are adjusting
    strLine = mvarSpecLines(index).LineName
    
    ' Check the level cap for the skill
    intLevel = cPointsToLevel(intValue)
    If intLevel > -Int(-mvarLevel / LineCap(index)) Then
        MsgBox "Sorry, you have reached your level cap"
        LineLevel(index) = -Int(-mvarLevel / LineCap(index))
        Exit Property
    End If
    
    ' Get the exact number of points we require and check if we have enough points
    intPoints = cLevelToPoints(intLevel)
    If PointsAvailable + mvarLinePoints(strLine) - intPoints < 0 Then
        MsgBox "Sorry, not enough points available"
        LinePoints(index) = PointsAvailable + mvarLinePoints(strLine)
        Exit Property
    End If
    
    
    ' Remove that line from the three collections
    On Error Resume Next
    mvarLinePercent.Remove (strLine)
    mvarLineLevel.Remove (strLine)
    mvarLinePoints.Remove (strLine)
    On Error GoTo 0
    
    ' Set the spend on the line as points
    mvarLinePoints.Add intPoints, strLine
    ' Set the level based on the point spend
    mvarLineLevel.Add intLevel, strLine
    ' Set the percent based on the level to current level
    mvarLinePercent.Add Int(intLevel / mvarLevel * 100), strLine

End Property

Public Property Get LinePoints(index As Integer) As Integer
Dim strLine As String

    ' Get the name of the line we are adjusting
    strLine = mvarSpecLines(index).LineName
    LinePoints = mvarLinePoints(strLine)
    
End Property

Private Sub ResetToonData()
' Reset the char data to blank
    
    Set mvarLinePercent = New Collection
    Set mvarLineLevel = New Collection
    Set mvarLinePoints = New Collection
    mvarLevel = 5
    
End Sub

Private Sub ResetClassData()

    mvarClass = "Class"

    ' Reset the data specific to class
    mvarAttributePrimary = vbNullString
    mvarAttributeSecondary = vbNullString
    mvarAttributeTertiary = vbNullString
    mvarMultiplier = 1
'    Set mvarLinePercent = New Collection
'    Set mvarLineLevel = New Collection
'    Set mvarLinePoints = New Collection
    Set mvarSpecLines = New Collection
    Set mvarNonSpecLines = New Collection
    Set mvarArmorLines = New clsLine
    Set mvarOtherLines = New clsLine
    
    ResetToonData
    
End Sub

Private Sub ResetRaceData()
' Reset the char data to blank

    mvarRace = "Race"
    
    Set mvarBaseAttributes = New Collection
    Set mvarBuffedAttributes = New Collection
    Set mvarAdditionalAttributes = New Collection
    
    ResetClassData
    
End Sub

Public Function DataVersion() As String
   
    DataVersion = ProfileGetItem("General", "Data Version", "unknown", App.Path & DATAFILE)
   
End Function

Private Sub ResetRealmData()
' Reset the char data to blank
    
    ' Currently nothing specific to reset
    mvarRealm = "Realm"
'    mvarRace = "Race"
    ResetRaceData
    
End Sub

Public Sub NewToon()
    ResetRealmData
    mvarName = "No Name"
End Sub

Private Sub LoadClassData(strClass As String)
' Load the data for a specific class
Dim i As Integer
Dim j As Integer
Dim Style As clsStyle
Dim strSpecLine As String
Dim strNonSpecLine As String
Dim strStyle As String
Dim intLevel As Integer
Dim strDescription As String
Dim varLine As clsLine
Dim strFilePath As String

    ' Reset the class data
    ResetClassData
    mvarClass = strClass
    
    strFilePath = App.Path & "\Data\" & mvarClass & ".dat"
    
    ' Populate the class with data from the relevant ini file
    mvarMultiplier = Val(ProfileGetItem(mvarClass, "Multiplier", "0", strFilePath))
    mvarAttributePrimary = ProfileGetItem(mvarClass, "Primary Attribute", "unknown", strFilePath)
    mvarAttributeSecondary = ProfileGetItem(mvarClass, "Secondary Attribute", "unknown", strFilePath)
    mvarAttributeTertiary = ProfileGetItem(mvarClass, "Tertiary Attribute", "unknown", strFilePath)

    
    ' Get all the spec lines for the class
    i = 0
    strSpecLine = vbNullString
    Do While strSpecLine <> "unknown"
        i = i + 1
        ' Get the classes first/next spec line
        strSpecLine = ProfileGetItem(mvarClass, "Specialisation " & i, "unknown", strFilePath)
        ' Bail if there isn't one
        If strSpecLine = "unknown" Then Exit Do
        ' Otherwise start a new line collection
        Set varLine = New clsLine
        ' Name the line
        varLine.LineName = strSpecLine
        varLine.Level = ProfileGetItem(mvarClass, "Specialisation " & i & " Level", "1", strFilePath)
        varLine.Cap = ProfileGetItem(mvarClass, "Specialisation " & i & " Cap", "1", strFilePath)
        varLine.Description = ProfileGetItem(strSpecLine, "Description", "No description", strFilePath)
        ' Get all the styles for the spec line
        j = 0
        strStyle = vbNullString
        Do While strStyle <> "unknown"
            j = j + 1
            ' Get the spec lines first/next style
            strStyle = ProfileGetItem(strSpecLine, "Style " & j, "unknown", strFilePath)
            ' Bail if there isn't one
            If strStyle = "unknown" Then
                Exit Do
            End If
            ' Otherwise get the other style properties
            intLevel = Val(ProfileGetItem(strSpecLine, "Style " & j & " Level", "unknown", strFilePath))
            strDescription = strStyle & " (" & strSpecLine & ")" & vbCrLf _
            & "Gained at level " & intLevel & vbCrLf _
            & ProfileGetItem(strSpecLine, "Style " & j & " Description", "unknown", strFilePath)
            ' Add the style to the line
            varLine.Add strStyle, intLevel, strDescription, strStyle
        Loop
            
        ' Now add the line collection to the spec lines collection
        mvarSpecLines.Add varLine, strSpecLine
        ' Add to collections containing percent. points, and level in line
        mvarLinePercent.Add 20, strSpecLine
        mvarLineLevel.Add 1, strSpecLine
        mvarLinePoints.Add 0, strSpecLine
        
    Loop
    RecalculateSpecLines
    
    ' Get all the non spec lines for the class
    i = 0
    strNonSpecLine = vbNullString
    Do While strNonSpecLine <> "unknown"
        i = i + 1
        ' Get the classes first/next spec line
        strNonSpecLine = ProfileGetItem(mvarClass, "Line " & i, "unknown", strFilePath)
        ' Bail if there isn't one
        If strNonSpecLine = "unknown" Then Exit Do
        ' Otherwise start a new line collection
        Set varLine = New clsLine
        ' Name the line
        varLine.LineName = strNonSpecLine
        ' Get all the styles for the spec line
        j = 0
        strStyle = vbNullString
        Do While strStyle <> "unknown"
            j = j + 1
            ' Get the spec lines first/next style
            strStyle = ProfileGetItem(strNonSpecLine, "Style " & j, "unknown", strFilePath)
            ' Bail if there isn't one
            If strStyle = "unknown" Then Exit Do
            ' Otherwise get the other style properties
            intLevel = Val(ProfileGetItem(strNonSpecLine, "Style " & j & " Level", "unknown", strFilePath))
            strDescription = strStyle & " (" & strNonSpecLine & ")" & vbCrLf _
            & "Gained at level " & intLevel & vbCrLf _
            & ProfileGetItem(strNonSpecLine, "Style " & j & " Description", "unknown", strFilePath)
            ' Add the style to the line
            varLine.Add strStyle, intLevel, strDescription, strStyle
        Loop
            
        ' Now add the line collection to the lines collection
        mvarNonSpecLines.Add varLine, strNonSpecLine
        
    Loop

    ' Get all the armor for the class
        j = 0
        strStyle = vbNullString
        Do While strStyle <> "unknown"
            j = j + 1
            ' Get the  first/next armor
            strStyle = ProfileGetItem(mvarClass, "Armor " & j, "unknown", strFilePath)
            ' Bail if there isn't one
            If strStyle = "unknown" Then Exit Do
            ' Otherwise get the other style properties
            intLevel = Val(ProfileGetItem(mvarClass, "Armor " & j & " Level", "unknown", strFilePath))
            strDescription = strStyle & " (Armor)" & vbCrLf _
            & "Gained at level " & intLevel & vbCrLf _
            & ProfileGetItem("Armor", strStyle & " Description", "unknown", strFilePath)
            ' Add the style to the line
            mvarArmorLines.Add strStyle, intLevel, strDescription, strStyle
        Loop
        
    ' Get all the other lines for the class
        j = 0
        strStyle = vbNullString
        Do While strStyle <> "unknown"
            j = j + 1
            ' Get the  first/next armor
            strStyle = ProfileGetItem(mvarClass, "Other " & j, "unknown", strFilePath)
            ' Bail if there isn't one
            If strStyle = "unknown" Then Exit Do
            ' Otherwise get the other style properties
            intLevel = Val(ProfileGetItem(mvarClass, "Other " & j & " Level", "unknown", strFilePath))
            strDescription = strStyle & " (Other)" & vbCrLf _
            & "Gained at level " & intLevel & vbCrLf _
            & ProfileGetItem("Other", strStyle & " Description", "unknown", strFilePath)
            ' Add the style to the line
            mvarOtherLines.Add strStyle, intLevel, strDescription, strStyle
        Loop
            
    ' Clean up
    Set varLine = Nothing

End Sub

Private Sub LoadRaceData(strRace As String)
Dim i As Integer

    ' Reset the race specific data
    ResetRaceData
    mvarRace = strRace
    
    'Populate the base attributes collection
    For i = 1 To UBound(AttributeArray)
        mvarBaseAttributes.Add ProfileGetItem(mvarRace, AttributeArray(i), "unknown", App.Path & DATAFILE), AttributeArray(i)
    Next i
    
End Sub

Private Sub LoadRealmData(strRealm As String)

    ' Reset the realm specific data
    ResetRealmData
    mvarRealm = strRealm
    ' Load the realm description
    
End Sub

Public Function SpecLine(index As Integer) As clsLine
    Set SpecLine = mvarSpecLines(index)
End Function

Public Function SpecLineCount() As String
    SpecLineCount = mvarSpecLines.Count
End Function

Public Function NonSpecLine(index As Integer) As clsLine
    Set NonSpecLine = mvarNonSpecLines(index)
End Function

Public Function NonSpecLineCount() As String
    NonSpecLineCount = mvarNonSpecLines.Count
End Function

Public Function ArmorLine() As clsLine
    Set ArmorLine = mvarArmorLines
End Function

Public Function ArmorLineCount() As String
    ArmorLineCount = mvarArmorLines.Count
End Function

Public Function OtherLine() As clsLine
    Set OtherLine = mvarOtherLines
End Function

Public Function OtherLineCount() As String
    OtherLineCount = mvarOtherLines.Count
End Function

Public Property Let Race(ByVal vData As String)
    If vData = vbNullString Then Exit Property
    If vData = mvarRace Then Exit Property
    SetDirty
    If vData = "Race" Then Exit Property
    LoadRaceData vData
End Property

Public Property Get SpecByPercent() As Boolean
    SpecByPercent = mvarSpecByPercent
End Property

Public Property Let SpecByPercent(ByVal vData As Boolean)
    SetDirty
    mvarSpecByPercent = vData
End Property

Public Property Get Race() As String
    Race = mvarRace
End Property

Public Property Let Realm(ByVal vData As String)
    If vData = vbNullString Then Exit Property
    If vData = mvarRealm Then Exit Property
    SetDirty
    If vData = "Realm" Then Exit Property
    LoadRealmData vData
    
End Property

Public Property Get Realm() As String
    Realm = mvarRealm
End Property

Public Property Let Class(ByVal vData As String)
    If vData = vbNullString Then Exit Property
    If mvarClass = vData Then Exit Property
    SetDirty
    If mvarClass = vData Then Exit Property
    If vData = "Class" Then Exit Property
    LoadClassData vData
End Property

Public Property Get Class() As String
    Class = mvarClass
End Property

Public Property Let CharName(ByVal vData As String)
    If vData = vbNullString Then Exit Property
    If vData = mvarName Then Exit Property
    mvarName = vData
End Property

Public Property Get CharName() As String
    CharName = mvarName
End Property

Public Function AttributePrimary() As String
    AttributePrimary = mvarAttributePrimary
End Function

Public Function AttributeSecondary() As String
    AttributeSecondary = mvarAttributeSecondary
End Function

Public Function AttributeTertiary() As String
    AttributeTertiary = mvarAttributeTertiary
End Function

Public Function Multiplier() As Single
    Multiplier = mvarMultiplier
End Function

Public Property Let Level(ByVal vData As Double)

    SetDirty
    If vData < 5 Then
        mvarLevel = 5
    ElseIf vData > MAXLEVEL Then
        mvarLevel = MAXLEVEL
    ElseIf (vData <= 40) Or (vData = Int(vData)) Then ' ensure that the level is an integer
        mvarLevel = vData
    ElseIf (vData - (1 / 2)) = Int(vData) Then 'or lvl + .5 if between 40 and 50
        mvarLevel = vData
    End If
    
End Property

Public Property Get Level() As Double
    Level = mvarLevel
End Property

Private Sub SetCollections()
' Populate the module lvl collections to use in functions
Dim intLevel As Integer
Dim intPoints As Integer
Dim intPointsEarnt1 As Integer
Dim intPointsEarnt15 As Integer
Dim intPointsEarnt2 As Integer
Dim intPointsEarnt22 As Integer
Dim intPointsEarnt25 As Integer
    ' Initialise the collections
    Set mvarPointsToLevel = New Collection
    Set mvarLevelToPoints = New Collection
    Set mvarLevelToPointsEarnt1 = New Collection
    Set mvarLevelToPointsEarnt15 = New Collection
    Set mvarLevelToPointsEarnt2 = New Collection
    Set mvarLevelToPointsEarnt22 = New Collection
    Set mvarLevelToPointsEarnt25 = New Collection
    
    For intLevel = 2 To MAXLEVEL
        ' Calculate total points for this level
        intPoints = intPoints + intLevel
        intPointsEarnt1 = intPointsEarnt1 + intLevel
        If intLevel > 5 Then
            intPointsEarnt15 = intPointsEarnt15 + Int(1.5 * intLevel)
            intPointsEarnt2 = intPointsEarnt2 + (2 * intLevel)
            intPointsEarnt22 = intPointsEarnt22 + Int(2.2 * intLevel)
            intPointsEarnt25 = intPointsEarnt25 + Int(2.5 * intLevel)
        Else
            intPointsEarnt15 = intPointsEarnt15 + intLevel
            intPointsEarnt2 = intPointsEarnt2 + intLevel
            intPointsEarnt22 = intPointsEarnt22 + intLevel
            intPointsEarnt25 = intPointsEarnt25 + intLevel
        End If
        
        ' Store in both collections for lookup later
        mvarPointsToLevel.Add intLevel, CStr(intPoints)
        mvarLevelToPoints.Add intPoints, CStr(intLevel)
        mvarLevelToPointsEarnt1.Add intPointsEarnt1, CStr(intLevel)
        mvarLevelToPointsEarnt15.Add intPointsEarnt15, CStr(intLevel)
        mvarLevelToPointsEarnt2.Add intPointsEarnt2, CStr(intLevel)
        mvarLevelToPointsEarnt22.Add intPointsEarnt22, CStr(intLevel)
        mvarLevelToPointsEarnt25.Add intPointsEarnt25, CStr(intLevel)
        ' If we are between 40 and 50 add a half level
        If (intLevel >= 40) And (intLevel < MAXLEVEL) Then
            ' add the additional points (half the level rouded down)
            intPointsEarnt1 = intPointsEarnt1 + Int(intLevel / 2)
            intPointsEarnt15 = intPointsEarnt15 + Int(1.5 * intLevel / 2)
            intPointsEarnt2 = intPointsEarnt2 + Int(intLevel)
            intPointsEarnt22 = intPointsEarnt22 + Int(2.2 * intLevel / 2)
            intPointsEarnt25 = intPointsEarnt25 + Int(2.5 * intLevel / 2)
            ' Store in collection
            mvarLevelToPointsEarnt1.Add intPointsEarnt1, CStr(intLevel + (1 / 2))
            mvarLevelToPointsEarnt15.Add intPointsEarnt15, CStr(intLevel + (1 / 2))
            mvarLevelToPointsEarnt2.Add intPointsEarnt2, CStr(intLevel + (1 / 2))
            mvarLevelToPointsEarnt22.Add intPointsEarnt22, CStr(intLevel + (1 / 2))
            mvarLevelToPointsEarnt25.Add intPointsEarnt25, CStr(intLevel + (1 / 2))
        End If
    Next intLevel
        
End Sub

Public Function PointsAcquired() As Integer
    If mvarLevel <= 1 Then
        PointsAcquired = 0
    Else
        Select Case mvarMultiplier
            Case 2.5
                PointsAcquired = mvarLevelToPointsEarnt25(CStr(mvarLevel))
            Case 2.2
                PointsAcquired = mvarLevelToPointsEarnt22(CStr(mvarLevel))
            Case 2
                PointsAcquired = mvarLevelToPointsEarnt2(CStr(mvarLevel))
            Case 1.5
                PointsAcquired = mvarLevelToPointsEarnt15(CStr(mvarLevel))
            Case Else ' should be 1 but...
                PointsAcquired = mvarLevelToPointsEarnt1(CStr(mvarLevel))
        End Select
        
        ' PointsAcquired = (mvarLevelToPointsEarnt(CStr(mvarLevel)) * mvarMultiplier) - ((mvarMultiplier - 1) * mvarLevelToPointsEarnt(CStr(5)))
    End If
    
End Function

Public Function PointsSpent() As Integer
Dim i As Integer
Dim intTotal As Integer

    For i = 1 To SpecLineCount
        intTotal = intTotal + mvarLinePoints(i)
    Next i
    
    PointsSpent = intTotal
    
End Function

Public Function PointsAvailable() As Integer
    PointsAvailable = PointsAcquired - PointsSpent
End Function

Public Function cPointsToLevel(intPointsA As Integer) As Integer
' Return the level of specialisation for a set number of points
Dim intLevel As Integer

    ' If the number of points is equal or greater than maximum points (1 as multiplier) then just return the maximum level
    If intPointsA >= MAXPOINTS Then
        cPointsToLevel = MAXLEVEL
        Exit Function
    End If
    ' If the number of points is 1, then return 1
    If intPointsA = 1 Then
        cPointsToLevel = 1
        Exit Function
    End If
    
    Do While intPointsA > 1
        ' See if we have a matching lvl for that number of points
        On Error Resume Next
        intLevel = mvarPointsToLevel(CStr(intPointsA))
        On Error GoTo 0
        ' If we have a match, exit the loop
        If intLevel > 0 Then Exit Do
        ' Otherwise reduce the number of point and repeat
        intPointsA = intPointsA - 1
    Loop
        
    ' Return the level
    cPointsToLevel = intLevel
    
End Function

Public Function cLevelToPoints(intLevelA As Integer) As Integer
' Return the number of points required to reach a certain level of specialisation
Dim intPoints As Integer

    ' Can't have a higher spec than current level
    If intLevelA > mvarLevel Then intLevelA = mvarLevel

    On Error Resume Next
    intPoints = mvarLevelToPoints(CStr(intLevelA))
    On Error GoTo 0
    
    cLevelToPoints = intPoints

End Function

Private Sub Class_Initialize()

    SetCollections
    
    ' Set the default char data
    ResetRealmData
    mvarName = "No Name"
    mvarRealm = "Realm"
    mvarRace = "Race"
    mvarClass = "Class"
  
End Sub

Private Sub Class_Terminate()

    Set mvarNonSpecLines = Nothing
    Set mvarSpecLines = Nothing
    Set mvarLevelToPoints = Nothing
    Set mvarLevelToPointsEarnt1 = Nothing
    Set mvarLevelToPointsEarnt15 = Nothing
    Set mvarLevelToPointsEarnt2 = Nothing
    Set mvarLevelToPointsEarnt22 = Nothing
    Set mvarLevelToPointsEarnt25 = Nothing
    Set mvarPointsToLevel = Nothing
    Set mvarBaseAttributes = Nothing
    Set mvarBuffedAttributes = Nothing
    Set mvarAdditionalAttributes = Nothing
    Set mvarLinePercent = Nothing
    Set mvarLineLevel = Nothing
    Set mvarLinePoints = Nothing
    Set mvarArmorLines = Nothing
    Set mvarOtherLines = Nothing
   
End Sub

Public Function BaseAttribute(intAttributeA As Integer) As Integer
Dim intAttribute As Integer

    On Error Resume Next
    intAttribute = mvarBaseAttributes(AttributeArray(intAttributeA))
    On Error GoTo 0
    
    BaseAttribute = intAttribute
    
End Function

Public Property Let AdditionalAttribute(intAttributeA As Integer, intValue As Integer)

    SetDirty
    ' Remove the item if it already exists
    On Error Resume Next
    mvarAdditionalAttributes.Remove AttributeArray(intAttributeA)
    On Error GoTo 0
    ' Add the new item
    mvarAdditionalAttributes.Add intValue, AttributeArray(intAttributeA)

End Property

Public Property Get AdditionalAttribute(intAttributeA As Integer) As Integer
Dim intAttribute As Integer

    On Error Resume Next
    intAttribute = mvarAdditionalAttributes(AttributeArray(intAttributeA))
    On Error GoTo 0

    AdditionalAttribute = intAttribute

End Property

Public Property Let AdditionalAttributePoints(intAttributeA As Integer, intPointsA As Integer)
Dim intAttribute As Integer

    SetDirty
    ' Convert points to attribute
    If intPointsA <= 10 Then
        intAttribute = intPointsA
    Else
        intPointsA = intPointsA - 10
        If intPointsA <= 10 Then
            intAttribute = 10 + intPointsA / 2
        Else
            intPointsA = intPointsA - 10
            intAttribute = 15 + intPointsA / 3
        End If
    End If
    
    AdditionalAttribute(intAttributeA) = intAttribute

End Property

Public Property Get AdditionalAttributePoints(intAttributeA As Integer) As Integer
Dim intAttribute As Integer
Dim intPoints As Integer

    On Error Resume Next
    intAttribute = mvarAdditionalAttributes(AttributeArray(intAttributeA))
    On Error GoTo 0

    If intAttribute > 10 Then
        If intAttribute > 15 Then
            intAttribute = intAttribute - 15
            intPoints = 20 + (intAttribute * 3)
        Else
            intAttribute = intAttribute - 10
            intPoints = 10 + (intAttribute * 2)
        End If
    Else
        intPoints = intAttribute
    End If
    AdditionalAttributePoints = intPoints

End Property

Public Property Let BuffedAttribute(intAttributeA As Integer, intValue As Integer)

    SetDirty
    ' Remove the item if it already exists
    On Error Resume Next
    mvarBuffedAttributes.Remove AttributeArray(intAttributeA)
    On Error GoTo 0
    ' Add the new item
    mvarBuffedAttributes.Add intValue, AttributeArray(intAttributeA)

End Property

Public Property Get BuffedAttribute(intAttributeA As Integer) As Integer
Dim intAttribute As Integer

    On Error Resume Next
    intAttribute = mvarBuffedAttributes(AttributeArray(intAttributeA))
    On Error GoTo 0

    BuffedAttribute = intAttribute
    
End Property

Public Function TotalAttribute(intAttributeA As Integer) As Integer
Dim intAttribute As Integer

    On Error Resume Next
    intAttribute = mvarBaseAttributes(AttributeArray(intAttributeA))
    intAttribute = intAttribute + EarnedAttribute(intAttributeA)
    intAttribute = intAttribute + AdditionalAttribute(intAttributeA)
    intAttribute = intAttribute + BuffedAttribute(intAttributeA)
    On Error GoTo 0
   
    TotalAttribute = intAttribute
    
End Function

Public Property Get EarnedAttribute(intAttributeA As Integer) As Integer
Dim intAttribute As Integer

    Select Case AttributeArray(intAttributeA)
        Case mvarAttributePrimary
            intAttribute = Int(mvarLevel - 6) + 1
        Case mvarAttributeSecondary
            intAttribute = Int((mvarLevel - 6) / 2) + 1
        Case mvarAttributeTertiary
            intAttribute = Int((mvarLevel - 6) / 3) + 1
        Case Else
            EarnedAttribute = 0
    End Select
    
    If intAttribute < 0 Then intAttribute = 0
    
    EarnedAttribute = intAttribute

End Property

Public Function DirtyBit() As Boolean

End Function

Public Sub SaveToon()
Dim i As Integer

    If mvarDirtyBit Then
        If mvarName = vbNullString Then mvarName = "no name"
        ' Delete existing information
        ProfileDeleteSection mvarName, App.Path & TOONDATA
        
        ' Write the new
        ProfileSaveItem mvarName, "Realm", mvarRealm, App.Path & TOONDATA
        ProfileSaveItem mvarName, "Race", mvarRace, App.Path & TOONDATA
        ProfileSaveItem mvarName, "Class", mvarClass, App.Path & TOONDATA
        ProfileSaveItem mvarName, "SpecByPercent", CStr(mvarSpecByPercent), App.Path & TOONDATA
        ProfileSaveItem mvarName, "Level", CStr(mvarLevel), App.Path & TOONDATA
        On Error Resume Next
        For i = LBound(AttributeArray) To UBound(AttributeArray)
            ProfileSaveItem mvarName, "Additional " & AttributeArray(i), CStr(mvarAdditionalAttributes(AttributeArray(i))), App.Path & TOONDATA
            ProfileSaveItem mvarName, "Buffed " & AttributeArray(i), CStr(mvarBuffedAttributes(AttributeArray(i))), App.Path & TOONDATA
        Next i
        For i = 1 To mvarSpecLines.Count
            ProfileSaveItem mvarName, mvarSpecLines(i).LineName, CStr(mvarLinePoints(mvarSpecLines(i).LineName)), App.Path & TOONDATA
        Next i
        On Error GoTo 0
        mvarDirtyBit = False
     End If

End Sub

Private Sub SetDirty()
    mvarDirtyBit = True
End Sub

Public Sub DeleteToon()

End Sub

Public Sub LoadToon()
Dim strResult As String
Dim i As Integer

    
    strResult = ProfileGetItem(mvarName, "Realm", "unknown", App.Path & TOONDATA)
    If strResult = "unknown" Then
        LoadFormFromChar
        Exit Sub
    End If
    Realm = strResult
    
    strResult = ProfileGetItem(mvarName, "Race", "unknown", App.Path & TOONDATA)
    If strResult <> "unknown" Then Race = strResult
    
    strResult = ProfileGetItem(mvarName, "Class", "unknown", App.Path & TOONDATA)
    If strResult <> "unknown" Then Class = strResult
    
    strResult = ProfileGetItem(mvarName, "SpecByPercent", "unknown", App.Path & TOONDATA)
    If strResult <> "unknown" Then SpecByPercent = CBool(strResult)
    
    strResult = ProfileGetItem(mvarName, "Level", "unknown", App.Path & TOONDATA)
    If strResult <> "unknown" Then Level = Val(strResult)
    
    strResult = vbNullString
    
    For i = LBound(AttributeArray) To UBound(AttributeArray)
        strResult = ProfileGetItem(mvarName, "Additional " & AttributeArray(i), "unknown", App.Path & TOONDATA)
        If strResult <> "unknown" Then AdditionalAttribute(i) = Val(strResult)
        strResult = ProfileGetItem(mvarName, "Buffed " & AttributeArray(i), "unknown", App.Path & TOONDATA)
        If strResult <> "unknown" Then BuffedAttribute(i) = Val(strResult)
    Next i
    
    For i = 1 To mvarSpecLines.Count
        strResult = ProfileGetItem(mvarName, mvarSpecLines(i).LineName, "unknown", App.Path & TOONDATA)
        If strResult <> "unknown" Then LinePoints(i) = Val(strResult)
    Next i
    
    ' Redraw form
    LoadFormFromChar
            
End Sub
