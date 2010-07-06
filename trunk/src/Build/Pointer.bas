Attribute VB_Name = "basPointer"
Option Explicit
Public Const DEBUGMODE = True

Public Const DATAFILE = "/Data/DataSet.dat"
Public Const TOONDATA = "/Data/ToonSet.dat"
Public Const MAXPOINTS = 1274 ' Using 1 as a multiplier
Public Const MAXLEVEL = 50
Public Const MINLEVEL = 5

Public bWaiting As Boolean

Public AttributeArray(1 To 8) As String

' Layout Constants
Public Const TABSTOP = 70 ' pixels from top of form
Public Const DESCRIPTIONTOP = 80 ' pixels from bottom of form

' Indexs into attribute grid
Public Const STRENGTH = 1
Public Const CONSTITUTION = 2
Public Const DEXTERITY = 3
Public Const QUICKNESS = 4
Public Const INTELLIGENCE = 5
Public Const PIETY = 6
Public Const EMPATHY = 7
Public Const CHARISMA = 8
Public Const ATTRIBUTECAPTION = 0
Public Const ATTRIBUTESTART = 1
Public Const ATTRIBUTEADDITIONAL = 2
Public Const ATTRIBUTELEVEL = 3
Public Const ATTRIBUTEBUFF = 4
Public Const ATTRIBUTETOTAL = 5

' Indexs into spec grid
Public Const SPEC1 = 1
Public Const SPEC2 = 2
Public Const SPEC3 = 3
Public Const SPEC4 = 4
Public Const SPEC5 = 5
Public Const SPEC6 = 6
Public Const SPEC7 = 7
Public Const SPEC8 = 8
Public Const SPECCAPTION = 0
Public Const SPECPERCENT = 1
Public Const SPECPOINTS = 2
Public Const SPECLEVEL = 3
Public Const SPECSTYLES = 4

' Indexs into non spec grid
Public Const NONSPECCAPTION = 0
Public Const NONSPECSTYLES = 1

Public Toon As clsChar

Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal strSectionNameA As String, ByVal strKeyNameA As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal strSectionNameA As String, ByVal strKeyNameA As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub DeleteCharData(strNameA As String)

    ' Delete the data section
    ProfileDeleteSection strNameA, App.Path & TOONDATA
    ' Names get rewritten on program close

End Sub

Public Sub SetListByValue(cmb As Control, vData As String)
Dim i As Integer

    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = vData Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i
    
    ' If we get this far the value is not in the list
    ' so add it
    cmb.AddItem vData
    cmb.ListIndex = cmb.ListCount - 1
    

End Sub

Public Sub RemoveListItem(cmb As Control, vData As String, Optional strDefault As String = vbNullString)
Dim i As Integer
Dim CurrentItem As String

    CurrentItem = cmb.Text
    
    For i = cmb.ListCount - 1 To 0 Step -1
        If cmb.List(i) = vData Then
            cmb.RemoveItem i
        End If
    Next i
    
    If CurrentItem = vData Then
        ' Set the index
        If cmb.ListCount > 0 Then
            cmb.ListIndex = 0
        Else
            ' empty list so add a default value
            If strDefault <> vbNullString Then SetListByValue cmb, strDefault
        End If
    End If
    
End Sub

Public Sub LoadFormFromChar()
Dim i As Integer

    With frmPointer
        SetListByValue .cmbName, Toon.CharName
        SetListByValue .cmbRealm, Toon.Realm
        LoadRaces Toon.Realm
        SetListByValue .cmbRace, Toon.Race
        LoadClasses Toon.Race
        SetListByValue .cmbClass, Toon.Class
    End With
' Load specific class data

    SystemWait
'    frmPointer.Refresh
'
'    frmProgress.Progress "Loading " & Toon.Class & " Data"
'    frmProgress.Move frmPointer.Left + (frmPointer.Width - frmProgress.Width) / 2, frmPointer.Top + (frmPointer.Height - frmProgress.Height) / 2
'    frmProgress.Show
'    frmProgress.Refresh
    ' Set the Toon class
'    Toon.Class = cmbClass.List(cmbClass.ListIndex)
    ' Set the attributes in the grid
    SetAttributeGrid
    SetPointsAcquired
    SetLevel
    SetSpecGrid
    SetNonSpecGrid
    UnlockLeveltxt
'
'    Unload frmProgress
'    frmPointer.SetFocus
    SystemContinue
    
'    SetAttributeGrid
    SetAvailableArmor
    SetAvailableNonSpecStyles
    SetAvailableOther
    SetAvailableSpecStyles
    SetSpecPoints
    
    If Toon.SpecByPercent Then
        frmPointer.chkSpecPercent = vbChecked
    Else
        frmPointer.chkSpecPercent = vbUnchecked
    End If
    
End Sub

Public Sub SetPointsAcquired()
    frmPointer.txtPointsAcquired.Text = Toon.PointsAcquired
End Sub

Public Sub SetPointsSpent()
    frmPointer.txtPointsSpent.Text = Toon.PointsSpent
End Sub

Public Sub SetLevel()
    frmPointer.txtLevel.Text = Toon.Level
End Sub


Public Sub LockLeveltxt()

    frmPointer.txtLevel.Locked = True
    frmPointer.txtLevel.BackColor = &HE0F0E0
    
End Sub

Public Sub UnlockLeveltxt()

    frmPointer.txtLevel.Locked = False
    frmPointer.txtLevel.BackColor = vbWhite

End Sub

Public Sub SetAttributeArray()
Dim i As Integer

    For i = 1 To 8
        AttributeArray(i) = ProfileGetItem("Attributes", "Attribute " & i, "unknown", App.Path & DATAFILE)
    Next i

End Sub

Public Sub ProfileSaveItem(strSectionNameA As String, strKeyNameA As String, strValueA As String, strInifileA As String)
   
   Call WritePrivateProfileString(strSectionNameA, strKeyNameA, strValueA, strInifileA)

End Sub

Public Sub SelectText(txtBox As TextBox)

    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)

End Sub

Public Function ProfileGetItem(strSectionNameA As String, strKeyNameA As String, strDefaultValueA As String, strInifileA As String) As String
        
   Dim success As Long
   Dim nSize As Long
   Dim ret As String
  
  ' Pad a string large enough to hold the data.
   ret = Space$(2048)
   nSize = Len(ret)
   success = GetPrivateProfileString(strSectionNameA, strKeyNameA, strDefaultValueA, ret, nSize, strInifileA)
   
   If success Then
      ProfileGetItem = Left$(ret, success)
   End If
   
End Function

Public Sub ProfileDeleteItem(strSectionNameA As String, strKeyNameA As String, strInifileA As String)
   
   Call WritePrivateProfileString(strSectionNameA, strKeyNameA, vbNullString, strInifileA)

End Sub

Public Sub ProfileDeleteSection(strSectionNameA As String, strInifileA As String)
   
   Call WritePrivateProfileString(strSectionNameA, vbNullString, vbNullString, strInifileA)

End Sub

Sub Main()

    SetAttributeArray
    Set Toon = New clsChar
    
    LoadDataSet
    SetAttributeGrid
    SetUpSpecGrid
    LoadNames
    frmPointer.Show
    Toon.LoadToon
    
End Sub

Public Function StripReturns(strTextA As String) As String
    StripReturns = Trim(Replace(strTextA, vbCrLf, " "))
End Function

Public Sub CleanUp()

    Set Toon = Nothing
    
End Sub

Public Sub LoadRealms()
Dim strRealm As String
Dim i As Integer

     ' Get all the realms
    Do While strRealm <> "unknown"
        i = i + 1
        ' Get the first/next realm
        strRealm = ProfileGetItem("Realms", "Realm " & i, "unknown", App.Path & DATAFILE)
        ' Bail if there isn't one
        If strRealm = "unknown" Then Exit Do
        ' Otherwise add the realm to the realm combo
        frmPointer.cmbRealm.AddItem strRealm
    Loop
   
End Sub

Public Sub LoadNames()
' Load available toon names into name list
Dim strResult As String
Dim i As Integer

     ' Get all the realms
    Do While strResult <> "unknown"
        i = i + 1
        ' Get the first/next realm
        strResult = ProfileGetItem("Names", "Name " & i, "unknown", App.Path & TOONDATA)
        ' Bail if there isn't one
        If strResult = "unknown" Then Exit Do
        ' Otherwise add the realm to the realm combo
        frmPointer.cmbName.AddItem strResult
    Loop

    ' Set to no name if nothing loaded
    If frmPointer.cmbName.ListCount = 0 Then frmPointer.cmbName.AddItem "No Name"
    
    frmPointer.cmbName.ListIndex = 0
    Toon.CharName = frmPointer.cmbName.List(0)
    
End Sub

Public Sub SaveNames()
'
Dim i As Integer

    ' Delete the old names section
    ProfileDeleteSection "Names", App.Path & TOONDATA

    For i = 1 To frmPointer.cmbName.ListCount
        ProfileSaveItem "Names", "Name " & i, frmPointer.cmbName.List(i - 1), App.Path & TOONDATA
    Next i

End Sub

Public Sub LoadRaces(strRealmA As String)
Dim strRace  As String
Dim i As Integer

    ' Clear the combo
    frmPointer.cmbRace.Clear
    SetListByValue frmPointer.cmbRace, "Race"
    ' Get all the races for the realm
    Do While strRace <> "unknown"
        i = i + 1
        ' Get the spec lines first/next style
        strRace = ProfileGetItem(strRealmA & " Races", "Race " & i, "unknown", App.Path & DATAFILE)
        ' Bail if there isn't one
        If strRace = "unknown" Then Exit Do
        ' Otherwise add the race to the race combo
        frmPointer.cmbRace.AddItem strRace
    Loop
    
    ' Bubble down to classes
    LoadClasses ("unknown")
    
End Sub

Public Sub LoadClasses(strRaceA As String)
Dim strClass  As String
Dim i As Integer

    ' Clear the combo
    frmPointer.cmbClass.Clear
    SetListByValue frmPointer.cmbClass, "Class"
    ' Get all the races for the realm
    Do While strClass <> "unknown"
        i = i + 1
        ' Get the spec lines first/next style
        strClass = ProfileGetItem(strRaceA, "Class " & i, "unknown", App.Path & DATAFILE)
        ' Bail if there isn't one
        If strClass = "unknown" Then Exit Do
        ' Otherwise add the race to the race combo
        frmPointer.cmbClass.AddItem strClass
    Loop
    
End Sub

Public Sub SetStartingAttributes()
' Load the data for a specific race
Dim i As Integer
Dim strAttribute As String

    With frmPointer.grdAttributes
        For i = LBound(AttributeArray) To UBound(AttributeArray)
            .TextArray((ATTRIBUTESTART * .Cols) + i) = Toon.BaseAttribute(i)
        Next i
    End With
    
    SetTotalAttributes
    
End Sub

Public Sub SetSpecPoints()

    With frmPointer
        .txtPointsSpent = Toon.PointsSpent
        .txtPointsAcquired = Toon.PointsAcquired
        .txtPointsAvailable = Toon.PointsAvailable
    End With
    
End Sub

Private Sub LoadDataSet()
' Load dat from file for interface and calculater

    LoadRealms
    
End Sub

Public Sub SetEarnedAttributes()
' Recalculate earned attributes based on level and class
Dim i As Integer
Dim intPoints As Integer

    With frmPointer.grdAttributes
        For i = 1 To 8
            intPoints = Toon.EarnedAttribute(i)
            If intPoints > 0 Then .TextArray((ATTRIBUTELEVEL * .Cols) + i) = intPoints
        Next i
    End With
    
    SetTotalAttributes
    
End Sub

Public Sub SetTotalAttributes()
' Recalculate total attributes
Dim i As Integer
Dim intPoints As Integer

    With frmPointer.grdAttributes
        For i = LBound(AttributeArray) To UBound(AttributeArray)
            .TextArray((ATTRIBUTETOTAL * .Cols) + i) = Toon.TotalAttribute(i)
        Next i
    End With

End Sub

Private Function AdditionalAttribute(intPointsIn As Integer) As Integer
' Return the amount of additional buff for points spent (max 30 points)
Dim intBuff As Integer

    If intPointsIn <= 10 Then
        intBuff = intPointsIn
    Else
        intPointsIn = intPointsIn - 10
        If intPointsIn <= 10 Then
            intBuff = 10 + intPointsIn / 2
        Else
            intPointsIn = intPointsIn - 10
            intBuff = 15 + intPointsIn / 3
        End If
    End If
    
    AdditionalAttribute = intBuff
    
End Function

Private Sub SetUpSpecGrid()
Dim intSpecLineCount As Integer
Dim i As Integer

    With frmPointer.grdSpec
        ' Reset
        .Clear
        .Cols = 1
        .ColWidth(0) = 100 * Screen.TwipsPerPixelX
        ' Write captions
        .TextArray(SPECPERCENT * .Cols) = "Percent Spec"
        .TextArray(SPECPOINTS * .Cols) = "Points Spent"
        .TextArray(SPECLEVEL * .Cols) = "Skill Level"
        .TextArray(SPECSTYLES * .Cols) = "Styles"
        .Cols = 1
    End With

End Sub

Public Sub ResetSpecGrid(Optional strClassA As String)
Dim intSpecLineCount As Integer
Dim i As Integer

    intSpecLineCount = frmPointer.grdSpec.Cols - 1
    
    SetUpSpecGrid
    If strClassA <> vbNullString Then
        intSpecLineCount = ProfileGetItem(strClassA, "Specialisation Line Count", "0", App.Path & DATAFILE)
    Else
        frmPointer.grdSpec.Cols = intSpecLineCount + 1
    End If
    
    With frmPointer.grdSpec
        .Cols = intSpecLineCount + 1
        ' Write  Spec Lines
        For i = 1 To intSpecLineCount
            .TextArray(SPECCAPTION + i) = ProfileGetItem(strClassA, "Specialisation " & i, "0", App.Path & DATAFILE)
        Next i
        If Toon.SpecByPercent Then
            .Row = SPECPERCENT
            For i = 1 To .Cols - 1
                .Col = i
                .CellBackColor = vbWhite
            Next i
        Else
            .Row = SPECPOINTS
            For i = 1 To .Cols - 1
                .Col = i
                .CellBackColor = vbWhite
            Next i
            .Row = SPECLEVEL
            For i = 1 To .Cols - 1
                .Col = i
                .CellBackColor = vbWhite
            Next i
        End If
    End With
    
End Sub

Public Sub SetAttributeGrid()
Dim i As Integer

    With frmPointer.grdAttributes
        ' Clear the grid
        .Clear
        ' Widen the first column
        .ColWidth(0) = 100 * Screen.TwipsPerPixelX
        ' Set the number of columns to the number of attributes (is 1 based)
        .Cols = UBound(AttributeArray) + 1
        ' Loop through each attribute/column
        For i = LBound(AttributeArray) To UBound(AttributeArray)
            .Col = i
            
            ' Set the back colour of editable cells
            .Row = ATTRIBUTEADDITIONAL
            .CellBackColor = vbWhite
            .Row = ATTRIBUTEBUFF
            .CellBackColor = vbWhite
            
            ' Set the caption
            .TextArray(ATTRIBUTECAPTION * .Cols + i) = AttributeArray(i)
            ' Set the base attributes
            .TextArray(ATTRIBUTESTART * .Cols + i) = Toon.BaseAttribute(i)
            ' Set the additional attributes
            .TextArray(ATTRIBUTEADDITIONAL * .Cols + i) = Toon.AdditionalAttribute(i)
            ' Set the earned attributes
            .TextArray(ATTRIBUTELEVEL * .Cols + i) = Toon.EarnedAttribute(i)
            ' Set the buffed attributes
            .TextArray(ATTRIBUTEBUFF * .Cols + i) = Toon.BuffedAttribute(i)
            ' Set the total attributes
            .TextArray(ATTRIBUTETOTAL * .Cols + i) = Toon.TotalAttribute(i)
        Next i
        
        ' Set the row titles
        .TextArray(ATTRIBUTESTART * .Cols) = "Starting"
        .TextArray(ATTRIBUTEADDITIONAL * .Cols) = "Additional (max 30)"
        .TextArray(ATTRIBUTELEVEL * .Cols) = "Earned (by lvl)"
        .TextArray(ATTRIBUTEBUFF * .Cols) = "Buffs"
        .TextArray(ATTRIBUTETOTAL * .Cols) = "Total"
    End With

End Sub

Public Sub SetSpecGrid()
Dim i As Integer
Dim j As Integer

    With frmPointer.grdSpec
        ' Clear the grid
        .Clear
        ' Widen the first column
        .ColWidth(0) = 100 * Screen.TwipsPerPixelX
        ' Set the number of columns to the number of toon spec lines
        .Cols = Toon.SpecLineCount + 1
        ' Loop through each attribute/column
        For i = 1 To .Cols - 1
            .Col = i
            ' Widen the column
            .ColWidth(i) = 100 * Screen.TwipsPerPixelX
            ' Set the back colour of editable cells
            If Toon.SpecByPercent Then
                ' Highlight percentage row
                .Row = SPECPERCENT
                .CellBackColor = vbWhite
            Else
                ' Highlight Level and points rows
                .Row = SPECLEVEL
                .CellBackColor = vbWhite
                .Row = SPECPOINTS
                .CellBackColor = vbWhite
            End If
            
            ' Set the caption
            .TextArray((SPECCAPTION * .Cols) + i) = Toon.SpecLine(i).LineName
            If Toon.SpecLine(i).Level > 1 Then
                .TextArray((SPECCAPTION * .Cols) + i) = .TextArray((SPECCAPTION * .Cols) + i) & " (lvl " & Toon.SpecLine(i).Level & ")"
            End If
            ' Set percent
            .TextArray((SPECPERCENT * .Cols) + i) = Toon.LinePercent(i)
            ' Set points
            .TextArray((SPECPOINTS * .Cols) + i) = Toon.LinePoints(i)
            ' Set level
            .TextArray((SPECLEVEL * .Cols) + i) = Toon.LineLevel(i)
            ' Write out styles
            ' Ensure we have enough rows
            If .Rows < Toon.SpecLine(i).Count + SPECSTYLES Then .Rows = Toon.SpecLine(i).Count + SPECSTYLES
            For j = 0 To Toon.SpecLine(i).Count - 1
                .TextArray(((SPECSTYLES + j) * .Cols) + i) = Toon.SpecLine(i).Item(j + 1).StyleName
                ' Now set the font colour if we can perform that style based on level
                If Toon.SpecLine(i).Item(j + 1).Level <= Toon.LineLevel(i) Then
                    .Row = SPECSTYLES + j
                    .CellForeColor = vbRed
                End If
            Next j
            
        Next i
        
        ' Write captions
        .TextArray(SPECPERCENT * .Cols) = "Percent Spec"
        .TextArray(SPECPOINTS * .Cols) = "Points Spent"
        .TextArray(SPECLEVEL * .Cols) = "Skill Level"
        .TextArray(SPECSTYLES * .Cols) = "Styles"
    End With

End Sub

Public Sub SetSpecPointsSpend()
Dim i As Integer
    
    With frmPointer.grdSpec
        For i = 1 To .Cols - 1
            ' Set percent
            .TextArray((SPECPERCENT * .Cols) + i) = Toon.LinePercent(i)
            ' Set points
            .TextArray((SPECPOINTS * .Cols) + i) = Toon.LinePoints(i)
            ' Set level
            .TextArray((SPECLEVEL * .Cols) + i) = Toon.LineLevel(i)
        Next i
    End With

End Sub
Public Sub Update()

'    Shell App.Path & "\launchupdater.exe"
    ShellEx App.Path & "\updaterlaunch.exe", , , , , frmPointer.hWnd

End Sub

Public Sub GotoURL(strURLA As String)

    ShellEx strURLA, , , , , frmPointer.hWnd

End Sub

Public Sub RunShellExecute(sTopic As String, sFIle As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()

   
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFIle, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
   
End Sub

Public Sub SetNonSpecGrid()
Dim i As Integer
Dim j As Integer
Dim MyLine As clsLine

    With frmPointer.grdNonSpec
        ' Clear the grid
        .Clear
        ' Set the number of columns to the number of toon spec lines
        .Cols = Toon.NonSpecLineCount
        ' Loop through each attribute/column
        For i = 0 To .Cols - 1
             ' Widen the column
            .ColWidth(i) = 100 * Screen.TwipsPerPixelX
           .Col = i
            Set MyLine = Toon.NonSpecLine(i + 1)
            ' Set the caption
            .TextArray(i) = MyLine.LineName
            If MyLine.Level > 1 Then
                .TextArray(i) = .TextArray(i) & " (lvl " & MyLine.Level & ")"
            End If            ' Write out styles
            ' Ensure we have enough rows
            If .Rows < MyLine.Count + 1 Then .Rows = MyLine.Count + 1
            For j = NONSPECSTYLES To MyLine.Count + NONSPECSTYLES - 1
                .TextArray((j * .Cols) + i) = MyLine(j - NONSPECSTYLES + 1).StyleName
                ' Now set the font colour if we can perform that style based on level
                If MyLine(j - NONSPECSTYLES + 1).Level <= Toon.Level Then
                    .Row = j
                    .CellForeColor = vbRed
                End If
            Next j
            
        Next i
        
        ' Add Armor
        ' Add a column
        .Cols = .Cols + 1
        i = .Cols - 1
        .Col = i
         ' Widen the column
        .ColWidth(i) = 100 * Screen.TwipsPerPixelX
         Set MyLine = Toon.ArmorLine
        ' Set the caption
        .TextArray(i) = "Armor"
        ' Write out styles
        ' Ensure we have enough rows
        If .Rows < MyLine.Count + 1 Then .Rows = MyLine.Count + 1
        For j = NONSPECSTYLES To MyLine.Count + NONSPECSTYLES - 1
            .TextArray((j * .Cols) + i) = MyLine(j - NONSPECSTYLES + 1).StyleName
            ' Now set the font colour if we can perform that style based on level
            If MyLine(j - NONSPECSTYLES + 1).Level <= Toon.Level Then
                .Row = j
                .CellForeColor = vbRed
            End If
        Next j
        
        ' Add Armor
        ' Add a column
        .Cols = .Cols + 1
        i = .Cols - 1
        .Col = i
         ' Widen the column
        .ColWidth(i) = 100 * Screen.TwipsPerPixelX
         Set MyLine = Toon.OtherLine
        ' Set the caption
        .TextArray(i) = "Other"
        ' Write out styles
        ' Ensure we have enough rows
        If .Rows < MyLine.Count + 1 Then .Rows = MyLine.Count + 1
        For j = NONSPECSTYLES To MyLine.Count + NONSPECSTYLES - 1
            .TextArray((j * .Cols) + i) = MyLine(j - NONSPECSTYLES + 1).StyleName
            ' Now set the font colour if we can perform that style based on level
            If MyLine(j - NONSPECSTYLES + 1).Level <= Toon.Level Then
                .Row = j
                .CellForeColor = vbRed
            End If
        Next j
    End With
    
    Set MyLine = Nothing

End Sub

Public Sub SetAvailableSpecStyles(Optional index As Integer = -1)
Dim i As Integer

    If index = -1 Then
        For i = 1 To Toon.SpecLineCount
            SetAvailableSpecStyles (i)
        Next i
    Else
        For i = 0 To Toon.SpecLine(index).Count - 1
            ' Now set the font colour if we can perform that style based on level
            If Toon.SpecLine(index).Item(i + 1).Level <= Toon.LineLevel(index) Then
                frmPointer.grdSpec.Row = SPECSTYLES + i
                frmPointer.grdSpec.CellForeColor = vbRed
            Else
                frmPointer.grdSpec.Row = SPECSTYLES + i
                frmPointer.grdSpec.CellForeColor = vbBlack
            End If
        Next i
    End If
    
End Sub

Public Sub SetAvailableNonSpecStyles(Optional index As Integer = -1)
Dim i As Integer
Dim MyLine As clsLine

    If index = -1 Then
        For i = 0 To Toon.NonSpecLineCount - 1
            SetAvailableNonSpecStyles (i)
        Next i
    Else
        frmPointer.grdNonSpec.Col = index
        Set MyLine = Toon.NonSpecLine(index + 1)
        For i = 0 To MyLine.Count - 1
            ' Now set the font colour if we can perform that style based on level
            If MyLine(i + 1).Level <= Toon.Level Then
                frmPointer.grdNonSpec.Row = 1 + i
                frmPointer.grdNonSpec.CellForeColor = vbRed
            Else
                frmPointer.grdNonSpec.Row = 1 + i
                frmPointer.grdNonSpec.CellForeColor = vbBlack
            End If
        Next i
        Set MyLine = Nothing
    End If

End Sub

Public Sub SetAvailableArmor()
Dim i As Integer
Dim MyLine As clsLine

    frmPointer.grdNonSpec.Col = frmPointer.grdNonSpec.Cols - 2
    Set MyLine = Toon.ArmorLine
    For i = 0 To MyLine.Count - 1
        ' Now set the font colour if we can perform that style based on level
        If MyLine(i + 1).Level <= Toon.Level Then
            frmPointer.grdNonSpec.Row = 1 + i
            frmPointer.grdNonSpec.CellForeColor = vbRed
        Else
            frmPointer.grdNonSpec.Row = 1 + i
            frmPointer.grdNonSpec.CellForeColor = vbBlack
        End If
    Next i
    Set MyLine = Nothing

End Sub

Public Sub SetAvailableOther()
Dim i As Integer
Dim MyLine As clsLine

    frmPointer.grdNonSpec.Col = frmPointer.grdNonSpec.Cols - 1
    Set MyLine = Toon.OtherLine
    For i = 0 To MyLine.Count - 1
        ' Now set the font colour if we can perform that style based on level
        If MyLine(i + 1).Level <= Toon.Level Then
            frmPointer.grdNonSpec.Row = 1 + i
            frmPointer.grdNonSpec.CellForeColor = vbRed
        Else
            frmPointer.grdNonSpec.Row = 1 + i
            frmPointer.grdNonSpec.CellForeColor = vbBlack
        End If
    Next i
    Set MyLine = Nothing

End Sub

Public Sub SystemWait()

    bWaiting = True
    frmPointer.MousePointer = vbHourglass
    
End Sub

Public Sub SystemContinue()

    bWaiting = False
    frmPointer.MousePointer = vbDefault

End Sub
